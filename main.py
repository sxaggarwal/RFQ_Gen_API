import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from threading import Thread
from src.general_class import TableManger
from src.mie_trak import MieTrak
from src.helper import create_dict_from_excel, transfer_file_to_folder, pk_info_dict
import os

class LoadingScreen(tk.Toplevel):
    def __init__(self, master, max_progress):
        super().__init__(master)
        self.title("Generating RFQ")
        self.geometry("300x100")
        self.protocol("WM_DELETE_WINDOW", self.disable_close_button)
        self.progressbar = ttk.Progressbar(self, orient="horizontal", length=200, mode="determinate", maximum=max_progress)
        self.progressbar.pack(pady=10)
    
    def set_progress(self, value):
        self.progressbar["value"] = value
        if value >= self.progressbar["maximum"]:
            self.destroy()
    
    def disable_close_button(self):
        pass

class RfqGen(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RFQGen")
        self.geometry("305x415")
        self.data_base_conn = MieTrak()
        self.customer_names = self.data_base_conn.get_customer_data(names=True)
        self.quote_assembly_table = TableManger("QuoteAssembly")
        self.make_combobox()
    
    def filter_combobox(self, event):
        current_text = self.customer_select_box.get().lower()
        self.customer_select_box['values'] = ()
        filtered_values = [name for name in self.customer_names if name.lower().startswith(current_text)]
        self.filtered_indices = [idx for idx, name in enumerate(self.customer_names) if name in filtered_values]
        self.customer_select_box['values'] = filtered_values
    
    def make_combobox(self):
        tk.Label(self, text="Select Customer: ").grid(row=0, column=0)
        self.customer_select_box = ttk.Combobox(self, values=self.customer_names, state="normal")
        self.customer_select_box.grid(row=1, column=0)
        
        tk.Label(self, text="Selected Customer Info: ").grid(row=2, column=0)
        self.customer_info_text = tk.Text(self, height=4, width=30)
        self.customer_info_text.grid(row=3, column=0)

        self.filtered_indices = []
        self.customer_select_box.bind('<KeyRelease>', self.filter_combobox)
        self.customer_select_box.bind("<<ComboboxSelected>>", self.update_customer_info)

        tk.Label(self, text="Enter RFQ Number: ").grid(row=4, column=0)
        self.rfq_number_text = tk.Entry(self, width=20)
        self.rfq_number_text.grid(row=5, column=0)

        tk.Label(self, text="Parts Requested File:").grid(row=6, column=0)
        self.file_path_PR_entry = tk.Listbox(self, height=2, width=50)
        self.file_path_PR_entry.grid(row=7, column=0)

        browse_button_1 = tk.Button(self, text="Browse Files", command=lambda: self.browse_files_parts_requested("Excel files", self.file_path_PR_entry))
        browse_button_1.grid(row=8, column=0)

        tk.Label(self, text="Part Lists File (PL):").grid(row=9, column=0)
        self.file_path_PL_entry = tk.Listbox(self, height=2, width=50)
        self.file_path_PL_entry.grid(row=10, column=0)

        browse_button_part_list = tk.Button(self, text="Browse Files", command=lambda: self.browse_files_parts_requested("All files", self.file_path_PL_entry))
        browse_button_part_list.grid(row=11, column=0)

        self.itar_restricted_var = tk.BooleanVar()
        self.itar_restricted_checkbox = tk.Checkbutton(self, text="ITAR RESTRICTED", variable=self.itar_restricted_var)
        self.itar_restricted_checkbox.grid(row=12, column=0)

        generate_button = tk.Button(self, text="Generate RFQ", command=self.generate_rfq_with_loading_screen)
        generate_button.grid(row=13, column=0)
    
    def generate_rfq_with_loading_screen(self):
        self.loading_screen = LoadingScreen(self, max_progress=100)
        Thread(target=self.generate_rfq, args=(self.loading_screen,)).start()  # Start RFQ generation in a separate thread
   
    def update_customer_info(self, event=None):
        self.customer_info_text.delete(1.0, tk.END)
        if self.filtered_indices:
            current_index = self.customer_select_box.current()
            selected_customer_index = self.filtered_indices[current_index]
            short_name, email, party_pk = self.data_base_conn.get_customer_data(selected_customer_index=selected_customer_index)
            self.customer_info_text.insert(tk.END, f"Name: {short_name}\nEmail: {email}")
            self.party_pk = party_pk
        else:
            short_name, email, party_pk = self.data_base_conn.get_customer_data(selected_customer_index=self.customer_select_box.current())
            self.customer_info_text.insert(tk.END, f"Name: {short_name}\nEmail: {email}")
            self.party_pk = party_pk
        
    def browse_files_parts_requested(self, filetype: str, list_box):
        """ Browse button for Part requested section, filetype only accepts -> "All files", "Excel files" """
        if filetype == "Excel files":
            param = (filetype, "*.xlsx;*.xls")
        else:
            param = (filetype, "*.*")

        try:
            self.filepaths = [filepath for filepath in filedialog.askopenfilenames(title="Select Files", filetypes=(param,))]

            # entering all file paths in the listbox
            list_box.delete(0, tk.END)
            for path in self.filepaths:
                list_box.insert(0, path)

        except FileNotFoundError as e:
            print(f"Error during file browse: {e}")
            messagebox.showerror("File Browse Error", "An error occurred during file selection. Please try again.")
    
    def generate_rfq(self, loading_screen):
        """ Main function for generating RFQ """
        if self.customer_select_box.get() and self.file_path_PR_entry.get(0):
            party_pk = self.party_pk
            billing_details, state, country = self.data_base_conn.get_address(party_pk)
            customer_rfq_number = self.rfq_number_text.get()
            rfq_pk = self.data_base_conn.insert_into_rfq(party_pk, billing_details, state, country, customer_rfq_number=customer_rfq_number)
            path_dict = {} #dictionary with file path as key and the pk of the document group
            user_selected_file_paths = list(self.file_path_PR_entry.get(0, tk.END) + self.file_path_PL_entry.get(0, tk.END)) #making a list of file paths that user uploaded
            i = 1
            y = 1
            count = 1
            info_dict = create_dict_from_excel(self.file_path_PR_entry.get(0, tk.END)[0])
            my_dict = pk_info_dict(info_dict)
            restricted = False
            quote_pk_dict = {}
            loading_screen.set_progress(10)
            ct=20
            for key, value in info_dict.items():
                if value[13] is None:
                    if self.itar_restricted_var.get(): # checking if the user clicked on Restricted box or not and based on that destination path is decided
                        destination_path = rf'y:\PDM\Restricted\{self.customer_select_box.get()}\{key}'
                        restricted = True
                    else:
                        destination_path = rf'y:\PDM\Non-restricted\{self.customer_select_box.get()}\{key}'
                    
                    for file in user_selected_file_paths:
                        # folder is get or created and file is copied to this folder
                        file_path_to_add_to_rfq = transfer_file_to_folder(destination_path, file)
                        path = file_path_to_add_to_rfq.lower()
                        if "_pl_" in path or "spdl" in path or "psdl" in path or "pl" in os.path.basename(path):
                            path_dict[file_path_to_add_to_rfq] = 26
                        elif "dwg" in path or "drw" in path:
                            path_dict[file_path_to_add_to_rfq] = 27
                        elif "step" in path or "stp" in path:
                            path_dict[file_path_to_add_to_rfq] = 30
                        elif "zsp" in path or "speco" in path:
                            path_dict[file_path_to_add_to_rfq] = 33
                        elif ".cat" in path:
                            path_dict[file_path_to_add_to_rfq] = 16
                        else:
                            path_dict[file_path_to_add_to_rfq] = None
                    
                    for file, pk in path_dict.items():
                        if count==1:
                            if restricted:
                                self.data_base_conn.upload_documents(file, rfq_fk=rfq_pk, document_type_fk=6, secure_document=1, document_group_pk=pk)
                            else:
                                self.data_base_conn.upload_documents(file, rfq_fk=rfq_pk, document_type_fk=6, document_group_pk=pk)
                    count+=1

                    item_pk = self.data_base_conn.get_or_create_item(key, description=value[0], purchase=0, service_item=0, manufactured_item=1)
                    matching_paths = {path:pk for path,pk in path_dict.items() if key in path}
                    for url, pk in matching_paths.items():
                            if restricted:
                                self.data_base_conn.upload_documents(url, item_fk=item_pk, document_type_fk=2, secure_document=1, document_group_pk=pk)
                            else:
                                self.data_base_conn.upload_documents(url, item_fk=item_pk, document_type_fk=2, document_group_pk=pk)
                    
                    quote_pk = self.data_base_conn.create_quote(party_pk, item_pk, 0, key)
                    quote_pk_dict[key] = quote_pk
                    self.data_base_conn.add_operation_to_quote(quote_pk)

                    a = [6,21,22]
                    quote_assembly_fk = []

                    for x in a:
                        quote_assembly_pk = self.quote_assembly_table.get("QuoteAssemblyPK", QuoteFK=quote_pk, SequenceNumber=x)
                        quote_assembly_fk.append(quote_assembly_pk[0][0])
                    
                    if key in my_dict:
                        dict_values = my_dict[key]
                        for j, k, l in zip(dict_values, quote_assembly_fk, a):   # noqa: E741
                            if j is not None:
                                self.data_base_conn.create_bom_quote(quote_pk, j, k, l, y, part_length=value[1], part_width=value[3], thickness=value[2])
                                y+=1
                    
                    if value[12] is None:
                        rfq_line_pk = self.data_base_conn.create_rfq_line_item(item_pk, rfq_pk, i, quote_pk, quantity=value[10])
                        i+=1
                        self.data_base_conn.rfq_line_qty(rfq_line_pk, value[10])
                    else:
                        part_num = value[12]
                        fk = quote_pk_dict.get(part_num)
                        self.data_base_conn.create_assy_quote(quote_pk, fk, qty_req=value[10])
                    
                    if key in info_dict:
                        dict_values = [value[1], value[2], value[3], value[4], value[8], value[9], value[11]]
                        self.data_base_conn.insert_part_details_in_item(item_pk, key, dict_values)
                        pk_value = my_dict[key]
                        for j in pk_value[1:]:
                            if j:
                                self.data_base_conn.insert_part_details_in_item(j, key, dict_values)
                                for url, pk in matching_paths.items():
                                    if restricted:
                                        self.data_base_conn.upload_documents(url, item_fk=j, document_type_fk=2, secure_document=1, document_group_pk=pk, print_with_purchase_order=1)
                                    else:
                                        self.data_base_conn.upload_documents(url, item_fk=j, document_type_fk=2, document_group_pk=pk, print_with_purchase_order=1)
                        if pk_value[0]:
                            self.data_base_conn.insert_part_details_in_item(pk_value[0], key, dict_values, item_type='Material')

                else:
                    part_num = value[12]
                    fk = quote_pk_dict.get(part_num)
                    quote_assembly_pk = self.quote_assembly_table.get("QuoteAssemblyPK", QuoteFK=fk, SequenceNumber=24)
                    item_fk = self.data_base_conn.get_or_create_item(key, item_type_fk=3, description=value[0], calculation_type_fk=12, purchase_account_fk=130, cogs_acc_fk=130, mps_item=0, forecast_on_mrp=0,mps_on_mrp=0,service_item=0,ship_loose=0,bulk_ship=0)
                    self.data_base_conn.create_bom_quote(fk, item_fk, quote_assembly_pk[0][0], 24, y)
                    y+=1
                loading_screen.set_progress(ct)
                if ct<100:
                    ct+=10
            loading_screen.set_progress(100)
            messagebox.showinfo("Success", f"RFQ generated successfully! RFQ Number: {rfq_pk}")
            self.file_path_PL_entry.delete(0, tk.END)
            self.file_path_PR_entry.delete(0, tk.END)
            self.rfq_number_text.delete(0, tk.END)
        else:
            self.loading_screen.destroy()
            messagebox.showerror("ERROR", "Select Customer/ Upload Parts Requested File")
            self.file_path_PR_entry.delete(0, tk.END)
            self.file_path_PL_entry.delete(0, tk.END)
            self.rfq_number_text.delete(0, tk.END)


if __name__ == "__main__":
    r = RfqGen()
    r.mainloop()