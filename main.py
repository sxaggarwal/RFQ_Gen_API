import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import Calendar
from threading import Thread
from src.general_class import TableManger
from src.mie_trak import MieTrak
from src.helper import (
    create_dict_from_excel,
    transfer_file_to_folder,
    pk_info_dict,
    check_and_create_tooling,
)
import os
import datetime
import re
from typing import Dict, Any
from mie_trak_api import bom, item, party, request_for_quote, quote, router
from base_logger import getlogger
from src.excel_parser import create_dict_from_excel_new, generate_item_pks
from pprint import pprint


LOGGER = getlogger("Main")


class LoadingScreen(tk.Toplevel):
    """Class to display a loading screen whine generating the RFQ"""

    def __init__(self, master, max_progress):
        super().__init__(master)
        self.title("Generating RFQ")
        self.geometry("300x100")
        self.protocol("WM_DELETE_WINDOW", self.disable_close_button)
        self.attributes("-topmost", True)  # Ensure loading screen stays on top
        self.grab_set()
        self.progressbar = ttk.Progressbar(
            self,
            orient="horizontal",
            length=200,
            mode="determinate",
            maximum=max_progress,
        )
        self.progressbar.pack(pady=10)

    def set_progress(self, value):
        self.progressbar["value"] = value
        if value >= self.progressbar["maximum"]:
            self.destroy()

    def disable_close_button(self):
        """The RFQ gen process is not affected if user by mistake clicks on close button in the Loading screen"""
        pass


class AddBuyerScreen(tk.Toplevel):
    """Display window when user tries to add a buyer"""

    def __init__(self, master, party_pk, party_name):
        super().__init__(master)
        self.title(f"Add Buyer for Customer: {party_name}, PartyPK: {party_pk}")
        self.geometry("500x300")
        self.attributes("-topmost", True)
        self.grab_set()
        tk.Label(self, text="Name: ").grid(row=0, column=0)
        self.buyer_name_box = tk.Entry(self, width=20)
        self.buyer_name_box.grid(row=0, column=1)
        tk.Label(self, text="Short Name: ").grid(row=1, column=0)
        self.short_name_box = tk.Entry(self, width=20)
        self.short_name_box.grid(row=1, column=1)
        tk.Label(self, text="Email: ").grid(row=2, column=0)
        self.email_id_box = tk.Entry(self, width=20)
        self.email_id_box.grid(row=2, column=1)
        tk.Label(self, text="Phone Number: ").grid(row=3, column=0)
        self.phone_number_box = tk.Entry(self, width=20)
        self.phone_number_box.grid(row=3, column=1)
        tk.Label(self, text="Title: ").grid(row=4, column=0)
        self.title_box = tk.Entry(self, width=20)
        self.title_box.grid(row=4, column=1)
        save_button = tk.Button(
            self, text="Save", command=lambda: self.save_buyer_info(party_pk)
        )
        save_button.grid(row=5, column=1)

    def save_buyer_info(self, party_pk):
        """Inserts the buyers data in the database"""
        if self.buyer_name_box.get():
            buyer_info_dict = {
                "Name": self.buyer_name_box.get(),
                "Email": self.email_id_box.get(),
                "Phone": self.phone_number_box.get(),
                "ShortName": self.short_name_box.get(),
                "Title": self.title_box.get(),
                "Buyer": 1,
                "HardwareCertificationFK": 1,
                "MaterialCertificationFK": 1,
                "OutsideProcessingCertificationFK": 1,
                "QualityLevelFK": 2,
                "KeepDocumentOnFile": 1,
                "FirstArticleFK": 1,
            }
            buyer_pk = MieTrak().create_buyer(buyer_info_dict, party_pk)
            messagebox.showinfo(
                "Success", f"Buyer created successfully! BuyerPK: {buyer_pk}"
            )
            self.master.update_buyer_combobox()  # TODO: this should be a callback
            self.destroy()
        else:
            messagebox.showerror("ERROR", "Please Enter Name")


class RfqGen(tk.Tk):
    """Main class with main window and generate rfq function"""

    def __init__(self):
        super().__init__()
        self.title("RFQGen")
        self.geometry("950x500")
        self.data_base_conn = MieTrak()
        self.customer_data = party.get_all_party_data()
        self.filtered_dict = None  # we append search results to this.
        self.quote_assembly_table = TableManger("QuoteAssembly")
        self.make_combobox()

    def filter_combobox(self, event):
        """Filter for selecting customers, type and search"""
        current_text = self.customer_select_box.get().lower()
        self.customer_select_box["values"] = ()
        self.filtered_dict = {
            pk: name
            for pk, name in self.customer_data.items()
            if name.lower().startswith(current_text)
        }
        self.customer_select_box["values"] = list(self.filtered_dict.values())

    def filter_buyer_box(self, event):
        """filtering buyer, type and search"""
        current_text = self.buyer_select_box.get().lower()
        self.buyer_select_box["values"] = ()
        filtered_values = [
            name
            for name in list(self.buyer_dict.keys())
            if name.lower().startswith(current_text)
        ]
        self.buyer_select_box["values"] = filtered_values

    def make_combobox(self):
        """Main window GUI"""

        # Customer select combobox
        tk.Label(self, text="Select Customer: ").grid(row=0, column=1)
        self.customer_select_box = ttk.Combobox(
            self, values=list(self.customer_data.values()), state="normal"
        )
        self.customer_select_box.grid(row=1, column=1)

        tk.Label(self, text="Selected Customer/ Buyer Info: ").grid(row=5, column=1)
        self.customer_info_text = tk.Text(self, height=4, width=30)
        self.customer_info_text.grid(row=6, column=1)

        # Applying filter
        self.filtered_indices = []
        self.customer_select_box.bind("<KeyRelease>", self.filter_combobox)

        # Bind the combobox selection event to update customer information
        self.customer_select_box.bind("<<ComboboxSelected>>", self.update_customer_info)

        # Buyer Selection box
        tk.Label(self, text="Select Buyer: ").grid(row=2, column=1)
        self.buyer_select_box = ttk.Combobox(self, state="normal")
        self.buyer_select_box.grid(row=3, column=1)

        self.buyer_select_box.bind("<KeyRelease>", self.filter_buyer_box)

        add_buyer_button = tk.Button(
            self, text="ADD Buyer", command=self.open_add_buyer_screen
        )
        add_buyer_button.grid(row=4, column=1)

        self.buyer_select_box.bind("<<ComboboxSelected>>", self.update_buyer_info)

        tk.Label(self, text="Enter Customer RFQ Number: ").grid(row=7, column=1)
        self.rfq_number_text = tk.Entry(self, width=50)
        self.rfq_number_text.grid(row=8, column=1)

        # Entrybox for the requested parts in an Excel file. (upload for Excel file)
        tk.Label(self, text="Parts Requested File:").grid(row=9, column=1)
        self.file_path_PR_entry = tk.Listbox(self, height=2, width=50)
        self.file_path_PR_entry.grid(row=10, column=1)

        browse_button_1 = tk.Button(
            self,
            text="Browse Files",
            command=lambda: self.browse_files_parts_requested(
                "Excel files", self.file_path_PR_entry
            ),
        )
        browse_button_1.grid(row=11, column=1)

        # Selection/ Upload for PartList
        tk.Label(self, text="Part Lists File (PL):").grid(row=9, column=2)
        self.file_path_PL_entry = tk.Listbox(self, height=2, width=50)
        self.file_path_PL_entry.grid(row=10, column=2)

        browse_button_part_list = tk.Button(
            self,
            text="Browse Files",
            command=lambda: self.browse_files_parts_requested(
                "All files", self.file_path_PL_entry
            ),
        )
        browse_button_part_list.grid(row=11, column=2)

        tk.Label(self, text="Estimating Documents:").grid(row=9, column=0)
        self.file_path_estimating_entry = tk.Listbox(self, height=2, width=50)
        self.file_path_estimating_entry.grid(row=10, column=0)

        browse_button_estimating = tk.Button(
            self,
            text="Browse Files",
            command=lambda: self.browse_files_parts_requested(
                "All files", self.file_path_estimating_entry
            ),
        )
        browse_button_estimating.grid(row=11, column=0)

        # Checkbox for ITAR RESTRICTED
        self.itar_restricted_var = tk.BooleanVar()
        self.itar_restricted_checkbox = tk.Checkbutton(
            self, text="ITAR RESTRICTED", variable=self.itar_restricted_var
        )
        self.itar_restricted_checkbox.grid(row=13, column=2)

        # main button for generating RFQ
        generate_button = tk.Button(
            self, text="Generate RFQ", command=self.generate_rfq_with_loading_screen
        )
        generate_button.grid(row=17, column=0)

        # Add or Update Item button
        add_item_button = tk.Button(self, text="ADD/Update Item", command=self.add_item)
        add_item_button.grid(row=17, column=2)

        # Calendar widgets for selecting Inquiry and Due dates
        tk.Label(self, text="Enter Inquiry Date (MM/DD/YYYY): ").grid(row=12, column=0)
        self.inquiry_date_box = tk.Entry(self, width=20)
        self.inquiry_date_box.grid(row=13, column=0)
        cal_button = tk.Button(self, text="Cal", command=self.open_calendar)
        cal_button.grid(row=14, column=0)

        tk.Label(self, text="Enter Due Date (MM/DD/YYYY): ").grid(row=12, column=1)
        self.due_date_box = tk.Entry(self, width=20)
        self.due_date_box.grid(row=13, column=1)
        cal_due_button = tk.Button(self, text="Cal", command=self.open_due_calendar)
        cal_due_button.grid(row=14, column=1)

        # Entry box for RFQ Number that needs to be updated
        tk.Label(self, text="Enter the RFQ number to be updated: ").grid(
            row=15, column=1
        )
        self.update_rfq_number_text = tk.Entry(self, width=20)
        self.update_rfq_number_text.grid(row=16, column=1)

        update_rfq_button = tk.Button(self, text="Update RFQ", command=self.update_rfq)
        update_rfq_button.grid(row=17, column=1)

    def open_calendar(self):
        """Opens the Calendar and selects the date on double click"""
        top = tk.Toplevel(self)
        top.grab_set()
        self.inq_cal = Calendar(top, selectmode="day", date_pattern="mm/dd/y")
        self.inq_cal.pack(padx=20, pady=20)

        top.bind("<Double-1>", self.get_selected_inquiry_date)

        btn = tk.Button(
            top, text="Get Selected Date", command=self.get_selected_inquiry_date
        )
        btn.pack(pady=10)

    def get_selected_inquiry_date(self, event=None):
        """gets the selected inquiry date"""
        selected_date = self.inq_cal.get_date()
        self.inquiry_date_box.delete(0, tk.END)
        self.inquiry_date_box.insert(tk.END, selected_date)
        self.inq_cal.master.destroy()

    def open_due_calendar(self):
        """Due date calendar widget"""
        top = tk.Toplevel(self)
        top.grab_set()
        self.due_cal = Calendar(top, selectmode="day", date_pattern="mm/dd/y")
        self.due_cal.pack(padx=20, pady=20)

        top.bind("<Double-1>", self.get_selected_due_date)

        btn = tk.Button(
            top, text="Get Selected Date", command=self.get_selected_due_date
        )
        btn.pack(pady=10)

    def get_selected_due_date(self, event=None):
        """gets the selected due date"""
        selected_date = self.due_cal.get_date()
        self.due_date_box.delete(0, tk.END)
        self.due_date_box.insert(tk.END, selected_date)
        self.due_cal.master.destroy()

    def generate_rfq_with_loading_screen(self):
        """Applying thread so that the screen doesn't freeze and the generate RFQ function is run on background"""
        # TODO: loading screen needs to be centered
        self.loading_screen = LoadingScreen(self, max_progress=100)
        Thread(
            target=self.generate_rfq, args=(self.loading_screen,)
        ).start()  # Start RFQ generation in a separate thread

    def open_add_buyer_screen(self):
        """Opens the Add buyer window when the Add button is clicked"""
        if self.customer_select_box.get():
            party_pk = self.party_pk
            name = self.customer_select_box.get()
            AddBuyerScreen(self, party_pk, name)
        else:
            messagebox.showerror("ERROR", "First Select Customer")

    def update_customer_info(self, event=None):
        """Update customer information label when a customer is selected."""
        self.customer_info_text.delete(1.0, tk.END)
        self.buyer_select_box.set("")

        current_index = self.customer_select_box.current()
        customer_data_dict = (
            self.filtered_dict if self.filtered_dict else self.customer_data
        )

        user_selected_party_pk = list(customer_data_dict.keys())[current_index]

        short_name, email = party.get_party_shortname_email(user_selected_party_pk)

        self.customer_info_text.insert(tk.END, f"Name: {short_name}\nEmail: {email}")
        self.party_pk = user_selected_party_pk

        self.update_buyer_combobox()

    def update_buyer_info(self, event=None):
        """Update customer information label when a Buyer is selected"""
        self.customer_info_text.delete(1.0, tk.END)
        user_selected_idx = self.buyer_select_box.current()
        buyer_fk = list(self.buyer_dict.keys())[user_selected_idx]
        # short_name, email = self.data_base_conn.get_buyer_info(buyer_fk)
        short_name, email = party.get_party_shortname_email(buyer_fk)
        self.customer_info_text.insert(tk.END, f"Name: {short_name}\nEmail: {email}")

    def browse_files_parts_requested(self, filetype: str, list_box):
        """Browse button for Part requested section, filetype only accepts -> "All files", "Excel files" """
        if filetype == "Excel files":
            param = (filetype, "*.xlsx;*.xls")
        else:
            param = (filetype, "*.*")

        try:
            self.filepaths = [
                filepath
                for filepath in filedialog.askopenfilenames(
                    title="Select Files", filetypes=(param,)
                )
            ]

            # entering all file paths in the listbox
            list_box.delete(0, tk.END)
            for path in self.filepaths:
                if "PDM" in path or "Estimating" in path:
                    messagebox.showerror(
                        "Error", "Dude! Files from PDM and Estimating can't be uploaded"
                    )
                    list_box.delete(0, tk.END)
                    return
                else:
                    list_box.insert(0, path)

        except FileNotFoundError as e:
            print(f"Error during file browse: {e}")
            messagebox.showerror(
                "File Browse Error",
                "An error occurred during file selection. Please try again.",
            )

    def update_buyer_combobox(self, event=None):
        """Updates the buyer combobox when a customer is selected"""
        try:
            self.buyer_dict = party.get_all_buyers_for_party(self.party_pk)
            self.buyer_select_box["values"] = list(
                self.buyer_dict.values()
            )  # values are sorted from the database
        except RuntimeError as e:
            messagebox.showerror(title="Error in database", message=f"{e}")
            self.buyer_dict = None

    def reset_gui(self):
        self.customer_select_box.set("")
        self.buyer_select_box.set("")
        self.customer_info_text.delete(1.0, tk.END)
        self.file_path_PR_entry.delete(0, tk.END)
        self.file_path_PL_entry.delete(0, tk.END)
        self.file_path_estimating_entry.delete(0, tk.END)
        self.rfq_number_text.delete(0, tk.END)
        self.inquiry_date_box.delete(0, tk.END)
        self.due_date_box.delete(0, tk.END)

    # -------------------------------------------------------------------------------------------------------------

    def generate_rfq(self, loading_screen, update_rfq_pk=None):
        """Main function for generating RFQ, adding line items and creating a quote"""

        # TODO: self.cusotmer_select_box.get() should be partypk instead.
        if (
            not self.customer_select_box.get() or not self.file_path_PR_entry.get(0)
        ):  # checking if user uploaded the part request excel file and selected the customer or not
            self.loading_screen.destroy()
            messagebox.showerror(
                "ERROR", "Select Customer/ Upload Parts Requested File"
            )
            self.reset_gui()
            return

        LOGGER.debug("Extracting excel...")
        info_dict: Dict[str, Dict[str, Any]] = create_dict_from_excel_new(
            self.file_path_PR_entry.get(0, tk.END)[0]
        )
        LOGGER.debug("Excel file values:")
        pprint(info_dict)

        if not info_dict:
            self.loading_screen.destroy()
            messagebox.showerror("ERROR", "Edit Excel File and try Again")
            self.reset_gui()
            return

        customer_rfq_number = self.rfq_number_text.get()  # user input

        # Getting current date, inquiry date and due date
        current_date = datetime.date.today()
        current_date_formatted = (
            f"{current_date.strftime('%m-%d-%Y')} 12:00:00 AM" if current_date else None
        )

        inquiry_date = self.inquiry_date_box.get()
        inq_date = f"{inquiry_date} 12:00:00 AM" if inquiry_date else None

        due_date = self.due_date_box.get()
        due_date_formated = f"{due_date} 12:00:00 AM" if due_date else None

        # we are doing this twice, once when the user makes a selection and second time when presses generate
        buyer_selection_idx = self.buyer_select_box.current()
        buyer_fk = (
            list(self.buyer_dict.keys())[buyer_selection_idx]
            if buyer_selection_idx and self.buyer_dict
            else None
        )

        party_pk = self.party_pk  # getting the pk of the selected customer

        address_dict = party.get_party_address(self.party_pk)
        if update_rfq_pk:
            rfq_pk = update_rfq_pk
        else:  # for new rfqs
            rfq_pk = request_for_quote.insert_into_rfq(
                party_pk,
                address_dict,
                customer_rfq_number=customer_rfq_number,
                buyer_fk=buyer_fk,
                inquiry_date=inq_date,
                due_date=due_date_formated,
                create_date=current_date_formatted,
            )  # creating the rfq with selected customer details

        if not rfq_pk:
            messagebox.showerror(
                title="RFQ error",
                message="RFQ might not be generated, database did not return a value for the insertion. Check last RFQ in MT and regenerate.",
            )
            return None

        LOGGER.debug(f"Created RFQ with pk: {rfq_pk}.")

        # dictionary with file path as key and the pk of the document group
        path_dict = {}
        estimation_path_dict = {}
        user_selected_file_paths = list(self.file_path_PL_entry.get(0, tk.END))

        # TODO: file path pr entry to estimation folder docs
        estimation_folder_docs = list(
            self.file_path_PR_entry.get(0, tk.END)
            + self.file_path_estimating_entry.get(0, tk.END)
        )
        order_by_counter = 1
        count = 1

        LOGGER.debug("Generating items for Mat, HT and OP for...")
        part_mat_ht_op_dict = generate_item_pks(info_dict)
        pprint(part_mat_ht_op_dict)

        item_pk_dict = {}  # {"PartNumber": ItemPK}
        restricted = False
        quote_pk_dict = {}
        loading_screen.set_progress(10)
        ct = 20
        LOGGER.debug("Starting loop to insert items...")
        for new_key, value in info_dict.items():
            key = (
                new_key.split("_____")[0] if self.ends_with_suffix(new_key) else new_key
            )
            LOGGER.debug(f"Current key: {key}")

            hardware_or_supplies = value.get("hardware_or_supplies", None)
            LOGGER.debug(hardware_or_supplies)
            if (
                not hardware_or_supplies
                or hardware_or_supplies == "Tooling - Manufactured"
            ):  # if main part of tooling.
                LOGGER.debug("Hardware/supplies is not None or tooling.")
                if self.itar_restricted_var.get():  # checking if the user clicked on Restricted box or not and based on that destination path is decided
                    # pass
                    destination_path = (
                        rf"y:\PDM\Restricted\{self.customer_select_box.get()}\{key}"
                    )
                    estimation_destinatoin_path = rf"y:\Estimating\Restricted\{self.customer_select_box.get()}\{self.rfq_number_text.get()}"
                    restricted = True
                else:
                    destination_path = (
                        rf"y:\PDM\Non-restricted\{self.customer_select_box.get()}\{key}"
                    )
                    estimation_destinatoin_path = rf"y:\Estimating\Non-restricted\{self.customer_select_box.get()}\{self.rfq_number_text.get()}"

                for file in user_selected_file_paths:
                    # folder is get or created and file is copied to this folder

                    # TODO: here instead of user_selected_file_path it should be estimating
                    file_path_to_add_to_rfq = transfer_file_to_folder(
                        destination_path, file
                    )
                    path = file_path_to_add_to_rfq.lower()
                    if (
                        "_pl_" in path
                        or "spdl" in path
                        or "psdl" in path
                        or "pl" in os.path.basename(path)
                    ):
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

                for file_p in estimation_folder_docs:
                    # TODO: Create different dict for estimation folder docs.
                    file_path_to_add_to_rfq = transfer_file_to_folder(
                        estimation_destinatoin_path, file_p
                    )
                    path = file_path_to_add_to_rfq.lower()
                    if (
                        "_pl_" in path
                        or "spdl" in path
                        or "psdl" in path
                        or "pl" in os.path.basename(path)
                    ):
                        estimation_path_dict[file_path_to_add_to_rfq] = 26
                    elif "dwg" in path or "drw" in path:
                        estimation_path_dict[file_path_to_add_to_rfq] = 27
                    elif "step" in path or "stp" in path:
                        estimation_path_dict[file_path_to_add_to_rfq] = 30
                    elif "zsp" in path or "speco" in path:
                        estimation_path_dict[file_path_to_add_to_rfq] = 33
                    elif ".cat" in path:
                        estimation_path_dict[file_path_to_add_to_rfq] = 16
                    else:
                        estimation_path_dict[file_path_to_add_to_rfq] = None

                # Uploading documents to the RFQ with a counter so that the same document is not uploaded more than once
                for file, pk in estimation_path_dict.items():
                    # TODO; Use the estiamtion folder dict to upload here
                    if count == 1:
                        if restricted:
                            self.data_base_conn.upload_documents(
                                file,
                                rfq_fk=rfq_pk,
                                document_type_fk=6,
                                secure_document=1,
                                document_group_pk=pk,
                            )
                        else:
                            self.data_base_conn.upload_documents(
                                file,
                                rfq_fk=rfq_pk,
                                document_type_fk=6,
                                document_group_pk=pk,
                            )
                count += 1
                # searching for the part on MIE Trak and returns the PK, if the part doesn't exist then it creates an item and returns the pk
                item_dict = {
                    "PartNumber": key,
                    "Description": value.get("description", ""),
                    "Purchase": 0,
                    "ServiceItem": 0,
                    "ManufacturedItem": 1,
                    "ItemTypeFK": 7
                    if hardware_or_supplies == "Tooling - Manufactured"
                    else None,
                }
                item_pk = item.get_or_create_item(**item_dict)
                item_pk_dict[key] = item_pk

                # uploading the documents of the item or part
                matching_paths = {
                    path: pk for path, pk in path_dict.items() if key in path
                }
                for url, pk in matching_paths.items():
                    if restricted:
                        self.data_base_conn.upload_documents(
                            url,
                            item_fk=item_pk,
                            document_type_fk=2,
                            secure_document=1,
                            document_group_pk=pk,
                        )
                    else:
                        self.data_base_conn.upload_documents(
                            url,
                            item_fk=item_pk,
                            document_type_fk=2,
                            document_group_pk=pk,
                        )

                # creating a quote for the Part and getting QuotePk
                quote_pk = quote.create_quote_new(party_pk, item_pk, 0, key)
                quote_pk_dict[key] = quote_pk
                LOGGER.debug("quote_pk_dict updated with new quote number")
                quote.copy_operations_to_quote(quote_pk)

                pprint(quote_pk_dict)

                # Sequence number in Operations for IssueMat, HT, FIN resp
                seq_nums = [6, 21, 22]

                # list of Quote Assembly pk in order MAT, HT, FIN
                quote_assembly_fks = [
                    quote.get_quote_assembly_pk(
                        **{"QuoteFK": quote_pk, "SequenceNumber": x}
                    )
                    for x in seq_nums
                ]

                # creating a Bill of Material for a quote
                mat_ht_fin_pks: tuple = part_mat_ht_op_dict[key]

                LOGGER.debug("Attaching BOM to Item...")
                for pk, quote_ass_fk, num in zip(
                    mat_ht_fin_pks, quote_assembly_fks, seq_nums
                ):
                    LOGGER.debug(f"{pk}, {quote_ass_fk}, {num}")
                    if pk is not None:
                        bom.create_bom_quote(
                            quote_pk,
                            pk,
                            quote_ass_fk,
                            num,
                            order_by_counter,
                            PartLength=value.get("length", ""),
                            PartWidth=value.get("width", ""),
                            Thickness=value.get("thickness", ""),
                        )
                        order_by_counter += 1

                if mat_ht_fin_pks[2]:  # if OP finish is not none
                    op_finish_pk = mat_ht_fin_pks[2]
                    op_part_number = f"{key} - OP Finish"
                    finish_description = value.get("finish_code", "")
                    self.create_finish_router(
                        finish_description, op_finish_pk, op_part_number
                    )

                # Inserting dimensional and other values to the item table for a part and attaching Document to OP, HT, FIN
                # if key in info_dict:
                self.data_base_conn.insert_part_details_in_item(item_pk, key, value)
                pk_value = part_mat_ht_op_dict[key]
                for pk in pk_value[1:]:
                    if pk:
                        self.data_base_conn.insert_part_details_in_item(pk, key, value)
                if pk_value[0]:
                    self.data_base_conn.insert_part_details_in_item(
                        pk_value[0], key, value, item_type="Material"
                    )

            else:
                # if hardware or tooling then adding it to the BOM of its Assembly part accordingly
                part_num = value.get("assy_for", "")

                # TODO: throw error if parent now found??

                fk = quote_pk_dict.get(part_num)
                if value.get("hardware_or_supplies", "") == "Hardware":
                    quote_assembly_pk = quote.get_quote_assembly_pk(
                        **{"QuoteFK": fk, "SequenceNumber": 24}
                    )
                    item_fk = check_and_create_tooling(value.get("description", ""))
                    bom.create_bom_quote(
                        fk,
                        item_fk,
                        quote_assembly_pk,
                        24,
                        order_by_counter,
                        QuantityRequired=value.get("quantity_required", 1.00),
                    )
                    order_by_counter += 1
                elif value.get("hardware_or_supplies", "") == "Tooling":
                    quote_assembly_pk = self.quote_assembly_table.get(
                        "QuoteAssemblyPK", QuoteFK=fk, SequenceNumber=8
                    )
                    item_fk = item.get_or_create_item(
                        **{
                            "PartNumber": key,
                            "Description": value.get("description"),
                            "ItemTypeFK": 7,
                            "MpsItem": 0,
                            "Purchase": 0,
                            "ForecastOnMRP": 0,
                            "MpsOnMRP": 0,
                            "ServiceItem": 0,
                            "ShipLoose": 0,
                            "BulkShip": 0,
                            "CanNotCreateWorkOrder": 1,
                            "CanNotInvoice": 1,
                            "ManufacturedItem": 1,
                        }
                    )
                    bom.create_bom_quote(
                        fk,
                        item_fk,
                        quote_assembly_pk,
                        8,
                        order_by_counter,
                        QuantityRequired=value.get("quantity_required", 1.00),
                    )
                    order_by_counter += 1
            loading_screen.set_progress(ct)
            if ct < 90:
                ct += 10

        # TODO: throw error if no RFQ PK?
        self.create_rfq(
            quote_pk_dict, item_pk_dict, rfq_pk, info_dict
        )  # checking if the Assy or Detail and creating the line item and adding quotes of assembly to the BOM of Assy Line Quotes

        for value in quote_pk_dict.values():
            quote.create_quote_assembly_formula_variable(value)

        loading_screen.set_progress(100)
        messagebox.showinfo(
            "Success", f"RFQ generated successfully! RFQ Number: {rfq_pk}"
        )

        self.reset_gui()

    # -------------------------------------------------------------------------------------------------------------

    # Update: May10
    def create_finish_router(self, finish_description, item_fin_pk, part_num):
        "Adds a router for every finish"
        finish_code = finish_description.split("\n")
        finish_pks = []

        if finish_code:
            for code in finish_code:
                finish_codes_pk = item.get_or_create_item(
                    **{
                        "PartNumber": code[:100],
                        "Description": code[
                            :490
                        ],  # TODO: Fix this as its crossing the limit, add this to the comments.
                        "Inventoriable": 0,
                        "ItemTypeFK": 5,
                        "CertReqdBySupplier": 1,
                        "CanNotCreateWorkOrder": 1,
                        "CanNotInvoice": 1,
                        "PurchaseAccountFK": 125,
                        "CogsAccFk": 125,
                        "CalculationTypeFK": 17,
                        "Comment": code,
                    }
                )
                finish_pks.append(finish_codes_pk)

        router_pk = router.create_router(item_fin_pk, part_num)
        LOGGER.debug(f"Created Router PK: {router_pk}")

        for idx, pk in enumerate(finish_pks, start=1):
            router.create_router_work_center(pk, router_pk, idx)

    def ends_with_suffix(self, s):
        return re.search(r"_____\d+$", s) is not None

    def create_rfq(
        self,
        quote_pk_dict,
        item_pk_dict,
        rfq_pk,
        info_dict: Dict[str, Dict[str, Any]],
        parent_key=None,
        parent_quote_fk=None,
        i=1,
    ):
        """checks if its Assy or Detail and accordingly creates the line item and adds quotes of assembly to the BOM of Assy Line Quotes"""
        main_part_number = None
        main_quote_pk = None

        parent_quote_assembly_pk_dict = {}
        for new_key, value in info_dict.items():
            if self.ends_with_suffix(new_key) is None:
                key = new_key
            else:
                key = new_key.split("_____")[0]

            part_number = key
            assy_for = value.get("assy_for", None)
            quote_pk = quote_pk_dict.get(part_number)
            item_pk = item_pk_dict.get(part_number)

            if not assy_for:
                rfq_line_pk = request_for_quote.create_rfq_line_item_with_qty(
                    item_pk,
                    rfq_pk,
                    i,
                    quote_pk,
                    quantity=value.get("quantity_required"),
                )
                LOGGER.debug(f"{rfq_line_pk}")
                i += 1
                main_quote_pk = quote_pk
                main_part_number = part_number

            elif assy_for and not value.get("hardware_or_supplies"):
                LOGGER.debug("executing hardware or supplies...")

                if not main_part_number or not main_quote_pk:
                    raise ValueError("Data from excel sheet is not proper bruh.")

                if assy_for == main_part_number:
                    quote_fk = main_quote_pk
                    parent_quote_assembly_pk = quote.create_assy_quote(
                        quote_pk, quote_fk, value.get("quantity_required", "")
                    )
                    parent_quote_assembly_pk_dict[part_number] = (
                        parent_quote_assembly_pk
                    )

                else:
                    LOGGER.debug("executing else in create RFQ.")
                    parent_quote_fk = quote_pk_dict[assy_for]
                    if assy_for not in parent_quote_assembly_pk_dict:
                        raise KeyError(
                            f"Key '{assy_for}' not found in parent_quote_assembly_pk_dict"
                        )
                    parent_quote_assembly_pk_new = parent_quote_assembly_pk_dict[
                        assy_for
                    ]
                    parent_quote_assembly_pk = quote.create_assy_quote(
                        quote_pk,
                        main_quote_pk,
                        value.get("quantity_required", 1),
                        parent_quote_fk=parent_quote_fk,
                        parent_quote_asembly=parent_quote_assembly_pk_new,
                    )
                    parent_quote_assembly_pk_dict[part_number] = (
                        parent_quote_assembly_pk
                    )

    def add_item(self):
        """Adds/Update Items"""
        if self.file_path_PR_entry.get(0):
            if self.customer_select_box.get():
                party_pk = self.party_pk
            else:
                party_pk = None
            user_selected_file_paths = list(
                self.file_path_PR_entry.get(0, tk.END)
                + self.file_path_PL_entry.get(0, tk.END)
            )
            info_dict = create_dict_from_excel(
                self.file_path_PR_entry.get(0, tk.END)[0], rfq_generate=False
            )
            path_dict = {}
            restricted = False
            # for key, value in info_dict.items():
            for new_key, value in info_dict.items():
                if self.ends_with_suffix(new_key) is None:
                    key = new_key
                else:
                    key = new_key.split("_____")[0]
                if self.customer_select_box.get():
                    if self.itar_restricted_var.get():  # checking if the user clicked on Restricted box or not and based on that destination path is decided
                        destination_path = (
                            rf"y:\PDM\Restricted\{self.customer_select_box.get()}\{key}"
                        )
                        restricted = True
                    else:
                        destination_path = rf"y:\PDM\Non-restricted\{self.customer_select_box.get()}\{key}"

                    for file in user_selected_file_paths:
                        # folder is get or created and file is copied to this folder
                        file_path_to_add_to_rfq = transfer_file_to_folder(
                            destination_path, file
                        )
                        path = file_path_to_add_to_rfq.lower()
                        if (
                            "_pl_" in path
                            or "spdl" in path
                            or "psdl" in path
                            or "pl" in os.path.basename(path)
                        ):
                            path_dict[file_path_to_add_to_rfq] = 26
                        elif "dwg" in path or "drw" in path:
                            path_dict[file_path_to_add_to_rfq] = 27
                        elif "step" in path or "stp" in path:
                            path_dict[file_path_to_add_to_rfq] = 30
                        elif "zsp" in path or "speco" in path:
                            path_dict[file_path_to_add_to_rfq] = 33
                        elif ".cat" in path:
                            path_dict[file_path_to_add_to_rfq] = 16
                        elif "prt" in path:
                            path_dict[file_path_to_add_to_rfq] = 17
                        elif "lwg" in path:
                            path_dict[file_path_to_add_to_rfq] = 29
                        else:
                            path_dict[file_path_to_add_to_rfq] = None
                if value[13] == "Hardware":  # NOTE: need to add the '05-' function
                    item_pk = check_and_create_tooling(value[0])

                elif value[13] == "Tooling":
                    item_pk = self.data_base_conn.get_or_create_item(key)
                    if item_pk:
                        self.data_base_conn.insert_part_details_in_item_new(
                            item_pk, key, value
                        )
                    else:
                        item_pk = self.data_base_conn.create_item(
                            key,
                            party_pk,
                            value[15],
                            value[14],
                            value[2],
                            value[4],
                            item_type_fk=7,
                            description=value[0],
                            purchase=0,
                            forecast_on_mrp=0,
                            can_not_create_work_order=1,
                            can_not_invoice=1,
                            manufactured_item=1,
                            mps_item=0,
                            mps_on_mrp=0,
                            service_item=0,
                            bulk_ship=0,
                            ship_loose=0,
                        )
                elif value[13] == "Material":
                    item_pk = self.data_base_conn.get_or_create_item(key)
                    if item_pk:
                        self.data_base_conn.insert_part_details_in_item_new(
                            item_pk, key, value, item_type="Material"
                        )
                    else:
                        item_pk = self.data_base_conn.create_item(
                            key,
                            party_pk,
                            value[15],
                            value[14],
                            value[2],
                            value[4],
                            description=value[0],
                            item_type_fk=2,
                        )
                else:
                    item_pk = self.data_base_conn.get_or_create_item(key)
                    if item_pk:
                        self.data_base_conn.insert_part_details_in_item_new(
                            item_pk, key, value
                        )
                    else:
                        item_pk = self.data_base_conn.create_item(
                            key,
                            party_pk,
                            value[15],
                            value[14],
                            value[2],
                            value[4],
                            description=value[0],
                        )

                if self.customer_select_box.get():
                    matching_paths = {
                        path: pk for path, pk in path_dict.items() if key in path
                    }
                    for url, pk in matching_paths.items():
                        if restricted:
                            self.data_base_conn.upload_documents(
                                url,
                                item_fk=item_pk,
                                document_type_fk=2,
                                secure_document=1,
                                document_group_pk=pk,
                            )
                        else:
                            self.data_base_conn.upload_documents(
                                url,
                                item_fk=item_pk,
                                document_type_fk=2,
                                document_group_pk=pk,
                            )
            messagebox.showinfo("Success", "Item added successfully!")
            self.reset_gui()
        else:
            messagebox.showerror("ERROR", "Upload Parts to be added File")
            self.reset_gui()

    def update_rfq(self):
        """Updates RFQ by deleting old quotes and creating new quotes for a RFQ"""
        if (
            self.update_rfq_number_text.get()
            and self.customer_select_box.get()
            and self.file_path_PR_entry.get(0)
        ):
            rfq_pk = self.update_rfq_number_text.get()
            request_for_quote.reset_rfq(rfq_pk)
            loading_screen = LoadingScreen(self, max_progress=100)
            Thread(
                target=self.generate_rfq,
                args=(loading_screen,),
                kwargs={"update_rfq_pk": rfq_pk},
            ).start()
            # self.generate_rfq(loading_screen, update_rfq_pk=rfq_pk)
        else:
            messagebox.showerror("ERROR", "Please fill all required fields")


if __name__ == "__main__":
    r = RfqGen()
    r.mainloop()
    # filepath = r"C:\Users\saggarwal\Downloads\Test RFQ Big.xlsx"
    # info_dict = create_dict_from_excel(filepath)
    # r.create_rfq("test", "test1", "test2", info_dict)
