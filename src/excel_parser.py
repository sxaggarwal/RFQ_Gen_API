import pandas as pd
from typing import Optional
from pydantic import BaseModel, ValidationError


class PartData(BaseModel):
    part_number: str
    description: Optional[str]
    length: float = 0.0
    thickness: float = 0.0
    width: float = 0.0
    weight: float = 0.0
    material: Optional[str]
    finish_code: Optional[str]
    heat_treat: Optional[str]
    drawing_number: Optional[str]
    drawing_revision: Optional[str]
    quantity_required: int
    pl_revision: Optional[str]
    assy_for: Optional[str]
    hardware_or_supplies: Optional[str]
    stock_length: float = 0.0
    stock_width: float = 0.0
    stock_thickness: float = 0.0


def sanitize_value(value, default=None):
    """Sanitizes NaN values and strips strings."""
    if pd.isna(value):  # More robust check than math.isnan
        return default
    if isinstance(value, str):
        return value.strip()
    return value


def create_dict_from_excel_new(filepath: str):
    """Converts the Excel file into a dictionary with part number as key."""

    df = pd.read_excel(filepath, dtype=str).fillna("")  

    required_columns = {
        "Part": "part_number",
        "DESCRIPTION": "description",
        "PartLength": "length",
        "Thickness": "thickness",
        "PartWidth": "width",
        "Weight": "weight",
        "Material": "material",
        "FinishCode": "finish_code",
        "HeatTreat": "heat_treat",
        "DrawingNumber": "drawing_number",
        "DrawingRevision": "drawing_revision",
        "QuantityRequired": "quantity_required",
        "PLRevision": "pl_revision",
        "AssyFor": "assy_for",
        "Hardware/Tooling": "hardware_or_supplies",
        "StockLength": "stock_length",
        "StockWidth": "stock_width",
        "StockThickness": "stock_thickness"
    }

    # error checking for correct template
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")

    df = df.rename(columns=required_columns)

    df = df.map(sanitize_value)

    # Ensure numeric fields are converted
    numeric_fields = ["length", "thickness", "width", "weight", "stock_length", "stock_width", "stock_thickness"]
    df[numeric_fields] = df[numeric_fields].apply(pd.to_numeric, errors="coerce").fillna(0.0)

    df["quantity_required"] = df["quantity_required"].apply(pd.to_numeric, errors="coerce").fillna(0).astype(int)

    my_dict = {}

    errors = []
    for idx, (_, row) in enumerate(df.iterrows(), start=1):
        part_number = row["part_number"] or f"Tool-{idx}"  # Fallback naming

        try:
            part_data = PartData(**row.to_dict())  # Validate with Pydantic
        except ValidationError as e:
            errors.append(f"Row {idx}: {e}")
            continue  # Skip invalid row

        # Ensure unique part numbers
        original_part_number = part_number
        suffix = 1
        while part_number in my_dict:
            part_number = f"{original_part_number}_____{suffix}"
            suffix += 1

        # my_dict[part_number] = [val for  val in list(part_data.model_dump().values())[1:]]  # we skip the first one as its the key
        my_dict[part_number] = part_data.model_dump()

    if errors:
        raise ValueError(f"Data validation failed:\n" + "\n".join(errors))

    return my_dict

