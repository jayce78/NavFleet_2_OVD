"""Script to convert NavFleet files into OVD recognised format, using
information that is relevant but requires further work/clarififcation 
on column naming conventions"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import re
from tkinter import simpledialog
import datetime

def get_vessel_name():
    """
    Prompt the user to input a vessel name.
    """
    vessel_name = simpledialog.askstring("Vessel Name", "Enter the Vessel Name:")
    if not vessel_name:
        raise ValueError("Vessel Name is required.")
    return vessel_name.strip()

# Utility Functions
def load_file(file_path, file_type="csv", encodings=None):
    """
    Load a file with fallback encoding options.
    """
    if encodings is None:
        encodings = ['utf-8', 'latin1', 'windows-1252']
    for encoding in encodings:
        try:
            if file_type == "csv":
                return pd.read_csv(file_path, encoding=encoding, on_bad_lines='skip')
            elif file_type == "excel":
                return pd.read_excel(file_path)
        except Exception:
            continue
    raise ValueError(f"Unable to load {file_path} with supported encodings.")

def convert_decimal_to_dms(decimal, is_latitude=True):
    """
    Convert decimal degrees to DMS (Degrees, Minutes, Direction).
    """
    if pd.isna(decimal):
        return None, None, None
    direction = 'N' if decimal >= 0 and is_latitude else 'S' if is_latitude else 'E' if decimal >= 0 else 'W'
    decimal = abs(decimal)
    degrees = int(decimal)
    minutes = round((decimal - degrees) * 60, 1)  # Round minutes to 1 decimal place
    return degrees, minutes, direction

def adjust_to_utc(report_to, timezone):
    """
    Adjust time to UTC based on the provided timezone.
    Handles date changes when the time crosses 00:00.
    """
    try:
        # Try parsing with a specific date format first
        formats = ["%d/%m/%Y %H:%M", "%m/%d/%Y %H:%M", "%Y-%m-%d %H:%M"]
        report_to_dt = None

        for fmt in formats:
            try:
                report_to_dt = datetime.datetime.strptime(str(report_to), fmt)
                break  # Stop trying once one format works
            except ValueError:
                continue

        # If all fail, fall back to default parsing
        if report_to_dt is None:
            report_to_dt = pd.to_datetime(report_to, errors='coerce', dayfirst=True)

        if pd.isna(report_to_dt) or not timezone:
            return None, None  # Return None for both time and date if invalid
        
        # Handle UTC timezone directly
        if timezone.strip().upper() == "UTC":
            return report_to_dt.time(), report_to_dt.date()

        # Extract timezone offset from the string
        match = re.search(r'TVA/GMT([+-]\d+)', timezone)
        if match:
            offset = int(match.group(1))  # Extract offset value
            # Adjust time based on the offset
            adjusted_time = report_to_dt - pd.to_timedelta(offset, unit='h')
            return adjusted_time.time(), adjusted_time.date()
        
        # If timezone format is not recognized
        return None, None
    except Exception as e:
        print(f"Error adjusting time to UTC: {e}")
        return None, None

def map_values(column, reference_df, reference_key, target_key):
    """
    Map values from one column to another based on a reference DataFrame.
    """
    reference_map = dict(zip(reference_df[reference_key], reference_df[target_key]))
    return column.map(reference_map).fillna(column)

def map_fuel_type_consumption(navfleet_data):
    """
    Map fuel type consumption values for Fuel Type 1, Fuel Type 2, and Fuel Type 3.
    """
    # Define mappings for ME, AE, and Boiler consumption columns
    fuel_type_to_ME_column = {
        'hfo': 'ME_Consumption_HFO',
        'vlsfo2020': 'ME_Consumption_HFO',  # Treat vlsfo2020 as HFO
        'lfo': 'ME_Consumption_LFO',
        'mgo': 'ME_Consumption_MGO',
        'ulsmgo2020': 'ME_Consumption_MGO',  # Treat vlsfo2020 as MGO
        'mdo': 'ME_Consumption_MDO',
        'lng': 'ME_Consumption_LNG',
        'lpgp': 'ME_Consumption_LPGP',
        'lpgb': 'ME_Consumption_LPGB',
        'm': 'ME_Consumption_M',
        'e': 'ME_Consumption_E'
    }
    fuel_type_to_AE_column = {
        'hfo': 'AE_Consumption_HFO',
        'vlsfo2020': 'AE_Consumption_HFO',  # Treat vlsfo2020 as HFO
        'lfo': 'AE_Consumption_LFO',
        'mgo': 'AE_Consumption_MGO',
        'ulsmgo2020': 'ME_Consumption_MGO',  # Treat vlsfo2020 as MGO
        'mdo': 'AE_Consumption_MDO',
        'lng': 'AE_Consumption_LNG',
        'lpgp': 'AE_Consumption_LPGP',
        'lpgb': 'AE_Consumption_LPGB',
        'm': 'AE_Consumption_M',
        'e': 'AE_Consumption_E'
    }
    fuel_type_to_Boiler_column = {
        'hfo': 'Boiler_Consumption_HFO',
        'vlsfo2020': 'Boiler_Consumption_HFO',  # Treat vlsfo2020 as HFO
        'lfo': 'Boiler_Consumption_LFO',
        'mgo': 'Boiler_Consumption_MGO',
        'ulsmgo2020': 'ME_Consumption_MGOO',  # Treat vlsfo2020 as MGO
        'mdo': 'Boiler_Consumption_MDO',
        'lng': 'Boiler_Consumption_LNG',
        'lpgp': 'Boiler_Consumption_LPGP',
        'lpgb': 'Boiler_Consumption_LPGB',
        'm': 'Boiler_Consumption_M',
        'e': 'Boiler_Consumption_E'
    }

    # Ensure all columns exist in navfleet_data and initialize as float
    for col in fuel_type_to_ME_column.values():
        if col not in navfleet_data:
            navfleet_data[col] = 0.0  # Initialize as float
        else:
            navfleet_data[col] = navfleet_data[col].astype(float)

    for col in fuel_type_to_AE_column.values():
        if col not in navfleet_data:
            navfleet_data[col] = 0.0  # Initialize as float
        else:
            navfleet_data[col] = navfleet_data[col].astype(float)

    for col in fuel_type_to_Boiler_column.values():
        if col not in navfleet_data:
            navfleet_data[col] = 0.0  # Initialize as float
        else:
            navfleet_data[col] = navfleet_data[col].astype(float)

    # Map Fuel Type 1 to ME_Consumption_* columns
    if 'Fuel Type 1' in navfleet_data and 'Fuel Type 1 ME (MT)' in navfleet_data:
        navfleet_data['Fuel Type 1'] = navfleet_data['Fuel Type 1'].astype(str).fillna("")
        for fuel_type, target_col in fuel_type_to_ME_column.items():
            mask = navfleet_data['Fuel Type 1'].str.strip().str.lower() == fuel_type
            navfleet_data.loc[mask, target_col] += navfleet_data['Fuel Type 1 ME (MT)'].fillna(0)

    # Map Fuel Type 1 to AE_Consumption_* columns
    if 'Fuel Type 1' in navfleet_data and 'Fuel Type 1 AE (MT)' in navfleet_data:
        navfleet_data['Fuel Type 1'] = navfleet_data['Fuel Type 1'].astype(str).fillna("")
        for fuel_type, target_col in fuel_type_to_ME_column.items():
            mask = navfleet_data['Fuel Type 1'].str.strip().str.lower() == fuel_type
            navfleet_data.loc[mask, target_col] += navfleet_data['Fuel Type 1 AE (MT)'].fillna(0)

    # Map Fuel Type 1 to Boiler_Consumption_* columns
    if 'Fuel Type 1' in navfleet_data and 'Fuel Type 1 Aux. Boiler (MT)' in navfleet_data:
        navfleet_data['Fuel Type 1'] = navfleet_data['Fuel Type 1'].astype(str).fillna("")
        for fuel_type, target_col in fuel_type_to_ME_column.items():
            mask = navfleet_data['Fuel Type 1'].str.strip().str.lower() == fuel_type
            navfleet_data.loc[mask, target_col] += navfleet_data['Fuel Type 1 Aux. Boiler (MT)'].fillna(0)

    # Map Fuel Type 2 to AE_Consumption_* columns
    if 'Fuel Type 2' in navfleet_data and 'Fuel Type 2 Total (MT)' in navfleet_data:
        navfleet_data['Fuel Type 2'] = navfleet_data['Fuel Type 2'].astype(str).fillna("")
        for fuel_type, target_col in fuel_type_to_AE_column.items():
            mask = navfleet_data['Fuel Type 2'].str.strip().str.lower() == fuel_type
            navfleet_data.loc[mask, target_col] += navfleet_data['Fuel Type 2 Total (MT)'].fillna(0)

    # Map Fuel Type 3 to Boiler_Consumption_* columns
    if 'Fuel Type 3' in navfleet_data and 'Fuel Type 3 Total (MT)' in navfleet_data:
        navfleet_data['Fuel Type 3'] = navfleet_data['Fuel Type 3'].astype(str).fillna("")
        for fuel_type, target_col in fuel_type_to_Boiler_column.items():
            mask = navfleet_data['Fuel Type 3'].str.strip().str.lower() == fuel_type
            navfleet_data.loc[mask, target_col] += navfleet_data['Fuel Type 3 Total (MT)'].fillna(0)

    return navfleet_data

def transform_navfleet_data(navfleet_data):
    """
    Transform NavFleet data to match the OVD structure.
    """
    def process_event(event_text):
        """
        Process the event text according to the specified rules.
        """
        if pd.isna(event_text):
            return event_text

        event_text = str(event_text).strip()

        # If the event is "Sea", change to "Noon (Position) - Sea passage"
        if 'Sea' in event_text:
            return "Noon (Position) - Sea passage"
        
        # If the event is "Port", change to "Noon (Position) Port"
        if 'Port' in event_text:
            return "Noon (Position) Port"
        
        # If the event contains "Arrival" or "Departure", only retain that word
        if 'Arrival' in event_text:
            return "Arrival"
        if 'Departure' in event_text:
            return "Departure"
        
        # Return the event text unchanged if no rule applies
        return event_text

    transformed = pd.DataFrame({
        'IMO': navfleet_data['IMO No'],
        'Date_UTC': navfleet_data.apply(lambda row: adjust_to_utc(row['Report To'], row['Timezone'])[1], axis=1),  # Adjusted date
        'Time_UTC': navfleet_data.apply(lambda row: adjust_to_utc(row['Report To'], row['Timezone'])[0], axis=1),  # Adjusted time
        'Event': navfleet_data['Type'].apply(lambda x: process_event(x)),  # Apply the event transformation
        'Time_Since_Previous_Report': navfleet_data['Report Period'],
        'Distance': navfleet_data['GPS Dist.'],
        'Latitude_Degree': navfleet_data['Latitude'].apply(lambda x: convert_decimal_to_dms(x, True)[0]),
        'Latitude_Minutes': navfleet_data['Latitude'].apply(lambda x: convert_decimal_to_dms(x, True)[1]),
        'Latitude_North_South': navfleet_data['Latitude'].apply(lambda x: convert_decimal_to_dms(x, True)[2]),
        'Longitude_Degree': navfleet_data['Longitude'].apply(lambda x: convert_decimal_to_dms(x, False)[0]),
        'Longitude_Minutes': navfleet_data['Longitude'].apply(lambda x: convert_decimal_to_dms(x, False)[1]),
        'Longitude_East_West': navfleet_data['Longitude'].apply(lambda x: convert_decimal_to_dms(x, False)[2]),
        'Voyage_From': navfleet_data['Next Port'].shift(1),
        'Voyage_To': navfleet_data['Next Port'],
        'Cargo_Mt': navfleet_data['Cargo Quantity'],
        'Time_Elapsed_Sailing': navfleet_data['Report Period'],  # Duplicate value as we cannot calculate time drifting
        'Wind_Force_Bft': navfleet_data['True Wind Force'],
        'Wind_Force_Kn': navfleet_data['WNI Relative Wind Speed'],
        'Wind_Dir_Degree': pd.to_numeric(navfleet_data['WNI Relative Wind Direction'], errors='coerce').fillna(0).astype(int),
        'Sea_state_Force_Douglas': navfleet_data['Sea Swell'],
        #'Current_Dir_Degree': navfleet_data['Current Direction, Compass'],
        'Current_Speed': navfleet_data['WNI Current Speed'],
        'Current_Dir': pd.to_numeric(navfleet_data['WNI Current Direction'], errors='coerce').fillna(0).astype(int),
        'Speed_Through_Water': navfleet_data['Log Speed'],
        'Speed_GPS': navfleet_data['GPS Speed'],
        'Comments': navfleet_data['Comments'],
    })

    # Map fuel type consumption
    navfleet_data = map_fuel_type_consumption(navfleet_data)
   
    # Add Main Engine_Consumption_* columns to the transformed DataFrame
    for col in navfleet_data.columns:
        if col.startswith('ME_Consumption_'):
            transformed[col] = navfleet_data[col]

    # Add Aux Engine_Consumption_* columns to the transformed DataFrame
    for col in navfleet_data.columns:
        if col.startswith('AE_Consumption_'):
            transformed[col] = navfleet_data[col]

    # Add Boiler_Consumption_* columns to the transformed DataFrame
    for col in navfleet_data.columns:
        if col.startswith('Boiler_Consumption_'):
            transformed[col] = navfleet_data[col]

    # Combine Date_UTC and Time_UTC into a single datetime column
    transformed['DateTime_UTC'] = pd.to_datetime(
        transformed['Date_UTC'].astype(str) + ' ' + transformed['Time_UTC'].astype(str),
        errors='coerce'
    )

    # Sort by the combined datetime column
    transformed = transformed.sort_values(by='DateTime_UTC').reset_index(drop=True)

        # Optionally drop the combined datetime column after sorting
    transformed = transformed.drop(columns=['DateTime_UTC'])

    return transformed

def update_voyage_columns(output_data, reference_data):
    """
    Update Voyage_From and Voyage_To columns using a reference file.
    Ensure Date_UTC remains as a date-only field.
    """
    # Ensure the reference data has the expected structure
    reference_data['NameWoDiacritics'] = reference_data['NameWoDiacritics'].str.lower().str.strip()
    reference_data['PortCode'] = (reference_data['Country'] + reference_data['Location']).str[:5]

    for column in ['Voyage_From', 'Voyage_To']:
        if column in output_data:
            output_data[column] = map_values(
                output_data[column].str.lower().str.strip(),
                reference_data,
                'NameWoDiacritics',
                'PortCode'
            )
        else:
            print(f"Column {column} not found in output data.")

    # Ensure Date_UTC is date-only
    if 'Date_UTC' in output_data:
        output_data['Date_UTC'] = pd.to_datetime(output_data['Date_UTC'], errors='coerce').dt.date

    return output_data

# GUI Functions
def process_transformation():
    try:
        file_path = filedialog.askopenfilename(title="Select NavFleet CSV File")
        if not file_path:
            raise ValueError("No file selected.")

        # Load NavFleet data
        navfleet_data = load_file(file_path)

        # Transform data
        transformed_data = transform_navfleet_data(navfleet_data)

        # Get Vessel Name from the user
        vessel_name = get_vessel_name()

        # Generate a file name using vessel name and current date
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"{vessel_name}_Step1_Formatting_{timestamp}.xlsx"

        # Save the file automatically
        save_path = filedialog.asksaveasfilename(initialfile=default_name, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            transformed_data.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Data successfully transformed and saved as:\n{save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Transformation failed: {e}")

def process_update():
    try:
        output_file = filedialog.askopenfilename(title="Select Output File to Update")
        update_file = filedialog.askopenfilename(title="Select PortCode File")

        if not output_file or not update_file:
            raise ValueError("Files not selected.")

        # Load files
        output_data = load_file(output_file, file_type="excel")
        reference_data = load_file(update_file)

        # Update voyage columns
        updated_data = update_voyage_columns(output_data, reference_data)

        # Get Vessel Name from the user
        #vessel_name = get_vessel_name()

        # Generate a file name using vessel name and current date
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"OVD_READY_{timestamp}.xlsx"

        # Save the file automatically
        save_path = filedialog.asksaveasfilename(initialfile=default_name, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            updated_data.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Voyage columns updated and saved as:\n{save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Update failed: {e}")

def main():
    root = tk.Tk()
    root.title("NavFleet Data Processor")
    root.geometry("400x300")

    tk.Label(root, text="NavFleet to OVD Formatter", font=("Arial", 16, "bold")).pack(pady=20)
    tk.Button(root, text="Transform Data", command=process_transformation, padx=10, pady=5, bg="lightgreen", fg="black").pack(pady=10)
    tk.Button(root, text="Update Voyage Columns", command=process_update, padx=10, pady=5, bg="lightgreen", fg="black").pack(pady=10)
    tk.Button(root, text="Exit", command=root.quit, padx=10, pady=5, bg="red", fg="black").pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
