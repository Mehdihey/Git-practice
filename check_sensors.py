import yaml
import openpyxl
import csv 
import json
# Read settings from the YAML config file ---
# safe_load() converts the YAML content into a Python dict.
# (safe_load is used instead of load, because it only reads plain data
# and never executes code from the file.)

with open("config.yml", encoding = "utf-8")as file:
    config = yaml.safe_load(file)

max_days = config ["max_days_since_calibration"] # threshold in days (int)
output = config["output_file"] # name of the JSON output file


#Read sensor information from the Excel file ---
wb = openpyxl.load_workbook("sensors.xlsx")
ws = wb.active # the first (active) worksheet


# Build a lookup dict keyed by sensor_id, e.g.
sensors = {}

for row in ws.iter_rows(min_row=2, values_only=True): # min_row=2 skips the header
    sensor_id, lab_room, owner = row # unpack the row tuple
    sensors[sensor_id] = {"lab_room": lab_room, "owner":owner}
#print(sensors)


# Read calibration logs, join with sensor info and filter ---

overdue = [] # list of dicts for sensors that need calibration
with open("calibrations.csv", encoding="utf-8", newline="") as file:
    reader = csv.DictReader(file) # each row becomes a dict using the header as keys
    for row in reader:
        sensor_id = row["sensor_id"] # CSV values are always read as strings, so convert to int before comparing
        days_since_calibration = int(row["days_since_calibration"])

        days = int(row["days_since_calibration"])
        if days > max_days:
            info = sensors[sensor_id] # join: look up room and owner by sensor_id
            overdue.append({
                "sensor_id":sensor_id,
                "lab_room":info["lab_room"],
                "owner":info["owner"],
                "days_since_calibration": days,
        
    })
#Export the overdue sensors as a formatted JSON array.
with open(output, "w", encoding="utf-8") as file:
    json.dump(overdue,file, indent=2)

print(f"overdue sensors written to {output}")