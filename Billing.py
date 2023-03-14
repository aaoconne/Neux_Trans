import openpyxl
from datetime import datetime
import os 

# Define filename for Excel spreadsheet
today = datetime.today().strftime('%Y-%m-%d')

# Create new workbook if one doesn't already exist for today's date
if os.path.exists(f'./Billing_{today}.xlsx'):
    workbook = openpyxl.load_workbook(f'./Billing_{today}.xlsx')
    sheet = workbook.active 
else:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    # Define headers for worksheet & append heads to worksheet 
    hearder_row = ["Date", "Translator", "InterpretedFor", "Office", "Name", "Address", "AppointmentTime", "ArrivalTimeInterpretor", "ArrivalTimePatient", "EndTime", "ServicesProvided", "TotalMiles", "ParkingGarage", "Paid", "Billed", "Client"]
    sheet.append(hearder_row)

while True:
    # Collect user input for billing data 
    date = input("Date (YYYY-MM-DD): ")
    
    if datetime.strptime(date, '%Y-%m-%d') < datetime.today():
        confirm = input("You have entered a date prior to today. Are you sure you want to proceed? (Y/N): ")
        if confirm.upper() != "Y":
            continue 
            
    translator = input("Translator: ")
    interpreted_for = input("Interpreted For (Medical/Legal/Translation): ")
    office = input("Office: ")
    name = input("Name: ")
    address = input("Address: ")
    appointment_time = input("Appointment Time (HH:MM): ")
    arrival_time_interpreter = input("Arrival Time (Interpreter) (HH:MM): ")
    arrival_time_patient = input("Arrival Time (Patient) (HH:MM): ")
    end_time = input("End Time (HH:MM): ")
    services_provided = input("Services Provided: ")
    total_miles = input("Total Miles: ")
    parking_garage = input("Parking Garage: ")
    paid = input("Paid: ")
    billed = input("Billed: ")
    client = input("Client: ")
        
    # Create new row for worksheet 
    sheet.append([date, translator, interpreted_for, office, name, address, appointment_time, arrival_time_interpreter, arrival_time_patient, end_time, services_provided, total_miles, parking_garage, paid, billed, client])
    add_another_row = input("Do you need to fill out another billing statement?: (Y/N): ")
    if add_another_row.upper() != "Y":
        file_name = f"BillingStatement_{date.replace('/', '_')}.xlsx"
        file_path = os.path.join('C:\BillingStatements', file_name)
    
        # Save workbook
        workbook.save(file_path)
    
        print(f"New billing statement saved to {file_path}") 
        break   