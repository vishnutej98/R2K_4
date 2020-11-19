#!Python3
# app.py - This application will calculate the Gear ratio's
# Vehicle Entry speed and calculated Exit speed for the vehicle.
# From this Calculation based on user requirement the calculation can be stored
# Files can be saved in .PDF, .xlsx, .txt
from math import pi  # Ref 1import
from fpdf import FPDF   # Ref 2import
from openpyxl import Workbook   # Ref 3import
from openpyxl.styles import Font    # Ref 4import

print('''
Data Sheet of the Vehicle
please enter the required specification for the vehicle
''')

"""
The following line from ln 6 to ln 15, we are getting user input for the particular variables.
The values are stored based on the variable units.
"""
Vehicle_name = str(input("Enter the Vehicle Name: ")).title()
Gear = int(input("No. of Gears: "))
if Gear >= 6:
    print("Sorry! this file will calculate only for 4 or 5 speed gear box.")
    pass
else:
    Primary = float(input("Primary transmission ratio: "))
    Secondary = float(input("Secondary transmission ratio: "))
    Rolling_Radius = float(input("Rolling radius: "))
    peak_power_rpm = float(input("Peak Power RPM: "))
    Gear_2 = float(input("2nd Gear ratio: "))
    Gear_3 = float(input("3rd Gear ratio: "))
    total_2nd_gear_ratio = round(Primary * Secondary * Gear_2, 3)
    total_3rd_gear_ratio = round(Primary * Secondary * Gear_3, 3)


    def fact(pi, rolling_radius):   # Ref 1import
        factor = round(2 * pi * rolling_radius * 0.06, 3)
        return factor


    fact = fact(pi, Rolling_Radius)


    class calculation():
        @staticmethod
        def RPM_to_KPH_ratio_2nd_Gear(total_2nd_gear_ratio, fact):
            RPM_to_KPH_ratio_2nd_Gear = round(total_2nd_gear_ratio / fact, 3)
            return RPM_to_KPH_ratio_2nd_Gear

        @staticmethod
        def RPM_to_KPH_ratio_3rd_Gear(total_3rd_gear_ratio, fact):
            RPM_to_KPH_ratio_3rd_Gear = round(total_3rd_gear_ratio / fact, 3)
            return RPM_to_KPH_ratio_3rd_Gear

        @staticmethod
        def Entry_Speed_2nd_Gear(peak_power_rpm, total_2nd_gear_ratio, fact):
            Entry_Speed_2nd_Gear = round(float(((0.75 * peak_power_rpm) / total_2nd_gear_ratio) * fact), 2)
            return Entry_Speed_2nd_Gear

        @staticmethod
        def Entry_Speed_3rd_Gear(peak_power_rpm, Total_3rd_Gear_ratio, fact):
            Entry_Speed_3rd_Gear = round(float(((0.75 * peak_power_rpm) / Total_3rd_Gear_ratio) * fact), 2)
            return Entry_Speed_3rd_Gear

        @staticmethod
        def Exit_Speed_2nd_Gear(peak_power_rpm, Total_2nd_Gear_ratio, fact):
            Exit_Speed_2nd_Gear = round(float((peak_power_rpm / Total_2nd_Gear_ratio) * fact), 2)
            return Exit_Speed_2nd_Gear

        @staticmethod
        def Exit_Speed_3rd_Gear(peak_power_rpm, Total_3rd_Gear_ratio, fact):
            Exit_Speed_3rd_Gear = round(float((peak_power_rpm / Total_3rd_Gear_ratio) * fact), 2)
            return Exit_Speed_3rd_Gear


    spec_name = [
        "Vehicle name", "Gear Box", "Factor = (2 * PI * R * 0.06)", "Primary Gear Ratio",
        "Secondary Gear Ratio", "Dynamic Rolling Radius in meters", "Max Peak Power RPM", "Total 2ndGear ratio",
        "Total 3rdGear ratio", "2ndGear RPMtoKPH ratio", "3rdGear RPMtoKPH ratio", "Entry Speed at 2nd Gear in KPH",
        "Exit Speed at 2nd Gear in KPH", "Entry Speed at 3rd Gear in KPH", "Exit Speed at 3rd Gear in KPH"
    ]
    spec_list = [
        Vehicle_name, Gear, fact, Primary, Secondary, Rolling_Radius, peak_power_rpm, total_2nd_gear_ratio,
        total_3rd_gear_ratio,
        calculation.RPM_to_KPH_ratio_2nd_Gear(total_2nd_gear_ratio, fact),
        calculation.RPM_to_KPH_ratio_3rd_Gear(total_3rd_gear_ratio, fact),
        calculation.Entry_Speed_2nd_Gear(peak_power_rpm, total_2nd_gear_ratio, fact),
        calculation.Exit_Speed_2nd_Gear(peak_power_rpm, total_2nd_gear_ratio, fact),
        calculation.Entry_Speed_3rd_Gear(peak_power_rpm, total_3rd_gear_ratio, fact),
        calculation.Exit_Speed_3rd_Gear(peak_power_rpm, total_3rd_gear_ratio, fact)

    ]
    print(f"The data will be calculated for {Vehicle_name}.")
    print("""
    Please select the option below to save the data in a file:
    
        1. PDF(.pdf)
        2. Excel(.xlsx)
        3. Text(.txt)
    """)

    user_input = int(input("Enter numerical value only: "))

    if user_input == 1:
        file = open("data.txt", "w")
        for i in range(0, 7):
            file.write(f"{spec_name[i]} : {spec_list[i]}\n")

        file.write("\n")
        file.write("\n")

        for j in range(7, 11):
            file.write(f"{spec_name[j]} : {spec_list[j]}\n")

        file.write("\n")
        file.write("\n")

        for k in range(11, 15):
            file.write(f"{spec_name[k]} : {spec_list[k]}\n")
        file.close()

        # Ref 2import
        # from fpdf import FPDF

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Times", size=14)
        read_file = open("data.txt", "r")
        for x in read_file:
            pdf.cell(200, 10, txt=x, ln=1, align='L')
        pdf.output(f"{Vehicle_name}.pdf")
        print(f"Data is saved in PDF file named as {Vehicle_name}.")

    elif user_input == 2:
        # Ref 3import
        # from openpyxl import Workbook
        # Ref 4import
        # from openpyxl.styles import Font

        book = Workbook()
        Sheet = book.active
        n = 0
        m = 0
        for g in range(5, 20):  # Excel Result loop.
            Sheet[f'C{g}'] = f'{spec_name[n]}'
            Sheet[f'D{g}'] = f'{spec_list[m]}'
            n += 1
            m += 1

        #  Making Rows(Top two and bottom four rows) Bold.
        for b in list(range(4, 6)) + list(range(16, 21)):
            Sheet[f'C{b}'].font = Font(bold=True)
            Sheet[f'D{b}'].font = Font(bold=True)
            Sheet.column_dimensions['C'].width = 27.33
        book.save(str(Vehicle_name) + ".xlsx")
        print(f"Data is saved in Excel file named as {Vehicle_name}.")

    elif user_input == 3:
        file = open(f"{Vehicle_name}.txt", "w")
        for i in range(0, 7):
            file.write(f"{spec_name[i]}: {spec_list[i]}\n")

        file.write("\n")
        file.write("\n")

        for j in range(7, 11):
            file.write(f"{spec_name[j]}: {spec_list[j]}\n")

        file.write("\n")
        file.write("\n")

        for k in range(11, 15):
            file.write(f"{spec_name[k]}: {spec_list[k]}\n")
        file.close()
        print(f"Data is saved in Text file named as {Vehicle_name}.")
    else:
        print('Input Error!')
        print('Enter valid input.')
        print('The output has not been saved in a file.')

    print("Entry 2nd gear: {}".format(calculation.Entry_Speed_2nd_Gear(peak_power_rpm, total_2nd_gear_ratio, fact)))
    print("Entry 3rd gear: {}".format(calculation.Entry_Speed_3rd_Gear(peak_power_rpm, total_3rd_gear_ratio, fact)))
    print("Exit speed in 2nd gear: {}".format(
        calculation.Exit_Speed_2nd_Gear(peak_power_rpm, total_2nd_gear_ratio, fact)))
    print("Exit speed in 3rd gear: {}".format(
        calculation.Exit_Speed_3rd_Gear(peak_power_rpm, total_3rd_gear_ratio, fact)))

temp = input("please press any key to quit!")
print("Thanks! for using this application.")
