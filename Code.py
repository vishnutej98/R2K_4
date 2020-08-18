print("""
Data Sheet of the Vehicle.
please enter the required specification for the vehicle.
""")

Gear = int(input("No of Gears in The Vehicle: "))
if Gear >= 6:
    print("Sorry! the application is only created for 4 and 5 speed Gear Box.")

else:
    print(f"""
    The Vehicle will have {Gear}.
    The Entry speed and Exit speed will be calcuated for 2nd and 3rd Gear.
    """)
    from math import pi
    from openpyxl import Workbook
    from openpyxl.styles import Font

    PI = round(pi, 4)
    book = Workbook()
    Sheet = book.active
    Vehicle_name = str(input("Enter the Vehicle Name: ")).title()
    print(f"The Data is Calculated for {Vehicle_name}.")
    P = float(input("Primary Gear Ratio: "))  # Primary Transmission Ratio.
    S = float(input("Secondary Gear Ratio: "))  # Secondary Transmission Ratio.
    R = float(input('Wheel Rolling radius in meters: '))  # Dynamic Rolling Radius in meters.
    PPM = int(input('Peak Power RPM: '))  # Peak Power RPM.
    G2 = float(input('2nd Gear Ratio: '))  # 2nd Gear Ratio.
    G3 = float(input('3rd Gear Ratio: '))  # 3rd Gear Ratio.


    class calculation():
        def Gear_ratio_2nd(P, S, G2):
            R2 = P * S * G2
            return R2

        def Gear_ratio_3rd(P, S, G3):
            R3 = P * S * G3
            return R3

        def factor(PI, R):
            Factor = round(2 * PI * R * 0.06, 3)
            return Factor


    Gear_ratio_2nd = calculation.Gear_ratio_2nd(P, S, G2)
    Gear_ratio_3rd = calculation.Gear_ratio_3rd(P, S, G3)
    factor = calculation.factor(PI, R)


    class converstion_calculation():
        def RpmtoKph_G2(Gear_ratio_2nd, factor):
            RK2 = round(Gear_ratio_2nd / factor)
            return RK2

        def RpmtoKph_G3(Gear_ratio_3rd, factor):
            RK3 = round(Gear_ratio_3rd / factor)
            return RK3


    RK2 = converstion_calculation.RpmtoKph_G2(Gear_ratio_2nd, factor)
    RK3 = converstion_calculation.RpmtoKph_G3(Gear_ratio_3rd, factor)


    class Entry_Exit_Speed_Cal():
        def Entry_Speed_G2(PPM, Gear_ratio_2nd, factor):
            Entry_Speed_G2 = round(float(((0.75 * PPM) / Gear_ratio_2nd) * factor), 2)
            return Entry_Speed_G2

        def Entry_Speed_G3(PPM, Gear_ratio_3rd, factor):
            Entry_Speed_G3 = round(float(((0.75 * PPM) / Gear_ratio_3rd) * factor), 2)
            return Entry_Speed_G3

        def Exit_Speed_G2(PPM, Gear_ratio_2nd, factor):
            Exit_Speed_G2 = round(float((PPM / Gear_ratio_2nd) * factor), 2)
            return Exit_Speed_G2

        def Exit_Speed_G3(PPM, Gear_ratio_3rd, factor):
            Exit_Speed_G3 = round(float((PPM / Gear_ratio_3rd) * factor), 2)
            return Exit_Speed_G3


    ES2 = Entry_Exit_Speed_Cal.Entry_Speed_G2(PPM, Gear_ratio_2nd, factor)
    ES3 = Entry_Exit_Speed_Cal.Entry_Speed_G3(PPM, Gear_ratio_3rd, factor)
    EXS2 = Entry_Exit_Speed_Cal.Exit_Speed_G2(PPM, Gear_ratio_2nd, factor)
    EXS3 = Entry_Exit_Speed_Cal.Exit_Speed_G3(PPM, Gear_ratio_3rd, factor)

    print(f"""
        The {Vehicle_name} with {Gear} Gears has the following values has been saved in the file name of {Vehicle_name}.
        """)
    spec_list = [
        Vehicle_name, Gear, factor, P, S, R, PPM, Gear_ratio_2nd, Gear_ratio_3rd, RK2, RK3, ES2, EXS2, ES3, EXS3
    ]
    spec_names = [
        "Vehicle name", "Gear Box", "Factor = (2 * PI * R * 0.06)", "Primary Gear Ratio",
        "Secondary Gear Ratio", "Dynamic Rolling Radius in meters", "Max Peak Power RPM", "Total 2ndGear ratio",
        "Total 3rdGear ratio", "2ndGear RPMtoKPH ratio", "3rdGear RPMtoKPH ratio", "Entry Speed at 2nd Gear in KPH",
        "Exit Speed at 2nd Gear in KPH", "Entry Speed at 3rd Gear in KPH", "Exit Speed at 3rd Gear in KPH"
    ]

    # Output Result Display Variables.
    i = 0
    j = 0
    for f in range(0, 8):
        # Output Result Display loop.
        print(f"{spec_names[i]}: {spec_list[j]}")
        i += 1
        j += 1

    # Excel Result Variables.
    n = 0
    m = 0
    for g in range(5, 20):  # Excel Result loop.
        Sheet[f'C{g}'] = f'{spec_names[n]}'
        Sheet[f'D{g}'] = f'{spec_list[m]}'
        n += 1
        m += 1

    #  Making Rows(Top two and bottom four rows) Bold.
    for b in list(range(4, 6)) + list(range(16, 21)):
        Sheet[f'C{b}'].font = Font(bold=True)
        Sheet[f'D{b}'].font = Font(bold=True)
        Sheet.column_dimensions['C'].width = 27.33
    book.save(str(Vehicle_name) + ".xlsx")
    print("""
    The Limit for the Vehicles based on capacity of engine the rules and
    regulations are set as follows:
    1. from 80cc to 175cc the limit is 77dB.
    2. above 175cc the limit is 80dB.
    """)
    print(f"Data is saved in Excel file named as {Vehicle_name}.")
