import pandas as pd
import os

# COATO mapping dictionary
coato_mapping = {
    # Qoraqalpog'iston Respublikasi
    1735401: {"region": "Qoraqalpog'iston Respublikasi", "district": "Nukus shahri"},
    1735204: {"region": "Qoraqalpog'iston Respublikasi", "district": "Amudaryo tumani"},
    1735207: {"region": "Qoraqalpog'iston Respublikasi", "district": "Beruniy tumani"},
    1735209: {"region": "Qoraqalpog'iston Respublikasi", "district": "Bo'zatov tumani"},
    1735211: {"region": "Qoraqalpog'iston Respublikasi", "district": "Qorao'zak tumani"},
    1735212: {"region": "Qoraqalpog'iston Respublikasi", "district": "Kegeyli tumani"},
    1735215: {"region": "Qoraqalpog'iston Respublikasi", "district": "Qo'ng'irot tumani"},
    1735218: {"region": "Qoraqalpog'iston Respublikasi", "district": "Qanliko'l tumani"},
    1735222: {"region": "Qoraqalpog'iston Respublikasi", "district": "Mo'ynoq tumani"},
    1735225: {"region": "Qoraqalpog'iston Respublikasi", "district": "Nukus tumani"},
    1735228: {"region": "Qoraqalpog'iston Respublikasi", "district": "Taxiatash tumani"},
    1735230: {"region": "Qoraqalpog'iston Respublikasi", "district": "Taxtako'pir tumani"},
    1735233: {"region": "Qoraqalpog'iston Respublikasi", "district": "To'rtko'l tumani"},
    1735236: {"region": "Qoraqalpog'iston Respublikasi", "district": "Xo'jayli tumani"},
    1735240: {"region": "Qoraqalpog'iston Respublikasi", "district": "Chimboy tumani"},
    1735243: {"region": "Qoraqalpog'iston Respublikasi", "district": "Shumanay tumani"},
    1735250: {"region": "Qoraqalpog'iston Respublikasi", "district": "Ellikqal'a tumani"},

    # Andijon
    1703401: {"region": "Andijon", "district": "Andijon shahri"},
    1703408: {"region": "Andijon", "district": "Xonobod shahri"},
    1703202: {"region": "Andijon", "district": "Oltinko'l tumani"},
    1703203: {"region": "Andijon", "district": "Andijon tumani"},
    1703206: {"region": "Andijon", "district": "Baliqchi tumani"},
    1703209: {"region": "Andijon", "district": "Bo'ston tumani"},
    1703210: {"region": "Andijon", "district": "Buloqboshi tumani"},
    1703211: {"region": "Andijon", "district": "Jalaquduq tumani"},
    1703214: {"region": "Andijon", "district": "Izboskan tumani"},
    1703217: {"region": "Andijon", "district": "Ulug'nor tumani"},
    1703220: {"region": "Andijon", "district": "Qo'rg'ontepa tumani"},
    1703224: {"region": "Andijon", "district": "Asaka tumani"},
    1703227: {"region": "Andijon", "district": "Marhamat tumani"},
    1703230: {"region": "Andijon", "district": "Shahrixon tumani"},
    1703232: {"region": "Andijon", "district": "Paxtaobod tumani"},
    1703236: {"region": "Andijon", "district": "Xo'jaobod tumani"},

    # Buxoro
    1706401: {"region": "Buxoro", "district": "Buxoro shahri"},
    1706403: {"region": "Buxoro", "district": "Kogon shahri"},
    1706204: {"region": "Buxoro", "district": "Olot tumani"},
    1706207: {"region": "Buxoro", "district": "Buxoro tumani"},
    1706212: {"region": "Buxoro", "district": "Vobkent tumani"},
    1706215: {"region": "Buxoro", "district": "G'ijduvon tumani"},
    1706219: {"region": "Buxoro", "district": "Kogon tumani"},
    1706230: {"region": "Buxoro", "district": "Qorako'l tumani"},
    1706232: {"region": "Buxoro", "district": "Qorovulbozor tumani"},
    1706240: {"region": "Buxoro", "district": "Peshku tumani"},
    1706242: {"region": "Buxoro", "district": "Romitan tumani"},
    1706246: {"region": "Buxoro", "district": "Jondor tumani"},
    1706258: {"region": "Buxoro", "district": "Shofirkon tumani"},

    # Jizzax
    1708401: {"region": "Jizzax", "district": "Jizzax shahri"},
    1708201: {"region": "Jizzax", "district": "Arnasoy tumani"},
    1708204: {"region": "Jizzax", "district": "Baxmal tumani"},
    1708209: {"region": "Jizzax", "district": "G'allaorol tumani"},
    1708212: {"region": "Jizzax", "district": "Sh.Rashidov tumani"},
    1708215: {"region": "Jizzax", "district": "Do'stlik tumani"},
    1708218: {"region": "Jizzax", "district": "Zomin tumani"},
    1708220: {"region": "Jizzax", "district": "Zarbdor tumani"},
    1708223: {"region": "Jizzax", "district": "Mirzacho'l tumani"},
    1708225: {"region": "Jizzax", "district": "Zafarobod tumani"},
    1708228: {"region": "Jizzax", "district": "Paxtakor tumani"},
    1708235: {"region": "Jizzax", "district": "Forish tumani"},
    1708237: {"region": "Jizzax", "district": "Yangiobod tumani"},

    # Qashqadaryo
    1710401: {"region": "Qashqadaryo", "district": "Qarshi shahri"},
    1710405: {"region": "Qashqadaryo", "district": "Shahrisabz shahri"},
    1710207: {"region": "Qashqadaryo", "district": "G'uzor tumani"},
    1710212: {"region": "Qashqadaryo", "district": "Dehqonobod tumani"},
    1710220: {"region": "Qashqadaryo", "district": "Qamashi tumani"},
    1710224: {"region": "Qashqadaryo", "district": "Qarshi tumani"},
    1710229: {"region": "Qashqadaryo", "district": "Koson tumani"},
    1710232: {"region": "Qashqadaryo", "district": "Kitob tumani"},
    1710233: {"region": "Qashqadaryo", "district": "Mirishkor tumani"},
    1710234: {"region": "Qashqadaryo", "district": "Muborak tumani"},
    1710235: {"region": "Qashqadaryo", "district": "Nishon tumani"},
    1710237: {"region": "Qashqadaryo", "district": "Kasbi tumani"},
    1710240: {"region": "Qashqadaryo", "district": "Ko'kdala tumani"},
    1710242: {"region": "Qashqadaryo", "district": "Chiroqchi tumani"},
    1710245: {"region": "Qashqadaryo", "district": "Shahrisabz tumani"},
    1710250: {"region": "Qashqadaryo", "district": "Yakkabog' tumani"},

    # Navoiy
    1712401: {"region": "Navoiy", "district": "Navoiy shahri"},
    1712408: {"region": "Navoiy", "district": "Zarafshon shahri"},
    1712412: {"region": "Navoiy", "district": "G'ozg'on shahri"},
    1712211: {"region": "Navoiy", "district": "Konimex tumani"},
    1712216: {"region": "Navoiy", "district": "Qiziltepa tumani"},
    1712230: {"region": "Navoiy", "district": "Navbahor tumani"},
    1712234: {"region": "Navoiy", "district": "Karmana tumani"},
    1712238: {"region": "Navoiy", "district": "Nurota tumani"},
    1712244: {"region": "Navoiy", "district": "Tomdi tumani"},
    1712248: {"region": "Navoiy", "district": "Uchquduq tumani"},
    1712251: {"region": "Navoiy", "district": "Xatirchi tumani"},

    # Namangan
    1714401: {"region": "Namangan", "district": "Namangan shahri"},
    1714204: {"region": "Namangan", "district": "Mingbuloq tumani"},
    1714207: {"region": "Namangan", "district": "Kosonsoy tumani"},
    1714212: {"region": "Namangan", "district": "Namangan tumani"},
    1714216: {"region": "Namangan", "district": "Norin tumani"},
    1714219: {"region": "Namangan", "district": "Pop tumani"},
    1714224: {"region": "Namangan", "district": "To'raqo'rg'on tumani"},
    1714229: {"region": "Namangan", "district": "Uychi tumani"},
    1714234: {"region": "Namangan", "district": "Uchqo'rg'on tumani"},
    1714236: {"region": "Namangan", "district": "Chortoq tumani"},
    1714237: {"region": "Namangan", "district": "Chust tumani"},
    1714242: {"region": "Namangan", "district": "Yangiqo'rg'on tumani"},

    # Samarqand
    1718401: {"region": "Samarqand", "district": "Samarqand shahri"},
    1718406: {"region": "Samarqand", "district": "Kattaqo'rg'on shahri"},
    1718203: {"region": "Samarqand", "district": "Oqdaryo tumani"},
    1718206: {"region": "Samarqand", "district": "Bulung'ur tumani"},
    1718209: {"region": "Samarqand", "district": "Jomboy tumani"},
    1718212: {"region": "Samarqand", "district": "Ishtixon tumani"},
    1718215: {"region": "Samarqand", "district": "Kattaqo'rg'on tumani"},
    1718216: {"region": "Samarqand", "district": "Qo'shrabot tumani"},
    1718218: {"region": "Samarqand", "district": "Narpay tumani"},
    1718224: {"region": "Samarqand", "district": "Payariq tumani"},
    1718227: {"region": "Samarqand", "district": "Pastdarg'om tumani"},
    1718230: {"region": "Samarqand", "district": "Paxtachi tumani"},
    1718233: {"region": "Samarqand", "district": "Samarqand tumani"},
    1718235: {"region": "Samarqand", "district": "Nurobod tumani"},
    1718236: {"region": "Samarqand", "district": "Urgut tumani"},
    1718238: {"region": "Samarqand", "district": "Toyloq tumani"},

    # Surxondaryo
    1722401: {"region": "Surxondaryo", "district": "Termiz shahri"},
    1722201: {"region": "Surxondaryo", "district": "Oltinsoy tumani"},
    1722202: {"region": "Surxondaryo", "district": "Angor tumani"},
    1722203: {"region": "Surxondaryo", "district": "Bandixon tumani"},
    1722204: {"region": "Surxondaryo", "district": "Boysun tumani"},
    1722207: {"region": "Surxondaryo", "district": "Muzrabot tumani"},
    1722210: {"region": "Surxondaryo", "district": "Denov tumani"},
    1722212: {"region": "Surxondaryo", "district": "Jarqo'rg'on tumani"},
    1722214: {"region": "Surxondaryo", "district": "Qumqo'rg'on tumani"},
    1722215: {"region": "Surxondaryo", "district": "Qiziriq tumani"},
    1722217: {"region": "Surxondaryo", "district": "Sariosiyo tumani"},
    1722220: {"region": "Surxondaryo", "district": "Termiz tumani"},
    1722221: {"region": "Surxondaryo", "district": "Uzun tumani"},
    1722223: {"region": "Surxondaryo", "district": "Sherobod tumani"},
    1722226: {"region": "Surxondaryo", "district": "Sho'rchi tumani"},

    # Sirdaryo
    1724401: {"region": "Sirdaryo", "district": "Guliston shahri"},
    1724410: {"region": "Sirdaryo", "district": "Shirin shahri"},
    1724413: {"region": "Sirdaryo", "district": "Yangier shahri"},
    1724206: {"region": "Sirdaryo", "district": "Oqoltin tumani"},
    1724212: {"region": "Sirdaryo", "district": "Boyovut tumani"},
    1724216: {"region": "Sirdaryo", "district": "Sayxunobod tumani"},
    1724220: {"region": "Sirdaryo", "district": "Guliston tumani"},
    1724226: {"region": "Sirdaryo", "district": "Sardoba tumani"},
    1724228: {"region": "Sirdaryo", "district": "Mirzaobod tumani"},
    1724231: {"region": "Sirdaryo", "district": "Sirdaryo tumani"},
    1724235: {"region": "Sirdaryo", "district": "Xovos tumani"},

    # Toshkent
    1727401: {"region": "Toshkent", "district": "Nurafshon shahri"},
    1727404: {"region": "Toshkent", "district": "Olmaliq shahri"},
    1727407: {"region": "Toshkent", "district": "Angren shahri"},
    1727413: {"region": "Toshkent", "district": "Bekobod shahri"},
    1727415: {"region": "Toshkent", "district": "Ohangaron shahri"},
    1727419: {"region": "Toshkent", "district": "Chirchiq shahri"},
    1727424: {"region": "Toshkent", "district": "Yangiyo'l shahri"},
    1727206: {"region": "Toshkent", "district": "Oqqo'rg'on tumani"},
    1727212: {"region": "Toshkent", "district": "Ohangaron tumani"},
    1727220: {"region": "Toshkent", "district": "Bekobod tumani"},
    1727224: {"region": "Toshkent", "district": "Bo'stonliq tumani"},
    1727228: {"region": "Toshkent", "district": "Bo'ka tumani"},
    1727233: {"region": "Toshkent", "district": "Quyi Chirchiq tumani"},
    1727237: {"region": "Toshkent", "district": "Zangiota tumani"},
    1727239: {"region": "Toshkent", "district": "Yuqori Chirchiq tumani"},
    1727248: {"region": "Toshkent", "district": "Qibray tumani"},
    1727249: {"region": "Toshkent", "district": "Parkent tumani"},
    1727250: {"region": "Toshkent", "district": "Piskent tumani"},
    1727253: {"region": "Toshkent", "district": "O'rta Chirchiq tumani"},
    1727256: {"region": "Toshkent", "district": "Chinoz tumani"},
    1727259: {"region": "Toshkent", "district": "Yangiyo'l tumani"},
    1727265: {"region": "Toshkent", "district": "Toshkent tumani"},

    # Farg'ona
    1730401: {"region": "Farg'ona", "district": "Farg'ona shahri"},
    1730405: {"region": "Farg'ona", "district": "Qo'qon shahri"},
    1730408: {"region": "Farg'ona", "district": "Quvasoy shahri"},
    1730412: {"region": "Farg'ona", "district": "Marg'ilon shahri"},
    1730203: {"region": "Farg'ona", "district": "Oltiariq tumani"},
    1730206: {"region": "Farg'ona", "district": "Qo'shtepa tumani"},
    1730209: {"region": "Farg'ona", "district": "Bag'dod tumani"},
    1730212: {"region": "Farg'ona", "district": "Buvayda tumani"},
    1730215: {"region": "Farg'ona", "district": "Beshariq tumani"},
    1730218: {"region": "Farg'ona", "district": "Quva tumani"},
    1730221: {"region": "Farg'ona", "district": "Uchko'prik tumani"},
    1730224: {"region": "Farg'ona", "district": "Rishton tumani"},
    1730226: {"region": "Farg'ona", "district": "So'x tumani"},
    1730227: {"region": "Farg'ona", "district": "Toshloq tumani"},
    1730230: {"region": "Farg'ona", "district": "O'zbekiston tumani"},
    1730233: {"region": "Farg'ona", "district": "Farg'ona tumani"},
    1730236: {"region": "Farg'ona", "district": "Dang'ara tumani"},
    1730238: {"region": "Farg'ona", "district": "Furqat tumani"},
    1730242: {"region": "Farg'ona", "district": "Yozyovon tumani"},

    # Xorazm
    1733401: {"region": "Xorazm", "district": "Urganch shahri"},
    1733406: {"region": "Xorazm", "district": "Xiva shahri"},
    1733204: {"region": "Xorazm", "district": "Bog'ot tumani"},
    1733208: {"region": "Xorazm", "district": "Gurlan tumani"},
    1733212: {"region": "Xorazm", "district": "Qo'shko'pir tumani"},
    1733217: {"region": "Xorazm", "district": "Urganch tumani"},
    1733220: {"region": "Xorazm", "district": "Hazorasp tumani"},
    1733221: {"region": "Xorazm", "district": "To'roqqal'a tumani"},
    1733223: {"region": "Xorazm", "district": "Xonqa tumani"},
    1733226: {"region": "Xorazm", "district": "Xiva tumani"},
    1733230: {"region": "Xorazm", "district": "Shovot tumani"},
    1733233: {"region": "Xorazm", "district": "Yangiariq tumani"},
    1733236: {"region": "Xorazm", "district": "Yangibozor tumani"},

    # Toshkent shahri
    1726262: {"region": "Toshkent shahri", "district": "Uchtepa tumani"},
    1726264: {"region": "Toshkent shahri", "district": "Bektemir tumani"},
    1726266: {"region": "Toshkent shahri", "district": "Yunusobod tumani"},
    1726269: {"region": "Toshkent shahri", "district": "Mirzo Ulug'bek tumani"},
    1726273: {"region": "Toshkent shahri", "district": "Mirobod tumani"},
    1726277: {"region": "Toshkent shahri", "district": "Shayxontohur tumani"},
    1726280: {"region": "Toshkent shahri", "district": "Olmazor tumani"},
    1726283: {"region": "Toshkent shahri", "district": "Sergeli tumani"},
    1726287: {"region": "Toshkent shahri", "district": "Yakkasaroy tumani"},
    1726290: {"region": "Toshkent shahri", "district": "Yashnobod tumani"},
    1726292: {"region": "Toshkent shahri", "district": "Chilonzor tumani"},
}


def convert_coato_to_names(input_file, output_file):
    """
    Read Excel file, convert COATO codes to region and district names

    Parameters:
    input_file: path to input Excel file
    output_file: path to output Excel file
    """
    print(f"Current directory: {os.getcwd()}")
    print(f"Looking for file: {input_file}")
    print(f"File exists: {os.path.exists(input_file)}")

    # Read the Excel file
    df = pd.read_excel('average_gisht_costs_five_columns.xlsx')
    print(f"File read successfully! Rows: {len(df)}")

    # Check if all COATO codes belong to the same region
    regions = []
    for coato in df['COATO']:
        if coato in coato_mapping:
            regions.append(coato_mapping[coato]['region'])

    # Determine if all codes are from the same region
    unique_regions = list(set(regions))
    if len(unique_regions) == 1:
        region_name = unique_regions[0]
    else:
        region_name = "Mixed Regions"

    print(f"Region detected: {region_name}")

    # Convert COATO codes to district names
    df['COATO'] = df['COATO'].apply(
        lambda x: coato_mapping[x]['district'] if x in coato_mapping else f"Unknown ({x})"
    )

    # Save to Excel with region name in first row
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font, Alignment

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = 'Data'

    # Write region name in first row (merged across all columns)
    worksheet.append([region_name] + [''] * (len(df.columns) - 1))
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))

    # Style the region name
    worksheet['A1'].font = Font(bold=True, size=14)
    worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')

    # Write column headers in second row
    worksheet.append(df.columns.tolist())

    # Write data starting from third row
    for row in dataframe_to_rows(df, index=False, header=False):
        worksheet.append(row)

    # Save workbook
    workbook.save(output_file)

    print(f"File saved successfully!")
    print(f"Output file: {os.path.abspath(output_file)}")
    print(f"File exists: {os.path.exists(output_file)}")


# Example usage:
if __name__ == "__main__":
    # Replace with your actual file paths
    input_file = "average_gisht_costs_five_columns.xlsx"
    output_file = "mm.xlsx"

    try:
        convert_coato_to_names(input_file, output_file)
    except FileNotFoundError:
        print(f"ERROR: File '{input_file}' not found!")
        print(f"Current working directory: {os.getcwd()}")
        print("Please check the file path.")
    except Exception as e:
        print(f"ERROR occurred: {e}")
        import traceback

        traceback.print_exc()