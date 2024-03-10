#!/usr/bin/env python3
#Written by Michael Pistono, 2019


import pandas as pd
import os
import datetime
import xlsxwriter



def get_current_datetime():
    import datetime
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def list(pt_list):
    header = pt_list[0]
    max_lengths = [max(map(len, map(str, col))) for col in zip(*pt_list)]

    formatted_header = [" " * 3 + f"{value:{length}}" if value not in ("Employee?", "Empl. #")
                        else f"{value:^{length}}" for value, length in zip(header, max_lengths)]
    print("|".join(formatted_header))

    for i, row in enumerate(pt_list[1:], start=1):
        formatted_row = [f"{value:{length}}" if value not in ("y", "n", "Employee?")
                         else f"{value:^{length}}" for value, length in zip(row, max_lengths)]
        print(f"{i}. {' | '.join(formatted_row)}")
        print()
    input("Press Enter to return to the menu...")
    print() # Wait for user input before returning to the menu



def add(pt_list, names_list):
    name = input("Last name, First name: ").lower()

    while True:
        status = input("Employee Status (y/n): ").lower()

        if status == "y":
            employee_number = input("Employee Number: ")
            break
        elif status == "n":
            employee_number = "guest"
            break
        else:
            print("Please enter either Y or N only!")
            print()

    current_datetime = get_current_datetime()
    date, time = current_datetime.split()

    visitor_pt = [name, status, employee_number, date, time, "", ""]
    pt_list.append(visitor_pt)

    # Check if the visitor is already in names_list, update only the date_in and time_in
    visitor_names = next((v for v in names_list if v[:3] == [name, status, employee_number]), None)
    if visitor_names:
        # Check if there are available columns for additional times
        index = len(visitor_names)
        if index < 5:
            visitor_names[index:index + 2] = [date, time]
        else:
            visitor_names.extend(["", "", date, time])
    else:
        # If not found, add a new entry to names_list
        visitor_names = [name, status, employee_number, date, time, "", ""]
        names_list.append(visitor_names)

    print(f"{visitor_pt} was added.\n")
    print()

    input("Press Enter to return to the menu...")
    print()
def report(names_list):
    header = names_list[0]
    max_lengths = [max(map(len, map(str, col))) for col in zip(*names_list)]

    formatted_header = [" " * 3 + f"{value:{length}}" if value not in ("Employee?", "Empl. #")
                        else f"{value:^{length}}" for value, length in zip(header, max_lengths)]
    print("|".join(formatted_header))

    for i, row in enumerate(names_list[1:], start=1):
        formatted_row = [f"{value:{length}}" if value not in ("y", "n", "Employee?")
                         else f"{value:^{length}}" for value, length in zip(row, max_lengths)]
        print(f"{i}. {' | '.join(formatted_row)}")
        print()
    input("Press Enter to return to the menu...")# Wait for user input before returning to the menu
    print()






def delete(pt_list, names_list):
    header = pt_list[0]
    max_lengths = [max(map(len, map(str, col))) for col in zip(*pt_list)]

    formatted_header = [" " * 3 + f"{value:{length}}" if value not in ("Employee?", "Empl. #")
                        else f"{value:^{length}}" for value, length in zip(header, max_lengths)]
    print("|".join(formatted_header))

    for i, row in enumerate(pt_list[1:], start=1):
        formatted_row = [f"{value:{length}}" if value not in ("y", "n", "Employee?")
                         else f"{value:^{length}}" for value, length in zip(row, max_lengths)]
        print(f"{i}. {' | '.join(formatted_row)}")

    while True:
        number = input("Line number of departing visitor: ")
        print()
        if not number.isdigit():
            print("Invalid input. Please enter a valid number.")
            print()
        else:
            number = int(number)
            try:
                visitor = pt_list.pop(number)
                print(f"Visitor {number} ({visitor[0]}) was deleted.\n")
                print()
                print()# Print the name of the deleted visitor
                # Check if the visitor is in names_list and update the date_out and time_out
                name, status, employee_number = visitor[:3]
                for row in names_list[1:]:
                    if row[:3] == [name, status, employee_number] and row[-2:] == ["", ""]:
                        row[-2:] = get_current_datetime().split()

                break
            except IndexError:
                print("Invalid line number. Please enter a number within the range.")
                print()
    input("Press Enter to return to the menu...")
    print()








def display_menu():
    if os.name == 'nt':  # Windows
        os.system('cls')
    elif os.environ.get('TERM'):  # Unix-based system with TERM variable set
        os.system('clear')
    else:  # Fallback if TERM variable is not set
        print('\n' * 100)  # Print several newline characters to simulate clearing the screen

    print("-----------SEVEN POINT PROTECTION-----------")
    print("----------------COMMAND MENU----------------")
    print("list - (List current visitors on site)")
    print("report - (List all of today's visitors)")
    print("add - (Add a new visitor to list upon entry)")
    print("del - (Delete a visitor from list upon exit)")
    print("exit - (Finish and export daily report)")
    print("--------------------------------------------")
    print()
    print()
    print()
    print()




def save_to_excel(names_list):
    try:
        # Create a DataFrame from names_list
        df = pd.DataFrame(names_list[1:], columns=names_list[0])

        # Get the user's desktop directory
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

        # Construct the full path for the Excel file
        excel_file_path = os.path.join(desktop_path, "employee_report.xlsx")

        # Create a Pandas Excel writer using xlsxwriter as the engine
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book

            # Create a worksheet named 'Daily Visitors'
            worksheet = workbook.add_worksheet('Daily Visitors')

            # Define a title format with centered alignment, bold, and font size 16
            title_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 16})

            # Define an underline format for column names
            underline_format = workbook.add_format({'underline': True})

            worksheet.set_column('A:B', 33)
            worksheet.set_column('B:G', 14)

            # Merge cells for the title row and enter the text
            title_text = f"20200 Spence Rd | Personnel On Site - Date: {get_current_datetime().split()[0]}"
            worksheet.merge_range('A1:G1', title_text, title_format)

            # Write column headers starting from A4 with underlined format
            for col_num, value in enumerate(df.columns.values, start=1):
                worksheet.write(3, col_num - 1, value, underline_format)

            # Set column names starting from A4
            df.to_excel(writer, sheet_name='Daily Visitors', startrow=4, header=False, index=False)

            # Add a box for 'Total Visitors' in A2
            total_visitors_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
            worksheet.write('A2', 'Total Visitors:', total_visitors_format)
            
            # Calculate and write the total number of visitors in B2
            total_visitors = len(df)
            worksheet.write('B2', total_visitors, total_visitors_format)

        print(f"Employee report saved to {excel_file_path}")
    except Exception as e:
        print(f"An error occurred while saving the report: {e}")





def main():
    pt_list = [["Visitor's Name (last, first)", "Employee Status", "Employee #", "Date In", "Time In"]]
    names_list = [["Visitor's Name (last, first)", "Employee Status", "Employee #", "Date In", "Time In", "Date Out", "Time Out"]]

    while True:
        display_menu()  # Display menu at the beginning of each loop iteration

        try:
            command = input("Command:  ")
            print()
            print()
            if command.lower() == "list":
                list(pt_list)
                print()
            elif command.lower() == "add":
                add(pt_list, names_list)
                print()
            elif command.lower() == "del":
                delete(pt_list, names_list)
                print()
            elif command.lower() == "report":
                report(names_list)
                print()
            elif command.lower() == "exit":
                save_to_excel(names_list)
                print()
                break
            else:
                print("Invalid command. Try again!")
                print()

        except Exception as e:
            print(f"An error occurred: {e}")
            save_to_excel(names_list)  

if __name__ == "__main__":
    main()





