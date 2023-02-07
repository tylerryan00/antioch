from openpyxl import Workbook

help_description = """
Random Groups

This script has the goal of randomizing given groups
"""

def randomizer():
    pass

def workbook_setup(wb, day, names, van_size):
    worksheet_name = 'Day ' + str(day)
    ws = wb.create_sheet(worksheet_name)
    y=0
    
    for x in range(len(names)):
        if x % (van_size/2) == 0:
            y += 1
        if x % van_size == 0:
            y += 2
        
        ws.cell(row=(x % (van_size/2))+1, column=y, value=names[x])

    


def main():

    # test cases
    days_of_trip = 5
    van_size = 12
    with open('names.txt') as f:
        names = f.readlines()

    workbook = Workbook()
    ws = workbook.active
    ws.title = 'People'

    for x in range(len(names)):
        ws.cell(row=x+1, column=1, value=names[x])

    
    for day in range(1, days_of_trip+1):
        workbook_setup(workbook, day, names, van_size)
    
    workbook.save('test.xlsx')
    

if __name__ == "__main__":
    main()
