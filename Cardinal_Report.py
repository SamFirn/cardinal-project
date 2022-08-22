
import pandas as pd
import matplotlib.pyplot as plt

#{"c" : "Cleaned up"}

file = 'RejectSqFtReport.xlsx'

df = pd.read_excel(file)

new_columns = ['DateTime', 'Shift ', 'Order Location ID', 'Location ID','Order Number', 'Sched ID', 'Unit ID', 'Master Key', 'Parent Key','Part No', 'Height', 'Width', 'Sq Ft', 'NaN', 'Employee', 'Code', 'Station ID']

df.columns = new_columns

rdf = df.drop([0])

cdf = rdf.drop(columns=['DateTime', 'Order Location ID', 'Location ID','Order Number', 'Sched ID', 'Unit ID', 'Master Key', 'Parent Key','Part No', 'Height', 'Width', 'Sq Ft', 'NaN'])


shift_remakes = cdf.groupby(['Shift ']).count()
csr = shift_remakes.drop(columns=['Employee', 'Station ID'])
csr["Remakes"] = csr['Code']
ccsr = csr.drop(columns=['Code'])
#ccsr.to_excel('by_shift.xlsx')


employee_remakes = cdf.groupby(['Employee']).count()
cer = employee_remakes.drop(columns=['Shift ', 'Station ID'])
cer['Remakes'] = cer['Code']
ccer = cer.drop(columns=['Code'])
#ccer.to_excel('by_employee.xlsx')

station_remakes = cdf.groupby(['Station ID']).count()
cstr = station_remakes.drop(columns=['Shift ', 'Employee'])
cstr['Remakes'] = cstr['Code']
ccstr = cstr.drop(columns=['Code'])
#ccstr.to_excel('by_station.xlsx')


code_remakes = cdf.groupby(["Code"]).count()
cr = code_remakes.drop(columns=['Shift ', 'Station ID'])
cr['Remakes'] = cr['Employee']
ccr = cr.drop(columns=['Employee'])
#ccr.to_excel('by_code.xlxs')

def new_excel():
    make = input("Make new excels? \n y/n:")
    return make

new = new_excel()


def report():
    
    report_choice = input("report type:")
    
    if report_choice == 'code':
        report_choice = ccr
    elif report_choice == 'station':
        report_choice = ccstr
    elif report_choice == 'employee':
        report_choice = ccer
    elif report_choice =='shift':
        report_choice = ccsr
    
    try:
        report_choice.plot(kind='bar')
        plt.show()
    except AttributeError:
        print("Please check spelling of report type")

if new == "y" or new == "Y":
    ccr.to_excel('by_code.xlsx')
    ccstr.to_excel('by_station.xlsx')
    ccer.to_excel('by_employee.xlsx')
    ccsr.to_excel('by_shift.xlsx')
    print('Please check current directorty for new your files')

report()
