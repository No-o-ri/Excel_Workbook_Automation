from utils import process_workbook_data
compound = None
interest = None
time_duration = None
compound_num = None
try:
    interest = float(input("Enter the interest rate:"))
    compound = str(input("Is this compounded monthly, annually, or weekly,:"))
    time_duration = int(input("How long are you letting this build up:"))
except ValueError or TypeError:
    print("Invalid Values Entered")


compound_dictionary = {
    "monthly": 12,
    "annually": 1,
    "weekly": 52
}
if compound is not None:
    compound_num = compound_dictionary.get(compound.lower(), 'Not valid duration')


process_workbook_data("PyAuto.xlsx", interest, compound_num, time_duration)
