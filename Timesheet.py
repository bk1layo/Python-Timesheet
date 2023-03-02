import tkinter as tk
from tkinter import messagebox as tkMB
import datetime
import pandas as pd
import os

# Main window
window = tk.Tk()
window.geometry('250x120')

# Variables
currentTD = datetime.datetime.now()
currentDate = datetime.date.today()
currentTime = currentTD.strftime('%H:%M:%S')

# Realtime Clock function
def realtimeClock():
    rawTD = datetime.datetime.now()
    nowTime = rawTD.strftime('%H:%M:%S %p')
    time.config(text = nowTime)
    window.after(1000, realtimeClock)

# Excel File Function
def create_excel_file():
    if not os.path.exists('TimeSheet.xlsx'):
        data = {}
        timeSheet = pd.DataFrame(data)
        timeSheet.to_excel('TimeSheet.xlsx', sheet_name='Sheet1', index=False)
    else:
        timeSheet = pd.read_excel('TimeSheet.xlsx')
        if 'HoursWorked' not in timeSheet.columns:
            timeSheet['HoursWorked'] = ''
            timeSheet.to_excel('TimeSheet.xlsx', index=False)

# Calling the function  
create_excel_file()

# Button Function
def clockInBut():
    # Open the excel file and read the data
    openTs = pd.read_excel('TimeSheet.xlsx')
    
    # Create a new row with the current date and time
    new_row = {'Date':currentDate, 'ClockIn':time.cget('text'), 'ClockOut':' '}
    
    # Append the new row to the dataframe
    openTs = openTs.append(new_row, ignore_index=True)
    
    # Write the dataframe to the excel file
    openTs.to_excel('TimeSheet.xlsx', index=False)
    
    # Notification when Clocked In
    tkMB.showinfo("Clocked in at: ", currentDate.strftime("%A, %b %d, %Y")+" "+time.cget("text"))

def clockOutBut():
    # Open the excel file and read the data
    openTs = pd.read_excel('TimeSheet.xlsx')
    
    # Adding into ClockOut in the same row
    last_clock_in_index = openTs[openTs['ClockIn'] != ' '].index[-1]
    openTs.at[last_clock_in_index, 'ClockOut'] = time.cget('text')
    
    # Calculate hours worked and update the dataframe
    clock_in = datetime.datetime.strptime(openTs.at[last_clock_in_index, 'ClockIn'], '%H:%M:%S %p')
    clock_out = datetime.datetime.strptime(time.cget('text'), '%H:%M:%S %p')
    hours_worked = (clock_out - clock_in).total_seconds() / 3600
    openTs.at[last_clock_in_index, 'HoursWorked'] = round(hours_worked, 2)
    
    openTs.to_excel('TimeSheet.xlsx', index=False)
    
    # Notification when Clocked Out
    tkMB.showinfo("Clocked out at: ", currentDate.strftime("%A, %b %d, %Y")+" "+time.cget("text"))

def openTimeSheet():
    # Open the excel file and read the data
    openTs = pd.read_excel('TimeSheet.xlsx')
    
    # Printing the data file into terminal
    print(openTs)
    
# Window Title
window.title('TimeSheet')

# Buttons and Labels
blank = tk.Label(window)
date = tk.Label(window, text= currentDate.strftime("%A, %b %d, %Y"))
time = tk.Label(window, text= currentTD.strftime('%H:%M:%S %p'))
clockIn = tk.Button(window, text='Clock In', command = clockInBut)
clockOut = tk.Button(window, text='Clock Out', command = clockOutBut)
timeSheetOpen = tk.Button(window, text='Open TimeSheet', command = openTimeSheet)

# pack() method to put into window
clockIn.pack()
clockOut.pack()
timeSheetOpen.pack()
date.pack()
time.pack()

# Call realtimeClock() function
realtimeClock()

# mainloop() method to display window
window.mainloop()
