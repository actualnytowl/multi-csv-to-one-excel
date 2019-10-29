import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from os import listdir
from os.path import isfile, join
import time
import threading
no = 1
export = None

root = tk.Tk()
root.title("Chaithu's cheat tool")
dir = tk.Text(root, width =40, height =1)
dir.pack()
dir.insert("1.0","paste base directory path here")
canvas1 = tk.Canvas(root, width=300, height=300, relief='raised')
canvas1.pack()


label1 = tk.Label(root, text='File Conversion Tool')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)

label2 = tk.Label(root, text='STATUS : Yet to start')
label2.config(font=('helvetica', 12))
canvas1.create_window(150, 180, window=label2)



time1 = ''
clock = tk.Label(root, font=('times', 20, 'bold'))
clock.pack()
def tick():
    global time1
    time2 = time.strftime('%I:%M:%S')
    if time2 != time1:
        time1 = time2
        clock.config(text=time2)
    clock.after(500, tick)



def getCSVs():
    global t1
    t1 = threading.Thread(target=getCSV)
    t1.start()
def getCSV():
    global read_file,mypath,no,export
    mypath = dir.get("1.0", 'end-1c')
    try:
        csvfiles = [f for f in listdir(mypath) if isfile(join(mypath, f)) and f.endswith('.csv')]
    except FileNotFoundError:
        label2.config(text='STATUS : give a valid path and try again')
    else:
        label2.config(text="STATUS : Processing")
        for i in range(len(csvfiles)):
            import_file_path = mypath+"\\"+csvfiles[i]    #filedialog.askopenfilename()
            read_file = pd.read_csv(import_file_path, delimiter="|")
            convertToExcel(read_file)
        no = 1
        export = None
        if(len(csvfiles)==0):
            label2.config(text="STATUS : no csv files found in path")
        else:
            label2.config(text="STATUS : Finished")
        wrt.close()


browseButton_CSV = tk.Button(text="      Save Excel to..     ", command=getCSVs, bg='green', fg='white',
                             font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=browseButton_CSV)

def convertToExcel(read_file):
    global no
    global export
    global wrt
    if no == 1:
        export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
        export = export_file_path
        wrt = pd.ExcelWriter(export)

    sheet = "sheet"+str(no)

    read_file.to_excel(wrt, sheet_name=sheet, index=None, header=True)
    no += 1


def exitApplication():

    MsgBox = tk.messagebox.askquestion('Exit Application', 'Are you sure you want to exit the application',
                                       icon='warning')
    if MsgBox == 'yes':
        try:
            wrt.close()
            root.destroy()
            exit()
        except:
            root.destroy()
            exit()


exitButton = tk.Button(root, text='       Exit Application     ', command=exitApplication, bg='brown', fg='white',
                       font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 230, window=exitButton)
tick()


root.mainloop()