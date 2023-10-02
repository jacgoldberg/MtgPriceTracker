import os
import tkinter as tk
import MtgSheetHandler as mtgS
from tkinter import * 


def launchSheet():
    os.system("open -a '/Applications/Microsoft Excel.app' 'MagicCardTracker.xlsx'")

cardAccess = mtgS.CardSheet("MagicCardTracker.xlsx")


win = tk.Tk()
win.geometry("700x350")

T = Text(win, height = 0, width = 20, font=("Courier", 18))
T2 = Text(win, height = 0, width = 20, font=("Courier", 18))
l = Label(win, text = "Set Name")
l.config(font =("Courier", 20))
l2 = Label(win, text = "Card Name")
l2.config(font =("Courier", 20))

l3 = Label(win, text = "Card Price Tracker")
l3.config(font =("Courier", 30))

var1 = tk.IntVar()
var2 = tk.IntVar()
var3 = tk.IntVar()

c1 = tk.Checkbutton(win, text='Foil',variable=var1, onvalue=1, offvalue=0)
c2 = tk.Checkbutton(win, text='Borderless',variable=var2, onvalue=1, offvalue=0)
c3 = tk.Checkbutton(win, text='Show Case',variable=var2, onvalue=1, offvalue=0)


 
# Create button for next text.
b1 = Button(win, text = "Enter", command = lambda : cardAccess.addCard(T.get("1.0",END), T2.get("1.0",END), var1,var2,var3))
 
# Create an Exit button.
b2 = Button(win, text = "Exit", command = win.destroy)

b3 = Button(win, text = "Launch Spread Sheet", command = lambda : launchSheet())

b4 = Button(win, text = "Update Spread Sheet", command = lambda : cardAccess.update())

l3.place(relx= 0.5, rely= 0.1, anchor="center")
l.place(relx= 0.185, rely= 0.3)
T.place(relx= 0.1, rely= 0.4)
l2.place(relx= 0.685, rely= 0.3)
T2.place(relx= 0.6, rely= 0.4)
c1.place(relx= 0.1, rely= 0.6)
c2.place(relx= 0.175, rely= 0.6)
c3.place(relx= 0.32, rely= 0.6)
b1.place(relx= 0.7, rely= 0.85, anchor="center")
b2.place(relx= 0.8, rely= 0.85, anchor="center")
b3.place(relx= 0.5, rely= 0.85, anchor="center")
b4.place(relx= 0.25, rely= 0.85, anchor="center")
 
# Insert The Fact.
# T.insert(tk.END, Fact)
 
tk.mainloop()
