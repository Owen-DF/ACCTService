from tkinter import *



def submitButtonD():


    

    return




root = Tk()
w = Label(
    root, 
    text = "Accounting Tool",
    font = ("Helvetica", 24 ))
w.pack()

Dinput = Entry(root)
Dinput.pack()

DButton = Button(root, text = "submit directory", command = submitButtonD())
DButton.pack()

Sinput = Text(root, height=5)
Sinput.pack()

Sinput.insert('1.0', 'Enter each search term seperated by a coma')

SButton = Button(root, text = "submit search terms")
SButton.pack()


root.geometry('500x500')
root.mainloop()
