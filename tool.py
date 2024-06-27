from tkinter import *



def submitButton(button):
    

    return




root = Tk()
w = Label(
    root, 
    text = "Accounting Tool",
    font = ("Helvetica", 24 ))
w.pack()

Dinput = Entry(root)
Dinput.pack()

DButton = Button(root, text = "submit directory", command = submitButton(Dinput))
DButton.pack()


root.geometry('500x500')
root.mainloop()
