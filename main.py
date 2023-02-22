# ----------INITIALIZATION----------#

from tkinter import *


import openpyxl


# ----------CREATING_THE_EXCEL_CLASS----------#


wb = openpyxl.Workbook()
ws = wb.active


# ----------FUNCTIONS_SETUP----------#


def courier_form_initiate():
    """Δημιουργία function για το button όπου και θα δίνει εντολή
    τα labels να εισέρχονται στο αρχείο excel μας."""

    data = (
        ("ΟΝΟΜΑ", "ΕΠΙΘΕΤΟ", "ΔΙΕΥΘΥΝΣΗ", "ΤΚ", "ΠΡΟΙΟΝ", "ΒΛΑΒΗ", "ΠΛΗΡΟΦΟΡΙΕΣ"),
        (name_entry.get(), last_name_entry.get(), adress_entry.get(), postal_entry.get(), product_entry.get(),
         info_entry.get())
    )

    for i in data:
        ws.append(i)

    wb.save('output.xlsx')


def clear_text():
    """Function για να καθαρίζει το παράθυρο
     για να θέσουμε νέες τιμές"""
    name_entry.delete(0, END)
    name_entry.insert(END, string="Όνομα:")
    last_name_entry.delete(0, END)
    last_name_entry.insert(END, string="Επίθετο:")
    adress_entry.delete(0, END)
    adress_entry.insert(END, string="Διεύθυνση:")
    postal_entry.delete(0, END)
    postal_entry.insert(END, string="TK:")
    product_entry.delete(0, END)
    product_entry.insert(END, string="Προϊόν:")
    info_entry.delete(0, END)
    info_entry.insert(END, string="Λοιπές πληροφορίες:")


# ----------UI_SETUP----------#


background = "#363232"
FONT_NAME = "Courier"
RED = "#E90606"
window = Tk()
window.title("PLAISIO COURIER APP")
window.config(padx=150, pady=75, bg="black")

canvas = Canvas(width=500, height=370, bg=background, highlightthickness=0)
image_img = PhotoImage(file="image.png")
canvas.create_image(250, 175, image=image_img)
canvas.grid(column=1, row=1)

Welcome_label = Label(text="COURIER-PLAISIO", fg=RED, font=(FONT_NAME, 35, "bold"))
Welcome_label.grid(column=1, row=0)

form_button = Button(text="ΦΟΡΜΑ", height=3, width=15, highlightthickness=0, command=courier_form_initiate)
form_button.place(x=100, y=360)

clear_button = Button(text="ΚΑΘΑΡΙΣΜΟΣ", height=3, width=15, highlightthickness=0, command=clear_text)
clear_button.place(x=300, y=360)

name_entry = Entry(font="Courier 12")
name_entry.insert(END, string="Όνομα:")
last_name_entry = Entry(font="Courier 12")
last_name_entry.insert(END, string="Επίθετο:")
adress_entry = Entry(font="Courier 12")
adress_entry.insert(END, string="Διεύθυνση:")
postal_entry = Entry(font="Courier 12")
postal_entry.insert(END, string="TK:")
product_entry = Entry(font="Courier 12")
product_entry.insert(END, string="Προϊόν:")
info_entry = Entry(font="Courier 12")
info_entry.insert(END, string="Λοιπές πληροφορίες:")
name_entry.place(x=155, y=110, width=200, height=35)
last_name_entry.place(x=155, y=150, width=200, height=35)
adress_entry.place(x=155, y=190, width=200, height=35)
postal_entry.place(x=155, y=230, width=200, height=35)
product_entry.place(x=155, y=270, width=200, height=35)
info_entry.place(x=155, y=310, width=200, height=35)

window.mainloop()
