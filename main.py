import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl

def Submit_Details():
    #accepted = accept_var.get()

    #if accepted=="Accepted":
        # user info
        invoice = invoice_number_entry.get()
        collection = collection_combobox.get()
        delivery = delivery_combobox.get()

        title = title_combobox.get()
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()

        if title and firstname and lastname:
            telephone = phone_entry.get()

            # order deatils
            print("InvoiceNumber: ",invoice,"Collection Mode: ",collection,"Delivery: ",delivery,"Title: ",title,"FirstName: ",firstname, "LastName: ",lastname,"Tel Phone: ",telephone)

            filepath = "D:\niki\py\orderData.xlsx"

            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["InvoiceNumber","Collection Mode","Delivery Mode","Title","FirstName","LastName","Telephone"]
                sheet.append(heading)
                workbook.save(filepath)
                workbook=openpyxl.load_workbook(filepath)
                sheet = workbook.active
                sheet.append(invoice,collection,delivery,title,firstname,lastname,telephone)
                workbook.save(filepath)
            else:
                tkinter,messagebox.showwarning(title="Error",message="Title, First Name, Last Name is not updated")
        #else:
            #tkinter,messagebox.showwarning(title="Error",message="You have not accepted the terms")



window = tkinter.Tk()
window.title("mepro - Plan-B Laundry")

frame = tkinter.Frame(window)
frame.pack()

#save trans
trans_info_frame = tkinter.LabelFrame(frame,text="Transaction Info")
trans_info_frame.grid(row=0,column=0,sticky="news",padx=20,pady=20)

invoice_number_lable = tkinter.Label(trans_info_frame,text="Invoice Number")
invoice_number_lable.grid(row=0,column=0)
invoice_number_entry = tkinter.Entry(trans_info_frame)
invoice_number_entry.grid(row=1,column=0)

collection_lable = tkinter.Label(trans_info_frame,text="Collection Mode")
collection_lable.grid(row=0,column=1)
collection_combobox = ttk.Combobox(trans_info_frame,values=["Home Collection","Self Drop"])
collection_combobox.grid(row=1,column=1)

delivery_lable = tkinter.Label(trans_info_frame,text="Delivery Mode")
delivery_lable.grid(row=0,column=2)
delivery_combobox = ttk.Combobox(trans_info_frame,values=["Home Delivery","Self Pick-Up"])
delivery_combobox.grid(row=1,column=2)

#save info
user_info_frame = tkinter.LabelFrame(frame,text="Customer Info")
user_info_frame.grid(row=1,column=0,sticky="news",padx=20,pady=20)

title_lable = tkinter.Label(user_info_frame,text="Title")
title_lable.grid(row=0,column=0)
title_combobox = ttk.Combobox(user_info_frame,values=["Mr.","Ms.","Dr."])
title_combobox.grid(row=1,column=0)


first_name_lable = tkinter.Label(user_info_frame,text="First Name")
first_name_lable.grid(row=0,column=1)
first_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1,column=1)


last_name_lable = tkinter.Label(user_info_frame,text="Last Name")
last_name_lable.grid(row=0,column=2)
last_name_entry = tkinter.Entry(user_info_frame)
last_name_entry.grid(row=1,column=2)

phone_lable = tkinter.Label(user_info_frame,text="Tel Phone")
phone_lable.grid(row=0,column=3)
phone_entry = tkinter.Entry(user_info_frame)
phone_entry.grid(row=1,column=3)

# save bill
billing_info_frame = tkinter.LabelFrame(frame,text="Billing Info")
billing_info_frame.grid(row=2,column=0,padx=20,pady=20)

# save payment
payment_info_frame = tkinter.LabelFrame(frame,text="Payment Info")
payment_info_frame.grid(row=3,column=0,padx=20,pady=20)

# button
button = tkinter.Button(frame,text="Submit", command="Submit_Details")
button.grid(row=3,column=0,sticky="news",padx=20,pady=20)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)
for widget in trans_info_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)

window.mainloop()






