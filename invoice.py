import tkinter as tk
from tkinter import ttk,messagebox
import datetime
import sys, os
from docxtpl import DocxTemplate

# def resource_path(relative_path):
#     try:
#         base_path = sys._MEIPASS  # When running from .exe
#     except:
#         base_path = os.path.abspath(".")  # When running from .py
#     return os.path.join(base_path, relative_path)



window = tk.Tk()
window.title("Billing App")
window.geometry("900x600")

cName = tk.StringVar()
address = tk.StringVar()
phone = tk.StringVar()
email = tk.StringVar()

pName = tk.StringVar()
qty = tk.IntVar()
price = tk.StringVar()

listData = []

def addProduct():
    table.insert("", "end", values=(pName.get(),qty.get(),price.get(),int(qty.get())*float(price.get())))
    listData.append([pName.get(),qty.get(),price.get(),int(qty.get())*float(price.get())])
    pName.set("")
    qty.set(0)
    price.set("")

def generate():
    # doc = DocxTemplate(resource_path("pkg_file.docx"))
    doc = DocxTemplate("pkg_file.docx")
    total = sum(item[3] for item in listData)
    doc.render({'cName':cName.get(),'address':address.get(),'email':email.get(),'phone':phone.get(),'itemList':listData,'total':total})
    x = str(datetime.datetime.now()).replace(" ", "").replace("-", "").replace(":", "").replace(".", "")
    fileName = str(cName.get()).replace(" ","") + x[10:]
    doc.save(f"{fileName}.docx")
    cName.set("")
    email.set("")
    address.set("")
    phone.set("")
    messagebox.showinfo("Invoice Generated",f"{fileName} is saved")
    for item in table.get_children():
        table.delete(item)



window.rowconfigure(0,weight=1)
window.rowconfigure(1,weight=1)
window.rowconfigure(2,weight=1)
window.rowconfigure(3,weight=1)
window.rowconfigure(4,weight=1)
window.rowconfigure(5,weight=1)
window.rowconfigure(6,weight=10)
window.rowconfigure(7,weight=1)


window.columnconfigure(0,weight=1)
window.columnconfigure(1,weight=1)
window.columnconfigure(2,weight=1)
window.columnconfigure(3,weight=1)

l1 = ttk.Label(master=window,text="Customer Name").grid(row=1,column=0)
l2 = ttk.Label(master=window,text="Address").grid(row=1,column=1)
l3 = ttk.Label(master=window,text="Phone").grid(row=1,column=2)
l4 = ttk.Label(master=window,text="Email").grid(row=1,column=3)

e1 = ttk.Entry(master=window,textvariable=cName).grid(row=2,column=0)
e2 = ttk.Entry(master=window,textvariable=address).grid(row=2,column=1)
e3 = ttk.Entry(master=window,textvariable=phone).grid(row=2,column=2)
e4 = ttk.Entry(master=window,textvariable=email).grid(row=2,column=3)


l5 = ttk.Label(master=window,text="Product Name").grid(row=4,column=0)
l6 = ttk.Label(master=window,text="Quantity").grid(row=4,column=1)
l7 = ttk.Label(master=window,text="Price").grid(row=4,column=2)

e5 = ttk.Entry(master=window,textvariable=pName).grid(row=5,column=0)
e6 = ttk.Entry(master=window,textvariable=qty).grid(row=5,column=1)
e7 = ttk.Entry(master=window,textvariable=price).grid(row=5,column=2)

b1 = ttk.Button(master=window,text="Add Product",command=addProduct).grid(row = 5, column=3,padx=10,sticky="ew")
table = ttk.Treeview(master=window,columns=("name","quantity","price","total"),show="headings")
table.grid(row=6,column=0,sticky="ewns",columnspan=4,padx=20)

table.heading("name",text="Product Name")
table.heading("quantity",text="Quantity")
table.heading("price",text="Price")
table.heading("total",text="Total")


b2 = ttk.Button(master=window,text="Generate Invoice",command=generate).grid(row=7,column=3,sticky="ew",padx=10)





window.mainloop()