import tkinter as tk
import os
import sys
from tkinter import *
from PIL import ImageTk,Image
from tkinter import messagebox
from tkinter.filedialog import askopenfile, asksaveasfile
from converter import *
from xls_xlsx_csv import *
import argparse
import webbrowser


file_in, file_out = None, None


def open_file():
    """
    ask user for filepath of the .csv
    :return: filepath
    """
    global file_in
    filename =""
    if option_var.get() == 3:
        filename = askopenfile(filetypes=[("CSV files", "*.csv")])
    else :
        filename = askopenfile(filetypes=[("Excel files", "*.xlsx")])
    if filename:
        file_in = filename.name
        path_sv.set(file_in)


def output_location():
    """
    ask user where to store the .vcf
    :return: filepath
    """
    global file_out
    filename = asksaveasfile(mode="w", defaultextension=".vcf",filetypes=[("VCF files", "*.vcf")])
    if filename:
        file_out = filename.name
        pathOut_sv.set(file_out)

def callback(url):
    webbrowser.open_new(url)

def convert():
    if not file_in or not file_out:
        messagebox.showwarning("ERROR", "The following File is missing {}".format(
            "Input-CSV-File" if not file_in else "Output-CSV-File"))
        exit()

    c = Converter(file_in, file_out)
    success = c.convert()

    if success[0]:
        messagebox.showinfo("Conversion successful", "Successfully converted the CSV!")
    else:
        messagebox.showwarning("ERROR",
                               "Conversion aborted, because of the following error: {error}".format(error=success[1]))
    root.destroy()
def convert2():
    if not file_in:
        messagebox.showwarning("ERROR", "The following File is missing {}".format(
            "Input-xlsx-File" if not file_in else "Output-CSV-File"))
        exit()

    c =  xlsxToCsvHandler(file_in)
    root.destroy()
def convert3():
    if not file_in or not file_out:
        messagebox.showwarning("ERROR", "The following File is missing {}".format(
            "Input-xlsx-File" if not file_in else "Output-vcf-File"))
        exit()
    xlsxToCsvHandler(file_in)
    csv_filename = file_in.rsplit('/', 1)[-1].rsplit('.', 1)[0]
    xlsFilePath = file_in.rsplit('/', 1)[0].rsplit('.', 1)[0]
    csvFilePathName = xlsFilePath +"/"+csv_filename +".csv"
    c = Converter(csvFilePathName, file_out)
    success = c.convert()

    if success[0]:
        messagebox.showinfo("Conversion successful", "Successfully converted the CSV!")
    else:
        messagebox.showwarning("ERROR",
                               "Conversion aborted, because of the following error: {error}".format(error=success[1]))
    root.destroy()
if __name__ == "__main__":
    # switch modes
    # console
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description="Convert CSV to VCF")
        parser.add_argument('-in', dest='file', help='Input File Name', required=True)
        parser.add_argument('-out', dest='out', help='Output Path for VCF', default='contacts.vcf')
        args = parser.parse_args()

        file_in = args.file
        file_out = args.out

        c = Converter(file_in, file_out)
        c.convert()


    # gui
    else:
        root = tk.Tk()
        root.title("Excel To VCF converter")
        root.iconbitmap('images/vcf.ico')
        root.geometry("800x600")
        #frame = tk.Frame(root).pack()
        background_image = tk.PhotoImage(file="./images/background.png")
        background_label = tk.Label(root, image=background_image)
        background_label.photo = background_image
        background_label.place(x=0, y=0, relwidth=1, relheight=1)
        # logo
        logo_img=tk.PhotoImage(file="./images/idbees-logo5.png")
        logo_label=tk.Label(root,image=logo_img)
        logo_label.place(x=632,y=25)
        #hyperlink
        link1 = Label(root, text="http://www.idbees7.entej.com", fg="blue", cursor="hand2")
        link1.place(x=600,y=130)
        link1.bind("<Button-1>", lambda e: callback("http://www.idbees7.entej.com/idbees"))

        # background_image = tk.PhotoImage(file="registrationForm.png")
        # panel1 = Label(self, image=background_image, height=500, width=500)
        # panel1.pack()
        # panel1.image = background_image

        label_1 = tk.Label(root, text="Excel to VCF-Converter", font=("arial", 30, "bold"),bg="#daf5ec").place(x=80, y=50)

        path_sv = tk.StringVar()
        pathOut_sv = tk.StringVar()
        label_Frame3 = tk.LabelFrame(root, text="CSV to VCF", highlightthickness=1,
                                    width=750, height=200, bd=1, padx=0, pady=0, bg="#daf5ec")

        path_Entry = tk.Entry(label_Frame3, bd=1, justify=tk.CENTER, bg='#FFFFFF', width=65,
                              state='disabled', relief=tk.SOLID, textvariable=path_sv,font=("arial", 12, "normal")).place(x=10, y=20)
        button1 = tk.Button(label_Frame3, text="Input File", command=open_file,font=("arial", 12, "normal")).place(x=630, y=15)

        pathOut_Entry = tk.Entry(label_Frame3, bd=1, justify=tk.CENTER, bg='#FFFFFF', width=65,
                                 state='disabled', relief=tk.SOLID, textvariable=pathOut_sv,font=("arial", 12, "normal")).place(x=10, y=80)
        button2 = tk.Button(label_Frame3, text="Output File", command=output_location,font=("arial", 12, "normal")).place(x=630, y=75)
        button3 = tk.Button(label_Frame3, text="Convert!", font=("arial", 12, "bold"), command=convert).place(x=320,y=120)

        # *****************************************************************************************************************
        label_Frame2 = tk.LabelFrame(root, text="Excel to CSV utf 8", highlightthickness=1,
                                    width=750, height=140, bd=1, padx=0, pady=0,bg="#daf5ec")
        path_Entry2 = tk.Entry(label_Frame2, bd=1, justify=tk.CENTER, bg='#FFFFFF', width=65,
                              state='disabled', relief=tk.SOLID, textvariable=path_sv,font=("arial", 12, "normal")).place(x=10, y=20)
        button4 = tk.Button(label_Frame2, text="Input File", command=open_file,font=("arial", 12, "normal")).place(x=630, y=15)
        button5 = tk.Button(label_Frame2, text="Convert!", font=("arial", 12, "bold"), command=convert2).place(x=320,
                                                                                                              y=80)
        #***********************************************************************************************************************
        label_Frame = tk.LabelFrame(root, text="Excel to VCF", highlightthickness=1,
                                     width=750, height=200, bd=1, padx=0, pady=0,bg="#daf5ec")

        path_Entry = tk.Entry(label_Frame, bd=1, justify=tk.CENTER, bg='#FFFFFF', width=65,
                              state='disabled', relief=tk.SOLID, textvariable=path_sv,font=("arial", 12, "normal")).place(x=10, y=20)
        button6 = tk.Button(label_Frame, text="Input File", command=open_file,font=("arial", 12, "normal")).place(x=630, y=15)

        pathOut_Entry = tk.Entry(label_Frame, bd=1, justify=tk.CENTER, bg='#FFFFFF', width=65,
                                 state='disabled', relief=tk.SOLID, textvariable=pathOut_sv,font=("arial", 12, "normal")).place(x=10, y=80)
        button7 = tk.Button(label_Frame, text="Output File", command=output_location,font=("arial", 12, "normal")).place(x=630, y=75)
        button8 = tk.Button(label_Frame, text="Convert!", font=("arial", 12, "bold"), command=convert3).place(x=320,y=130)

        label_Frame.place(x=30, y=300)
        def option_select():
            if option_var.get() == 3:
                label_Frame.place_forget()
                label_Frame2.place_forget()
                label_Frame3.place(x=30, y=300)

            elif option_var.get() == 2:
                label_Frame.place_forget()
                label_Frame3.place_forget()
                label_Frame2.place(x=30, y=300)
            elif option_var.get() == 1:
                label_Frame2.place_forget()
                label_Frame3.place_forget()
                label_Frame.place(x=30, y=300)


        option_var = tk.IntVar()
        option_var.set(1)
        rb1 = tk.Radiobutton(root, text='Excel to VCF', variable=option_var, value=1,
                               command=option_select,bg="#daf5ec" , font=("arial", 12, "normal")).place(x=100,y=200)
        rb2 = tk.Radiobutton(root, text='Excel to CSV utf 8', variable=option_var, value=2,
                               command=option_select,bg="#daf5ec",font=("arial", 12, "normal")).place(x=250,y=200)

        rb3 = tk.Radiobutton(root, text='CSV to VCF', variable=option_var, value=3,
                             command=option_select,bg="#daf5ec",font=("arial", 12, "normal")).place(x=440,y=200)





        root.mainloop()
