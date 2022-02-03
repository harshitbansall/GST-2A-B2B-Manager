import os,sqlite3,shutil
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import pandas,xlrd,csv
userProfile = os.path.expanduser("~").split("\\")[-1]
class XLXFile:
    def __init__(self, path):
        self.path = path
        self.name = path.split("/")[-1].split(".")[0]
        self.ext = path.split(".")[1]

def browseFile(fileType):
    selectedFile = filedialog.askopenfilename(title = "Select {}".format(fileType), filetypes = (("Excel Files","*.xls*"),("CSV Files","*.csv*"),("All Files","*.*")))
    if selectedFile != "":
        if fileType == "2A":
            global file2A
            file2A = XLXFile(selectedFile)
            file2ALabel.config(text = file2A.name + file2A.ext)
        elif fileType == "B2B":
            global fileB2B
            fileB2B = XLXFile(selectedFile)
            fileB2BLabel.config(text = fileB2B.name + fileB2B.ext)

def findCols(cols):
        colList = []
        sort_order = ["name", "gstin", "invno", "invvalue", "cgst", "sgst"]
        for x in cols:
            if "name" in x.lower() or "party" in x.lower():
                colList += [("name",cols.index(x))]
            elif "gstin" in x.lower() or "gst of supplier" in x.lower():
                colList += [("gstin",cols.index(x))]
            elif "invoice no" in x.lower() or "invoice number" in x.lower():
                colList += [("invno",cols.index(x))]
            elif "invoice value" in x.lower():
                colList += [("invvalue",cols.index(x))]
            elif "cgst" in x.lower() or "central" in x.lower():
                colList += [("cgst",cols.index(x))]
            elif "sgst" in x.lower() or "state or ut tax" in x.lower() or "state/ut tax" in x.lower():
                colList += [("sgst",cols.index(x))]
        return ([t[1] for x in sort_order for t in colList if t[0] == x])

def pdfData(gstinList):
    global xlist
    from fpdf import FPDF
    pdf = FPDF(format='letter', unit='in')
    pdf.add_page()
    pdf.set_font('Times','',10.0) 
    epw = pdf.w - 2*pdf.l_margin
    col_width = epw/8
    th = pdf.font_size
    data3 = []
    for num,i in enumerate(gstinList):
        if i!=('',''):
            data3 = []
            pdf.ln(2*th),pdf.set_font('Times','B',14.0) 
            pdf.cell(epw, 0.0, txt = "{} - {}".format(i[0],i[1]), align = 'C')
            pdf.set_font('Times','',10.0) 
            pdf.ln(0.3)
            data1 = [[k[0],k[1],k[2],k[3]] for k in mainDB.execute("select invno,invvalue,cgst,sgst from file2A where gstin = '{}'".format(i[1]))]
            data2 = [[k[0],k[1],k[2],k[3]] for k in mainDB.execute("select invno,invvalue,cgst,sgst from fileB2B where gstin = '{}'".format(i[1]))]

            if len(data1)>len(data2):
                length = len(data1)
                for b in range(len(data1)-len(data2)):
                    data2+=[['-','-','-','-']]
            elif len(data2)>len(data1):
                length = len(data2)
                for b in range(len(data2)-len(data1)):
                    data1+=[['-','-','-','-']]
            else:
                length = len(data2)
            for f in range(length):
                if data1[f]!=['','','',''] and data2[f]!=['','','','']:
                    data3+=[data1[f]+data2[f]]
            pdf.set_font('Times','B',12.0)
            pdf.cell(epw/2, 2*th, file2A.name, border=1,align='C')
            pdf.cell(epw/2, 2*th, fileB2B.name, border=1,align='C')
            pdf.ln(2*th)
            pdf.set_font('Times','B',10.0)
            for i in range(2):
                for k in ["Invoice No","Invoice Value","CGST","SGST"]:
                    pdf.cell(epw/8, 2*th, k, border=1,align='C')
            pdf.ln(2*th)
            pdf.set_font('Times','',10.0)
            for row in data3:
                for datum in row:
                    pdf.cell(col_width, 1.5*th, str(datum), border=1,align='C')
                pdf.ln(1.5*th)
    if alterName.get() == "":
        pdf.output("C:/Users/{}/Desktop/{} VS {}.pdf".format(userProfile,file2A.name,fileB2B.name),"F")
        finalLabel.config(text = "{} VS {}.pdf Saved to Desktop.".format(file2A.name,fileB2B.name))
    else:
        pdf.output("C:/Users/{}/Desktop/{}.pdf".format(userProfile,alterName.get()),"F")
        finalLabel.config(text = "{}.pdf Saved to Desktop.".format(alterName.get()))


def mainProcess():
    global mainDB
    # CREATE TEMP FILES
    os.mkdir("2A-B2B Temp")
    if file2A.ext != "csv":pandas.read_excel(r'{}'.format(file2A.path)).to_csv(r'2A-B2B Temp/{}.csv'.format(file2A.name), index = None)
    else:shutil.copy(file2A.path,'2A-B2B Temp')
    if fileB2B.ext != "csv":pandas.read_excel(r'{}'.format(fileB2B.path)).to_csv(r'2A-B2B Temp/{}.csv'.format(fileB2B.name), index = None)
    else:shutil.copy(fileB2B.path,'2A-B2B Temp')

    rawCSVFile2A = open("2A-B2B Temp/{0}.csv".format(file2A.name),"r")
    rawCSVFileB2B = open("2A-B2B Temp/{0}.csv".format(fileB2B.name),"r")

    CSVReader2A = csv.reader(rawCSVFile2A)
    CSVReaderB2B = csv.reader(rawCSVFileB2B)

    #FINDING COLUMN INDEXES
    for row in CSVReader2A:
        if any("Invoice" in column for column in row):
            file2AColumns = row
            break
    for row in CSVReaderB2B:
        if any("Invoice" in column for column in row):
            fileB2BColumns = row
            break

    #INSERTING DATA IN SQLITE DATABASE
    mainDB = sqlite3.connect("2A-B2B Temp/mainDB.db")
    mainDB.execute("Create table file2A (name, gstin, invno, invvalue, cgst, sgst)")
    mainDB.execute("Create table fileB2B (name, gstin, invno, invvalue, cgst, sgst)")


    for row in CSVReader2A: mainDB.execute("insert into file2A values {}".format(tuple([row[x] for x in findCols(file2AColumns)])))
    
    if len(findCols(fileB2BColumns)) == 5:
        for row in CSVReaderB2B:
            if row != []: mainDB.execute("insert into fileB2B values {}".format(tuple(['']+[row[x] for x in findCols(fileB2BColumns)])))
    else:
        for row in CSVReaderB2B:
            if row != []: mainDB.execute("insert into fileB2B values {}".format(tuple([row[x] for x in findCols(fileB2BColumns)])))
    
    mainDB.commit()


    #SORTING DATA WITH GST AND MAKING PDF
    file2AGST = [x[0] for x in mainDB.execute("select gstin from file2A group by gstin")]
    fileB2BGST = [x[0] for x in mainDB.execute("select gstin from fileB2B group by gstin")]
    totalGSTList = []
    for num, gstin in enumerate(list(dict.fromkeys(file2AGST + fileB2BGST))):
        try:
            partyName = mainDB.execute("select name from file2A where gstin = '{}'".format(gstin)).fetchone()[0]
        except:
            partyName = ""
        totalGSTList +=[(partyName, gstin)]
    totalGSTList.sort(key = lambda e:e[0])
    pdfData(totalGSTList)
    
    mainDB.close(),rawCSVFile2A.close(),rawCSVFileB2B.close()
    shutil.rmtree("2A-B2B Temp")

mainRoot = Tk()
mainRoot.geometry("500x330")
mainRoot.title("Welcome to 2A / B2B Manager")
mainFrame = Frame(mainRoot)
mainFrame.pack(fill = BOTH, expand = 1, padx = 10, pady = 20)

Label(mainFrame, text = "Welcome to 2A / B2B Manager", font = ("JUST DO GOOD",20)).grid(column = 0, row = 0,columnspan = 4)
ttk.Separator(mainFrame, orient=HORIZONTAL).grid(sticky = E+W,row=1,column=0,columnspan=4, pady = 10)

Button(mainFrame, text ="Select 2A", command = lambda:[browseFile("2A")],height=2, relief = "ridge").grid(sticky = "ew",row=2,column=0,columnspan=2)
Button(mainFrame, text ="Select B2B", command = lambda:[browseFile("B2B")],height=2, relief = "ridge").grid(sticky = "ew",row=2,column=2,columnspan=2)

file2ALabel = Label(mainFrame, text="No file Selected",width = 30,height=2)
fileB2BLabel = Label(mainFrame, text="No file Selected",width = 30,height=2)
file2ALabel.grid(row=3,column=0,columnspan=2)
fileB2BLabel.grid(row=3,column=2,columnspan=2)

ttk.Separator(mainFrame, orient=HORIZONTAL).grid(sticky = E+W,row=4,column=0,columnspan=4, ipadx=100)

Label(mainFrame,text = "Alternate Name of File to be Saved :", relief = "groove",height = 2).grid(sticky = "ew",column = 0,row = 5, columnspan = 2,pady = 10)
alterName = Entry(mainFrame,width = 21,font=("Times New Roman",15),justify = "center")
alterName.grid(row = 5, column = 2,padx=(12,0))

ttk.Separator(mainFrame, orient=HORIZONTAL).grid(sticky = E+W,row=6,column=0,columnspan=4, ipadx=100)
Button(mainFrame,text = "Start",command = lambda: [mainProcess()], bg='#4DE8FF', fg='black', font= ("JUST DO GOOD",13)).grid(sticky = E+W, row=7,columnspan=4, pady = 10)

finalLabel = Label(mainFrame,text = "",width = 60)
finalLabel.grid(row=8,column=0,columnspan=4)

ttk.Separator(mainFrame, orient=HORIZONTAL).grid(sticky = E+W,row=9,column=0,columnspan=4, ipadx=100)
Label(mainFrame, text = "Copyright (c) Harshit Bansal (2021)").grid(row = 10, column = 0, columnspan = 4)
mainRoot.mainloop()
