#! python3

import os
from tkinter import *
from tkinter import _setit,filedialog
from tkinter.ttk import *
from ttkthemes.themed_tk import *
import PyPDF2
import subprocess


class pdfLine():
    # Defines a GUI line with file input, and start/end buttons to pick pages
    def __init__(self,app,name,lineNum):
        self.label = [Label(app,text=name)]
        self.label[0].grid(row=lineNum,column=0)
        self.file = Entry(app,state='readonly')
        self.file.grid(row=lineNum,column=1)
        self.dir_button = Button(app,text='...',command=lambda: self.pdfBrowse(lineNum))
        self.dir_button.grid(row=lineNum,column=2)
        self.PageStart = StringVar(app)
        self.PageEnd = StringVar(app)
        self.label.append(Label(app,text='Start'))
        self.label[1].grid(row=lineNum,column=3)
        self.startPop = OptionMenu(app,self.PageStart,0)
        self.startPop.grid(row=lineNum,column=4)
        self.label.append(Label(app,text='End'))
        self.label[2].grid(row=lineNum,column=5)
        self.endPop = OptionMenu(app,self.PageEnd,0)
        self.endPop.grid(row=lineNum,column=6)

    def grid_remove(self):
        for label in self.label:
            label.grid_forget()
        self.file.grid_forget()
        self.dir_button.grid_forget()
        self.startPop.grid_forget()
        self.endPop.grid_forget()

    def pdfBrowse(self,lineNum):
        dir = filedialog.askopenfilename(filetypes=[("PDF Files","*.pdf")])
        if dir == '':
            return
        self.file.configure(state='normal')
        self.file.delete(0,'end')
        self.file.insert(0,dir)
        self.file.configure(state='readonly')
        self.pdfPageOptions(lineNum)

    def pdfPageOptions(self,lineNum):
        pdfFile = open(self.file.get(),'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFile)
        pageNum = pdfReader.numPages
        self.startPop['menu'].delete(0,'end')
        self.PageStart.set('')
        self.endPop['menu'].delete(0,'end')
        self.PageEnd.set('')
        for i in range(1,pageNum+1):
            self.startPop['menu'].add_command(label=i, command=_setit(self.PageStart, i))
            self.endPop['menu'].add_command(label=i, command=_setit(self.PageEnd, i))

    def getEntry(self,_type):
        if _type == 'dir':
            return self.file.get()
        elif _type == 'start':
            return self.PageStart.get()
        elif _type == 'end':
            return self.PageEnd.get()


class outLine():
    # Defines a GUI line with output folder and output filename input
    def __init__(self,app,name,lineNum):
        self.name_label = Label(app,text=name)
        self.name_label.grid(row=lineNum,column=0)
        self.dir = Entry(app,state='readonly')
        self.dir.grid(row=lineNum,column=1)
        self.dir_button = Button(app,text='...',command=lambda: self.outBrowse())
        self.dir_button.grid(row=lineNum,column=2)
        self.out_label = Label(app,text='File name:')
        self.out_label.grid(row=lineNum,column=3)
        self.outName = Entry(app)
        self.outName.insert(0,'MergedPDF')
        self.outName.grid(row=lineNum,column=4,columnspan=3)

    def grid_remove(self):
        self.name_label.grid_forget()
        self.dir.grid_forget()
        self.dir_button.grid_forget()
        self.out_label.grid_forget()
        self.outName.grid_forget()

    def outBrowse(self):
        dir = filedialog.askdirectory()
        if dir == '':
            return
        self.dir.configure(state='normal')
        self.dir.delete(0,'end')
        self.dir.insert(0,dir)
        self.dir.configure(state='readonly')

    def getEntry(self,_type):
        if _type == 'dir':
            return self.dir.get()
        if _type == 'out':
            return self.outName.get()


class Application(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.grid()
        self.Line = []
        self.actionButton = []
        self.fileNum()

    def fileNum(self):
        lineNum = len(self.Line)
        Label(self,text='Number of files:').grid(row=lineNum,column=0)
        self.numFiles= StringVar(self)
        self.Line.append(OptionMenu(self,self.numFiles,'2',*tuple(range(1,10)),command=self.addFileLines))
        self.Line[lineNum].grid(row=lineNum,column=1)
        self.addFileLines(2)

    def addFileLines(self,num):
        length = len(self.Line)
        if length == 1: # only file amount line exists
            fileLines = 0
        if length > 1:  # there are atleast 4 lines, file amount, file output, merge button, atleast 1 input line
            fileLines = length-3
            if fileLines == num:
                return
            else:
                self.Line[-2].grid_remove()
                del self.Line[-2]
                self.Line[-1].grid_remove()
                del self.Line[-1]
            if fileLines > num:
                for i in range(num+1,fileLines+1):
                    self.Line[i].grid_remove()
                del self.Line[num+1:]

        while fileLines < num:
            fileLines = fileLines + 1
            self.create_input('File ' + str(fileLines) + ':',Type='pdf')
        self.create_input('Destination:',Type='out')
        self.create_action("Merge",self.mergePages)

    def create_input(self,name,Type='Standard'):
        lineNum = len(self.Line)
        if Type == 'Standard':
            Label(self,text=name).grid(row=lineNum,column=0)
            self.Line.append(Entry(self))
            self.Line[-1].grid(row=lineNum,column=1)
        if Type == 'pdf':
            self.Line.append(pdfLine(self,name,lineNum))
        if Type == 'out':
            self.Line.append(outLine(self,name,lineNum))

    def create_action(self,name,func):
        self.Line.append(Button(self,text=name,command=func))
        self.Line[-1].grid(column=5,columnspan=2)

    def mergePages(self):
        pdfWriter = PyPDF2.PdfFileWriter()
        pdffiles = []
        for line in self.Line:
            if type(line) is not pdfLine:
                continue
            pdffiles.append(open(line.getEntry('dir'),'rb'))
            pdfreader = PyPDF2.PdfFileReader(pdffiles[-1])
            start = int(line.getEntry('start'))
            end = int(line.getEntry('end'))
            pages = range(start,end+1) if start<end else range(start,end-1,-1)
            for pagenum in pages:
                pdfWriter.addPage(pdfreader.getPage(pagenum-1)) # pages begin at 0
        outname = os.path.join(self.Line[-2].getEntry('dir'),self.Line[-2].getEntry('out'))+'.pdf'
        pdfout = open(outname,'wb')
        pdfWriter.write(pdfout)
        pdfout.close()
        for file in pdffiles:
            file.close()
        subprocess.Popen(outname,shell=True)


root = ThemedTk()
root.title("PDF Merger")
root.set_theme('arc')
app = Application(master=root)
app.mainloop()