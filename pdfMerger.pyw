#!python3

import os
import sys
from tkinter import *
from tkinter import _setit, filedialog
from tkinter.ttk import *
from ttkthemes.themed_tk import *
import tkinter.messagebox as messagebox
import PyPDF2
import subprocess

f = open(os.devnull, 'w')  # change for debug
sys.stderr = sys.stdout = f  # must redirect stderr/stdout in .pyw files, since theres no terminal


class pdfLine():
    # Defines a GUI line with file input, and start/end buttons to pick pages
    def __init__(self, app, name, lineNum):
        # initiate all of the line's widgets, and add them to the grid in line "lineNum"
        self.lineNum = lineNum
        self.label = [Label(app, text=name)]
        self.label[0].grid(row=lineNum, column=0)
        self.file = Entry(app, state='readonly')
        self.file.grid(row=lineNum, column=1)
        self.dir_button = Button(app, text='...',
                                 command=self.pdfBrowse)  # make a button with press action - call self.pdfBrowse()
        self.dir_button.grid(row=lineNum, column=2)
        self.PageStart = StringVar(app)
        self.PageEnd = StringVar(app)
        self.label.append(Label(app, text='Start'))
        self.label[1].grid(row=lineNum, column=3)
        self.startPop = OptionMenu(app, self.PageStart, 0)
        self.startPop.grid(row=lineNum, column=4)
        self.label.append(Label(app, text='End'))
        self.label[2].grid(row=lineNum, column=5)
        self.endPop = OptionMenu(app, self.PageEnd, 0)
        self.endPop.grid(row=lineNum, column=6)

    def grid_remove(self):
        # remove all of the line's widgets from the grid
        for label in self.label:
            label.grid_forget()
        self.file.grid_forget()
        self.dir_button.grid_forget()
        self.startPop.grid_forget()
        self.endPop.grid_forget()

    def pdfBrowse(self):
        # used for the '...' button, opens a file dialog to get pdf name, and puts it into the entry
        dir = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if dir == '':
            return
        self.file.configure(state='normal')
        self.file.delete(0, 'end')
        self.file.insert(0, dir)
        self.file.configure(state='readonly')
        self.pdfPageOptions()

    def pdfPageOptions(self):
        # update the start/end options with relevant page numbers
        pdfFile = open(self.file.get(), 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFile)
        pageNum = pdfReader.getNumPages()
        pdfFile.close()
        if pageNum == 0:
            return
        self.startPop['menu'].delete(0, 'end')  # remove old options from start/end
        self.PageStart.set('1')
        self.endPop['menu'].delete(0, 'end')
        self.PageEnd.set(pageNum)
        for i in range(1, pageNum + 1):  # add the new options
            self.startPop['menu'].add_command(label=i, command=_setit(self.PageStart, i))
            self.endPop['menu'].add_command(label=i, command=_setit(self.PageEnd, i))

    def getEntry(self, _type):
        if _type == 'dir':
            return self.file.get()
        elif _type == 'start':
            return self.PageStart.get()
        elif _type == 'end':
            return self.PageEnd.get()


class outLine():
    # Defines a GUI line with output folder and output filename input
    def __init__(self, app, name, lineNum):
        # initiate all of the line's widgets, and add them to the grid in line "lineNum"
        self.name_label = Label(app, text=name)
        self.name_label.grid(row=lineNum, column=0)
        self.dir = Entry(app)
        self.dir.insert(0,os.getcwd())
        self.dir.configure(state='readonly')
        self.dir.grid(row=lineNum, column=1)
        self.dir_button = Button(app, text='...', command=self.outBrowse)
        self.dir_button.grid(row=lineNum, column=2)
        self.out_label = Label(app, text='File name:')
        self.out_label.grid(row=lineNum, column=3)
        self.outName = Entry(app)
        self.outName.insert(0, 'MergedPDF')
        self.outName.grid(row=lineNum, column=4, columnspan=3)

    def grid_remove(self):
        # remove all of the line's widgets from the grid
        self.name_label.grid_forget()
        self.dir.grid_forget()
        self.dir_button.grid_forget()
        self.out_label.grid_forget()
        self.outName.grid_forget()

    def outBrowse(self):
        # used for the '...' button
        dir = filedialog.askdirectory()
        if dir == '':
            return
        self.dir.configure(state='normal')
        self.dir.delete(0, 'end')
        self.dir.insert(0, dir)
        self.dir.configure(state='readonly')

    def getEntry(self, _type):
        if _type == 'dir':
            return self.dir.get()
        if _type == 'out':
            return self.outName.get()


class Application(Frame):
    # main window class
    def __init__(self, master=None):
        # initiate main window, and add first line
        super().__init__(master)
        self.grid()
        self.Line = []
        self.actionButton = []
        self.fileNum()

    def fileNum(self):
        # adds the first line, defaulted to 2 files
        lineNum = len(self.Line)
        default = 2
        maxFiles = 15
        Label(self, text='Number of files:').grid(row=lineNum, column=0)
        self.numFiles = StringVar(self)
        self.Line.append(
            OptionMenu(self, self.numFiles, str(default), *tuple(range(1, maxFiles + 1)), command=self.addFileLines))
        self.Line[-1].grid(row=lineNum, column=1)
        self.addFileLines(default)  # add the rest of the lines

    def addFileLines(self, num):
        # add/remove file input lines to match *num* input lines, followed by output line and Merge button
        length = len(self.Line)
        if length == 1:  # only file amount line exists
            fileLines = 0
        if length > 1:  # there are atleast 4 lines: file amount, file output, merge button, atleast 1 input line
            fileLines = length - 3
            if fileLines == num:
                return
            else:
                # delete output/button lines and spare input lines if needed
                self.Line[-2].grid_remove()
                del self.Line[-2]
                self.Line[-1].grid_remove()
                del self.Line[-1]
            if fileLines > num:
                for i in range(num + 1, fileLines + 1):
                    self.Line[i].grid_remove()
                del self.Line[num + 1:]

        while fileLines < num:
            fileLines = fileLines + 1
            self.create_input('File ' + str(fileLines) + ':', Type='pdf')
        self.create_input('Destination:', Type='out')
        self.create_action("Merge", self.mergePages)

    def create_input(self, name, Type='Standard'):
        # wrapper function for creating different user-input lines
        lineNum = len(self.Line)
        if Type == 'Standard':
            Label(self, text=name).grid(row=lineNum, column=0)
            self.Line.append(Entry(self))
            self.Line[-1].grid(row=lineNum, column=1)
        if Type == 'pdf':
            self.Line.append(pdfLine(self, name, lineNum))
        if Type == 'out':
            self.Line.append(outLine(self, name, lineNum))

    def create_action(self, name, func):
        # wrapper function for creating different action lines, only adds Merge button at the moment
        self.Line.append(Button(self, text=name, command=func))
        self.Line[-1].grid(column=5, columnspan=2)

    def mergePages(self):
        # get all the arguments given, and merge the relevant pages
        pdfWriter = PyPDF2.PdfFileWriter()  # pdf object for the output pdf
        pdffiles = []
        if self.Line[-2].getEntry('out') == "" or self.Line[-2].getEntry('dir') == "":
            messagebox.showerror("Error","Missing output parameters")
            return
        outname = os.path.join(self.Line[-2].getEntry('dir'), self.Line[-2].getEntry('out')) + '.pdf'
        try:
            for line in self.Line:
                if type(line) is not pdfLine:
                    continue
                cDir = line.getEntry('dir')
                if cDir == "":
                    messagebox.showerror("PDF Merger","Missing file #"+str(line.lineNum))
                    raise IOError
                if line.getEntry('start') == "" or line.getEntry('end') == "":
                    messagebox.showerror("PDF Merger","File #"+str(line.lineNum)+" is empty")
                    raise IOError
                pdffiles.append(open(cDir, 'rb'))
                pdfreader = PyPDF2.PdfFileReader(pdffiles[-1])  # pdf object for the file in the current line
                start = int(line.getEntry('start'))
                end = int(line.getEntry('end'))
                pages = range(start, end + 1) if start < end else range(start, end - 1,-1)  # determine the page range, and direction (normal/backwards)
                for pagenum in pages:
                    pdfWriter.addPage(pdfreader.getPage(pagenum - 1))  # pages begin at 0
        except FileNotFoundError as e:  # failed to open one of the files
            for file in pdffiles:
                file.close()
            messagebox.showerror("Error", "Failed to open file: "+e.filename)
            return
        except IOError:
            for file in pdffiles:
                file.close()
            return
        if os.path.exists(outname):
            choice = messagebox.askyesnocancel("PDF Merger","Do you want to overwrite "+outname+" ?\n(Choosing no will create the file with a different name)")
            if choice is None:
                for file in pdffiles:
                    file.close()
                return
            elif not choice:
                i = 1
                while os.path.exists(os.path.join(self.Line[-2].getEntry('dir'), self.Line[-2].getEntry('out')) + '_'+ str(i) +'.pdf'):
                    i+=1
                outname = os.path.join(self.Line[-2].getEntry('dir'), self.Line[-2].getEntry('out')) + '_'+ str(i) +'.pdf'
        pdfout = open(outname + '.temp', 'wb')
        pdfWriter.write(pdfout)  # write the new pdf to pdfout
        pdfout.close()
        for file in pdffiles:
            file.close()
        try:
            os.replace(outname + '.temp', outname)
            # TODO: update lines if the file was overwritten, below code doesn't work for some reason
            #for line in self.line:
            #    if type(line) is pdfline:
            #        if line.getEntry('dir') == outname:
            #            line.pdfPageOptions()
        except:
            os.remove(outname+'.temp')
            messagebox.showerror('Error','Error overwriting file:\n'+outname+'\nMake sure the file isn\'t open elsewhere')
            return
        subprocess.Popen(outname, shell=True)  # open output pdf for preview by the user


# main tkinter loop        
root = ThemedTk()
root.title("PDF Merger")
root.set_theme('arc')
app = Application(master=root)
app.mainloop()
