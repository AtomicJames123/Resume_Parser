import csv
import os
import filetype
import docx2txt
import docx
import PyPDF2
import re
import collections
from tkinter import filedialog
from tkinter import *
from subprocess import Popen
from docx import Document
from nameparser import HumanName
from pathlib import Path

#Title = input("Please enter a job title: ")
#Title = Title + ".csv"

email = "No Email Found"
phone = "No Phone Number Found"
name = "No Name was Found"

FileTypesArray = [("PDF Files", "*.pdf"), ("Word Files", "*.docx")]

#EmailParsingArray = ["@", ".com", ".net"]

root = Tk()
file = filedialog.askopenfilename(initialdir="/", title="Choose a PDF or Word file", filetypes=FileTypesArray)

name = str(Path(file).stem)
name = re.sub('[_]', " ", name)

if file is not None:

        kind = filetype.guess(file)
        mime = kind.mime
        ext = kind.extension

        if str(ext) == "pdf":
            fileReader = PyPDF2.PdfFileReader(file)
            number_Pages = fileReader.numPages
            c = collections.Counter(range(number_Pages))
            for i in c:
                page = fileReader.getPage(i)
                content = page.extractText()
                print(content)

        elif str(ext) == "zip":
            document = docx.Document(file)

            for i in document.paragraphs:
                temp = i.text

                #if all(x in temp for x in EmailParsingArray):
                #   email = temp

                if re.search(r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)", temp):
                    email = temp
                    email = email.strip()

                if re.search(r"\(?\d{3}\)?[-.\s]\d{3}[-.\s]\d{4}", temp):
                    phone = temp
                    phone = phone.strip()

            print("Name: " + str(name))
            print("Email: " + email)
            print("Phone: " + phone)

        else:
            print("Bad Work")

else:
    sys.exit("You did not choose a file. Terminating Program")
