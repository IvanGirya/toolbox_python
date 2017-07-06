from tkinter import *
from tkinter.filedialog import *
import tkinter.ttk as ttk
import pyodbc
import pathlib
import os
import datetime
import PIL
import pytesseract
import binascii
import getpass
import subprocess
import tkinter
import PyPDF2
import binascii
import pyhdb
import subprocess
import pyocr
import stopit
from pyocr import pyocr
import pyocr.builders
import pdfminer
from langdetect import detect
import io
import time
import image as BI
#import pptx
import textract
from wand.image import Image as WImage
from PIL import Image as PI
from PIL import ImageEnhance, ImageFilter
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files (x86)/Tesseract-OCR/tesseract'
from pytesseract import *
from PIL import Image, ImageTk
from PIL import ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES = True

from PyPDF2 import PdfFileReader
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

########### - START OF GLOBAL SETUPS AND VARIABLES - #########

global new_load
new_load = 0 #0 - continue previous load, 1 - start new load

global directory
directory = ""

hana_servernode = 'sdihost:30041'
hana_serverdb = 'SDI'
hana_uid = 'SYSTEM'
hana_password = 'Express1'
hana_connect_str = 'DRIVER={HDBODBC};SERVERNODE=' + hana_servernode + ';SERVERDB=' + hana_serverdb + ';UID=' + hana_uid + ';PWD=' + hana_password

image_files_extensions = [".jpg", ".jpeg", ".png", ".gif", ".tif", ".tif", ".bmp", ".JPG"]
stop_files_extensions = [".gz", ".rar", ".zip", ".tar", ".tmp", ".url", ".lnk", ".mp3", ".mp4", ".vmv", ".vid", ".avi", ".db", ".exe", ".bin", ".GID", ".dll", ".cab", ".msi", ".ex_", ".sys", ".cat", ".SYS", ".chm", ".dl_"]
texract_files_extensions = [".xls", ".xlsx"] #".pptx", ".msg", ".eml", ".epub"]

pdftotext_path = "d:/1/pdftotext" #location of PDF-to-Text package

python_treshold = 500 #Gb
content_treshold = 10 #Gb

file_size_max = 30000 #kb
img_size_max = 2000 #kb
pdf_page_max = 20 #number of pages
maximum_time = 180 #maximum time for processing one file - seconds

#get last ID in the target table
conn = pyodbc.connect(hana_connect_str) #connect to HANA
cur = conn.cursor() #Open a cursor
query = "select case when max(FILE_ID) is null then 0 else max(FILE_ID) end as FILE_ID from \"TOOLBOX\".\"gdpr-demo.tables::FILES_CONTENT\"" #specify current ID for Files
ret=cur.execute(query) #execute query to define File ID
ret=cur.fetchall()
for row in ret:
    for col in row:
        file_number = col #starting ID for new files to be loaded
cur.execute("COMMIT")
cur.close
#conn.close

##if new_load == 1:
##    directory = tkinter.filedialog.askdirectory(initialdir="d:/2/PoC/",title='Please select a loading directory') #directory to load from
##else:
##
#subdirectory = tkinter.filedialog.askdirectory(initialdir="d:/2/PoC/",title='Subdirectory of last load') #subdirectory when last load was finished

global last_file_seeker
last_file_seeker = 0

last_file =''

user_name = getpass.getuser() #user
computer_name = os.environ['COMPUTERNAME'] #computer name

load_date = time.strftime("%Y%m%d")

#tool = pyocr.get_available_tools()[0]
#lang = tool.get_available_languages()[1]

########### - END OF GLOBAL SETUPS AND VARIABLES - #########



def _readIt(): ## this is the variable the button changes
    global bg_read
    bg_read = 1
    TK.after(0, _profiling) ##this is how you make sure the GUI doesn't freeze up

#Function - converts pdf, returns its text content as a string
def convert_pdf(fname, pages=None):
    if not pages:
        pagenums = set()
    else:
        pagenums = set(pages)

    output = StringIO()
    manager = PDFResourceManager()
    converter = TextConverter(manager, output, laparams=LAParams())
    interpreter = PDFPageInterpreter(manager, converter)

    infile = open(fname, 'rb')
    #parser = PDFParser(fp)
    #document = PDFDocument(parser, password)
    #if not document.is_extractable:
     #   ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, 'PDF extraction not allowed')
    try:
        for page in PDFPage.get_pages(infile, pagenums, check_extractable=True):
            #print(time.strftime("%H:%M:%S") + " convert_PDF: processing pages")
            interpreter.process_page(page)
    except:
        output.write ('ZZZ')
    infile.close()
    converter.close()
    text = output.getvalue()
    output.close
    return text

#Function - populates HANA table of Rejected Files when error occurs in the loading
def file_rejected(file_number, computer_name, fname, file_size_kb, file_modification_date, user_name, file_extension, error_desc):
    conn_rej = pyodbc.connect(hana_connect_str) #connect to HANA
    cur_rej = conn_rej.cursor() #Open a cursor
    cur_rej.execute('INSERT INTO "TOOLBOX"."gdpr-demo.tables::FILES_REJECTED" ( FILE_ID, SOURCE_ID, FILENAME, FILESIZE, FILECHANGED, SOURCE_NAME, FILE_TYPE, LOADING_DATE, ERROR ) VALUES (?,?,?,?,?,?,?,?,?)',  #Save the content to the table
                    (file_number,
                     computer_name,
                     fname,
                     file_size_kb,
                     file_modification_date,
                     user_name,
                     file_extension,
                     load_date,
                     error_desc
                     ))
    cur_rej.execute("COMMIT") #Save the content to the table
    cur_rej.close() #Close the cursor
    conn_rej.close() #Close the connection
    print(time.strftime("%H:%M:%S") + " - ! FILE " + str(file_number) + " REJECTED: " + fname)


def img_driver_license(file_path, file_number):

    img = PI.open(file_path).convert('L')

    scale_value=80
    contrast = ImageEnhance.Contrast(img)  #converting to grayscale
    contrast_applied=contrast.enhance(scale_value)

    img = img.convert("RGBA")
    img = img.filter(ImageFilter.SMOOTH_MORE) #blurring to make noise weaker
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(0.95)  # NEED TO PLAY WITH THIS PARAMETER  # lint:ok

    pixdata = img.load()

    #print(time.strftime("%H:%M:%S") + " text extraction")
    content_text = pytesseract.image_to_string(img, lang="dan+eng", config="-psm 1")
    img = img.rotate(90)
    content_text = content_text + " " + pytesseract.image_to_string(img, lang="dan+eng", config="-psm 1")
    img = img.rotate(90)
    content_text = content_text + " " + pytesseract.image_to_string(img, lang="dan+eng", config="-psm 1")
    img = img.rotate(90)
    content_text = content_text + " " + pytesseract.image_to_string(img, lang="dan+eng", config="-psm 1")


    #!---- IN SOME CASES ADDITIONAL CHECK IS NECESSARY WITH EXTREME CONTRAST LEVELS
    img = img.rotate(90)
    # Make the letters bolder for easier recognition
    for y in range(img.size[1]):
        for x in range(img.size[0]):
            if pixdata[x, y][0] < 40:
                pixdata[x, y] = (0, 0, 0, 255)

    for y in range(img.size[1]):
        for x in range(img.size[0]):
            if pixdata[x, y][1] < 80:
                pixdata[x, y] = (0, 0, 0, 255)

    for y in range(img.size[1]):
        for x in range(img.size[0]):
            if pixdata[x, y][2] > 0:
                pixdata[x, y] = (255, 255, 255, 255)

    content_text = content_text + " " + pytesseract.image_to_string(img, lang="dan+eng", config="-psm 1")
    #!-----END OF EXTREME CONTRAST

    content_text = content_text.replace(r"\r\n", r" ").replace(r"\r", r" ").replace(r"\n", r" ").replace(r",", r", ")

    #Checking for KØREKORT
    driver_license_words = ["KØREKORT", "KEREKORT", "KGREKORT", "KQREKORT", "KQREKQRT", "KØRE"]
    for img_word in content_text.split():
        #print(img_word.upper())
        if any(fuzz.ratio(word, img_word.upper()) > 70 for word in driver_license_words):
            print("Driving license found! " + file_path + " - " + img_word)
            content_text = content_text + " KØREKORT"
            break
    return content_text


def _open(event):
    txt.delete(0,END)
    op = askdirectory()
    txt.insert(END,op)


## Radio-button selection
def rb_selected():
     global new_load
     global txt
     if rbut.get()==1:
        new_load = 1
        txt.delete(0,END) #clear Entry field
     else:
        print("Continue from:")
        new_load = 0
        conn_last_file = pyodbc.connect(hana_connect_str) #connect to HANA
        cur_last_file = conn_last_file.cursor() #Open a cursor

        #ask for directory
        global directory
        cur_last_file.execute('SELECT LAST_FILE, START_DIRECTORY FROM "TOOLBOX"."gdpr-demo.tables::LOADS_HISTORY" WHERE SOURCE_NAME=? AND LOAD_TIME IN (SELECT MAX (LOAD_TIME) FROM "TOOLBOX"."gdpr-demo.tables::LOADS_HISTORY" WHERE SOURCE_NAME=?)',  #Save the content to the table
                            (computer_name,
                             computer_name
                             ))
        ret_last_file = cur_last_file.fetchall()
        for row in ret_last_file:
            last_file = row.LAST_FILE
            directory = row.START_DIRECTORY
        print(last_file)
        cur_last_file.execute("COMMIT") #Save the content to the table
        cur_last_file.close() #Close the cursor
        conn_last_file.close() #Close the connection
        #put directory name into an Entry field
        txt.delete(0,END)
        txt.insert(END,directory)


def _exit(event):
     root.destroy()

def _profiling(event):
    ### LOADING FILES TO HANA
    global new_load
    new_load = rbut.get()
    global directory
    global txt
    directory = txt.get()
    print(directory)
    global file_number
    content = ''


    ### SPECIFY LOAD DIRECTORY (IF LOAD IS NEW) OR LAST_FILE PARAMETER (IF CONTINUING THE PREVIOUS LOAD)
    conn_last_file = pyodbc.connect(hana_connect_str) #connect to HANA
    cur_last_file = conn_last_file.cursor() #Open a cursor
    if new_load == 1:
        #ask for directory
        #directory = tkinter.filedialog.askdirectory(initialdir="d:/2/PoC/",title='Please select a loading directory') #directory to load from

        top_file = os.listdir(directory)[0]
        cur_last_file.execute('INSERT INTO "TOOLBOX"."gdpr-demo.tables::LOADS_HISTORY" ( SOURCE_NAME, LOAD_TIME, START_DIRECTORY, LAST_FILE ) VALUES (?,?,?,?)',  #Save the content to the table
                        (computer_name,
                         datetime.datetime.fromtimestamp(datetime.datetime.now().timestamp()),
                         directory,
                         top_file
                         ))
    else:
        cur_last_file.execute('SELECT LAST_FILE FROM "TOOLBOX"."gdpr-demo.tables::LOADS_HISTORY" WHERE SOURCE_NAME=? AND LOAD_TIME IN (SELECT MAX (LOAD_TIME) FROM "TOOLBOX"."gdpr-demo.tables::LOADS_HISTORY" WHERE SOURCE_NAME=?)',  #Save the content to the table
                        (computer_name,
                         computer_name
                         ))
        ret_last_file = cur_last_file.fetchall()
        for row in ret_last_file:
            last_file = row.LAST_FILE
        #directory = os.path.dirname(last_file)
    cur_last_file.execute("COMMIT") #Save the content to the table
    cur_last_file.close() #Close the cursor
    conn_last_file.close() #Close the connection



    last_file_seeker = 0
    for dirpath, dirnames, files in os.walk(directory, topdown=True):
    ##    curr_subdirectory = dirpath.replace("\\", "/")
    ##    #print (curr_subdirectory)
    ##    if (curr_subdirectory != subdirectory and last_file_seeker == 0):
    ##        continue
    ##    else:
    ##        last_file_seeker = 1

        #start from last file
        for name in files:
                if new_load != 1 and os.path.join(dirpath, name) != last_file and last_file_seeker == 0:
                    last_file_seeker = 0
                    continue
                elif new_load != 1 and os.path.join(dirpath, name) == last_file:
                    last_file_seeker = 1
                    continue

                #increment file index
                file_number +=1

                #get file metadata
                content = ''
                file_path = os.path.join(dirpath, name)
                file_extension = os.path.splitext(file_path)[1]
                file_modification_date = datetime.datetime.fromtimestamp(os.path.getmtime(file_path)) #file last modification - LATER CHANGE FOR DYNAMIC AND OS-SPECIFIC
                file_size_kb =  os.path.getsize(file_path)/1000


                #!!!!!!!
                #!!!!!!!
                #!!!!!!! Temporary to process images only
                #!!!!!!!
                #!!!!!!!
                #img_size_max = 3000 #kb
                #if file_extension not in image_files_extensions:
                    #file_number -=1
                    #continue

                #!!!!!!!
                #!!!!!!!
                #!!!!!!!
                #!!!!!!!
                #!!!!!!!

                #Print that load started
                print("--")
                print(time.strftime("%H:%M:%S") + " - File " + str(file_number) + " started: " + file_path)

                #skip unsupported file formats and other restrictions
                if file_extension in stop_files_extensions: #unsupported type
                    ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, 'Extension not supported')
                    continue
                elif (file_extension in image_files_extensions and file_size_kb >= img_size_max): #image too large
                    ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, 'Image size too large')
                    continue
                elif file_size_kb >= file_size_max: #file too large
                    ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, 'File too large')
                    continue



                if file_size_kb < file_size_max: #file size is OK
                    print(time.strftime("%H:%M:%S") + " Opening file")
                    file_to_hana = open(file_path, 'rb') #Open file in read-only and binary

                    ########
                    #special logics for images - with OCR recognition
                    if ((file_extension.lower() in image_files_extensions) and chbx_img.get() == 1):
                        print ("Image processing started")
                        ### TO SKIP IMAGES
                        #ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, 'Skipping images')
                        #continue
                        try:
                            img = PI.open(file_path).convert('L')
                        except:
                            ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, "Cannot identify image file")
                            continue

                        if file_size_kb < img_size_max and chbx_img.get() == 1: #if Process Images option is active
                            try:
                                print ("Launching text extraction from OCR")
                                content_text = img_driver_license(file_path, file_number)
                                #content_text = pytesseract.image_to_string(img, lang="eng+dan", config="-psm 1")
                            except:
                                ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, "Error during extracting text from image")
                                continue
                        else:
                            ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, 'Image size too large')
                            continue
                        #print (content_text)
                        content = bytes(content_text, 'utf-8')#convert recognized string into bytes to insert it into BLOB type in HANA
                        content = content.replace(br"\r\n", br" ").replace(br"\r", br" ").replace(br"\n", br" ").replace(br",", br", ")
                    elif ((file_extension in image_files_extensions) and chbx_img.get() != 1):
                        ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, "Skipping images")
                        continue
                    elif file_extension.lower() == '.pdf':

                        #try extracting text from file - CRITICAL NOT TO HAVE SPACES IN NAMES
                        #pdf_command = pdftotext_path + ' -layout "' + file_path + '" -'
                        #pdf_text = subprocess.check_output(pdf_command, shell=True)
                        #pdf_text = pdf_text.replace(b"\r\n", b" ").replace(b"\r", b" ") #replace line breaks with spaces

                        #print(time.strftime("%H:%M:%S") + "Calling Convert_PDF")

                        #Timer for PDF processing. Skip if too long
                        with stopit.ThreadingTimeout(maximum_time) as to_ctx_mgr:
                            assert to_ctx_mgr.state == to_ctx_mgr.EXECUTING
                            pdf_text = convert_pdf(file_path)
                            if to_ctx_mgr.state == to_ctx_mgr.TIMED_OUT or to_ctx_mgr.state == to_ctx_mgr.INTERRUPTED or to_ctx_mgr.state == to_ctx_mgr.CANCELED:
                                ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, 'File processing stopped by timeout')
                                continue

                        #print(time.strftime("%H:%M:%S") + "Replacing special symbols")
                        pdf_text = pdf_text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ").replace(",", ", ") #replace line breaks with spaces
                        if pdf_text == 'ZZZ':
                            ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, 'PDF doesnt allow text extraction')
                            continue
                        #we consider PDF with more than 10 words a text file
                        if len(pdf_text.split()) > 10:
                            #print ("pdf with text")
                            #content = bytes(pdf_text) #Save file content to variable
                            content = pdf_text #Save file content to variable


                        else: #else - PDF with scans

                            #ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, 'Skipping images')
                            #continue

                            #if Process Images option is active
                            if chbx_img.get() == 1:
                                #HARDCODE - correct later problem with files with special symbols
                                try:
                                    pdf_file = PdfFileReader(open(file_path.encode('utf-8'), "rb"))
                                except:
                                    ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, 'Error when reading PDF')
                                    continue
                                #count pages in PDF-scan
                                if pdf_file.isEncrypted:
                                    try:
                                        pages_num = pdf_file.getNumPages()
                                        hana_connect_str
                                    except:
                                        file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, "PDF is protected or encrypted")
                                        pages_num = 999
                                        continue
                                    #if pages_num != 999: #HARDCODED
                                        #pages_num = pdf_file.getNumPages()
                                else:
                                    pages_num = pdf_file.getNumPages()

                                if pages_num < pdf_page_max:
                                    req_image = []
                                    final_text = []
                                    image_pdf = WImage(filename=file_path, resolution=300)
                                    image_jpeg = image_pdf.convert('jpeg')

                                    for img in image_jpeg.sequence: #process set of images from PDF
                                        img_page = WImage(image=img)
                                        req_image.append(img_page.make_blob('jpeg')) #list of images from PDF created

                                    for img in req_image: #each image converted to text and appended to list of strings
                                        txt = pytesseract.image_to_string(
                                            PI.open(io.BytesIO(img)),
                                            lang="dan+eng",
                                            builder=pyocr.builders.TextBuilder()
                                            )
                                        final_text.append(txt)
                                    content = ' '.join(final_text) #convert list to strings separated by spacestess
                                    content = content.replace("\r\n", " ").replace("\r", " ").replace("\n", " ") #replace line breaks with spaces
                                else:
                                    ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, "File is too big")
                                    continue
                            else:
                                ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, "Skipping PDF with only scans/images")
                                continue

                    ########
                    #use texract for files like xls, ppt, msg etc.
                    elif file_extension.lower() in texract_files_extensions:
                        try:
                            content = textract.process(file_path)
                        except:
                            ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, "Textract couldn't process the file")
                            continue
                        content = content.replace(br"\r\n", br" ").replace(br"\r", br" ").replace(br"\n", br" ").replace(br",", br", ")


                    else: #all other formats
                        content = file_to_hana.read()
                        content = content.replace(br"\r\n", br" ")



                else:
                    ret = file_rejected(file_number, computer_name, file_path, file_size_kb, file_modification_date, user_name, file_extension, "File is too big")
                    continue


                cur.execute('INSERT INTO "TOOLBOX"."gdpr-demo.tables::FILES_CONTENT" (FILE_ID,SOURCE_ID,CONTENT,LANG,FILENAME,FILESIZE,FILECHANGED,SOURCE_NAME) VALUES (?,?,?,?,?,?,?,?)', #Save the content to the table
                            (file_number,
                             computer_name,
                             content,
                             'en',
                             file_path,
                             file_size_kb,
                             file_modification_date,
                             user_name
                             ))
                #update the last_file variable
                cur.execute('UPDATE "TOOLBOX"."gdpr-demo.tables::LOADS_HISTORY" SET LAST_FILE = ?, START_DIRECTORY = ? WHERE SOURCE_NAME= ? AND LOAD_TIME IN (SELECT MAX (LOAD_TIME) FROM "TOOLBOX"."gdpr-demo.tables::LOADS_HISTORY" WHERE SOURCE_NAME=?)',
                            (file_path, directory, computer_name, computer_name))
                cur.execute("COMMIT") #Save the content to the table
                 #Close the file
                #Lis.insert(END, file_path)
                print(time.strftime("%H:%M:%S") + " - File " + str(file_number) + " loaded: " + file_path)

    #cur.close() #Close the cursor
    #conn.close() #Close the connection
    print("**** All files loaded to HANA ****")


## Gui creation
root = Tk()
root.wm_title("GDPR Unstructured Data loading")
root["bg"] = "white"
root.geometry('600x200+50+50')
a=StringVar()

image = Image.open('dm_logo.png')
photo = ImageTk.PhotoImage(image)

lbl_image = Label(root, image=photo, bg="white")#.place(x=500,y=5)


#Radiobuttons
rbut = IntVar()
rbut.set("1")
r1 = Radiobutton(root, text="New load", font="Montserrat 10", variable=rbut, value=1,command=rb_selected).place(x=50,y=15)
r2 = Radiobutton(root, text="Continue", font="Montserrat 10", variable=rbut, value=0,command=rb_selected).place(x=150,y=15)


#Checkboxes
chbx_img = IntVar()
Checkbutton(root, text="Process images", onvalue=1, offvalue=0, variable=chbx_img, bg="yellow").place(x=250,y=15)


## Labels description
#lab = Label(root, text="Select directory to start profling", font="Montserrat 16", fg="gray15", bg="white")
global txt
txt = Entry(root, textvariable=a, font="Montserrat 8", width="50")#.place(x=60,y=105)
lbl = Label(root, text="Source folder", font="Montserrat 10", bg="white", justify='right')
dummy_lbl = Label(root, text="", font="Arial 10", bg="white")
lbl_progress = Label(root, text="", font="Montserrat 8", bg="white", justify='left')
lbl_info = Label(root, text="", font=("Montserrat", 10, "bold"), bg="white")


## Buttons description
but_start = Button(root, text="Start profiling", width = 20, bg="gray15",fg="DarkGoldenrod1", font="Montserrat 10")#.place(x=60,y=135)
but_exit = Button(root, text="Exit", width = 20, bg="gray15",fg="DarkGoldenrod1", font="Montserrat 10")
but = Button(root, text="Select folder", width = 20, bg="gray15",fg="DarkGoldenrod1", font="Montserrat 10")#.place(x=400,y=105)
but.bind("<Button-1>", _open)
but_start.bind("<Button-1>", _profiling)
but_exit.bind("<Button-1>", _exit)


## Grid assesment
lbl_image.grid(row=0,column=2)
lbl.grid(row=1,column=0)
txt.grid(row=1,column=1)
but.grid(row=1,column=2)
dummy_lbl.grid(row=2,column=0)
but_start.grid(row=3,column=1)
but_exit.grid(row=3,column=2)
lbl_info.grid(row=4,column=0)


root.mainloop()

