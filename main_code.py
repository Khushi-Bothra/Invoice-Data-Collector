
import tkinter as tk
from tkinter import filedialog

from adobe.pdfservices.operation.auth.credentials import Credentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.execution_context import ExecutionContext
from adobe.pdfservices.operation.io.file_ref import FileRef
from adobe.pdfservices.operation.pdfops.extract_pdf_operation import ExtractPDFOperation
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_pdf_options import ExtractPDFOptions
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_element_type import ExtractElementType

from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_renditions_element_type import \
    ExtractRenditionsElementType

import logging
import os.path
import zipfile
import json
import csv
import pandas as pd
import shutil

"""
Here the below code is commented as we will not be taking the input of names of file 
individually since we already know the names of the file

file_names = []  # List to store the file names

file_no=int(input("Enter how many files you want extracted"))

for i in range(file_no): #Dynamically taking input of the names of files
    file_name = input("Enter the name of file {}: ".format(i + 1))
    file_names.append(file_name) """
    
header=[None]*19
header[0]="Bussiness__City"
header[1]="Bussiness__Country"
header[2]="Bussiness__Description"
header[3]="Bussiness__Name"
header[4]="Bussiness__StreetAddress"
header[5]="Bussiness__Zipcode"
header[6]="Customer__Address__line1"
header[7]="Customer__Address__line2"
header[8]="Customer__Email"
header[9]="Customer__Name"
header[10]="Customer__PhoneNumber"
header[11]="Invoice__BillDetails__Name"
header[12]="Invoice__BillDetails__Quantity"
header[13]="Invoice__BillDetails__Rate"
header[14]="Invoice__Description"
header[15]="Invoice__DueDate"
header[16]="Invoice__IssueDate"
header[17]="Invoice__Number"
header[18]="Invoice__Tax"

file = open('ExtractedDataByMe_Final_Test_File.csv', 'w', newline='')  #Writing the header for the file
writer = csv.writer(file)
writer.writerow(header)
file.close()

def process(file_path):
    #Here we have already set the range as 100 as it was given in the challenge
    for i in range(1):  #Starting a loop to extract all the files and insert the data in csv file one by one
        input_pdf = file_path
        
        zip_file_1 = "./ExtractTextInfoFromPDF_Final_Test.zip"   #Zipfile where the extracted json is stored
        
        if os.path.isfile(zip_file_1):  #Removing the said zipfile in case it already exists
            os.remove(zip_file_1)
        
        try:
        
            #Initial setup, create credentials instance.
            credentials = Credentials.service_account_credentials_builder()\
                .from_file("./pdfservices-api-credentials.json") \
                .build()

            #Create an ExecutionContext using credentials and create a new operation instance.
            execution_context = ExecutionContext.create(credentials)
            extract_pdf_operation = ExtractPDFOperation.create_new()

            #Set operation input from a source file.
            source = FileRef.create_from_local_file(input_pdf)
            extract_pdf_operation.set_input(source)

            #Build ExtractPDF options and set them into the operation
            extract_pdf_options: ExtractPDFOptions = ExtractPDFOptions.builder() \
                .with_element_to_extract(ExtractElementType.TEXT) \
                .build()
            extract_pdf_operation.set_options(extract_pdf_options)

            #Execute the operation.
            result: FileRef = extract_pdf_operation.execute(execution_context)

            #Save the result to the specified location.
            result.save_as(zip_file_1)

            print("Successfully extracted information from PDF \n")

            archive = zipfile.ZipFile(zip_file_1, 'r')   #Opening the zipfile
            jsonentry = archive.open('structuredData.json')
            jsondata = jsonentry.read()  #Reading the json file in order to get the required data
            data = json.loads(jsondata)
            archive.close()  #Closing the zipfile
            jsonentry.close()
        
        except (ServiceApiException, ServiceUsageException, SdkException):
            logging.exception("Exception encountered while executing operation")
        
        logging.basicConfig(level=os.environ.get("LOGLEVEL", "INFO"))

        zip_file_2 = "./ExtractTextTableInfoWithCharBoundsFromPDF_Final_Test.zip"  #Using another zipfile to extract the data from table

        if os.path.isfile(zip_file_2): #Removing the said zipfile in case it already exists
            os.remove(zip_file_2)

        try:
            # get base path.
            #base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

            # Initial setup, create credentials instance.
            credentials = Credentials.service_account_credentials_builder() \
                .from_file("./pdfservices-api-credentials.json") \
                .build()

            # Create an ExecutionContext using credentials and create a new operation instance.
            execution_context = ExecutionContext.create(credentials)
            extract_pdf_operation = ExtractPDFOperation.create_new()

            # Set operation input from a source file.
            source = FileRef.create_from_local_file(input_pdf)
            extract_pdf_operation.set_input(source)

            # Build ExtractPDF options and set them into the operation
            extract_pdf_options: ExtractPDFOptions = ExtractPDFOptions.builder() \
                .with_elements_to_extract([ExtractElementType.TEXT, ExtractElementType.TABLES]) \
                .with_element_to_extract_renditions(ExtractRenditionsElementType.TABLES) \
                .with_get_char_info(True) \
                .build()
            extract_pdf_operation.set_options(extract_pdf_options)

            # Execute the operation.
            result: FileRef = extract_pdf_operation.execute(execution_context)

            # Save the result to the specified location.  base_path + "/output/ExtractTextTableInfoWithCharBoundsFromPDF.zip"
            result.save_as(zip_file_2)
            
            print("Successfully extracted table information from PDF \n");
            
        except (ServiceApiException, ServiceUsageException, SdkException):
            logging.exception("Exception encountered while executing operation")
                    
        zip_path ="./ExtractTextTableInfoWithCharBoundsFromPDF_Final_Test.zip" #Opening the zipfile where the excel file containing the required data exists
        folder_name = 'tables/'
        
        if os.path.isdir('./'+folder_name):  #Deleting the folder where excel file will reside after extraction
            shutil.rmtree('./'+folder_name)

        archive = zipfile.ZipFile(zip_path, 'r')
        archive.extractall('.') #Extracting everything from the zipfile that contains data of the table
        archive.close() #Closing the file
        no=0
        while(True):  #Checking in which of the excel file the required data resides
            xlsx_filename = 'fileoutpart' + str(no) + '.xlsx'
            xlsx_path = './' + folder_name + xlsx_filename
            df = pd.read_excel(xlsx_path,header=None) #Reading the excel file
            if(df.iloc[0, 0][:-8] == "ITEM"): #We use slicing here as we get " _x000D_" in the end of all outputs of df.iloc[0, 0]
                break
            else:
                no+=2 #We are moving ahead by 2 as there is an image file after every excel file
        
        xlsx_filename = 'fileoutpart' + str(no+2) + '.xlsx' #Since we got the value of no such that the excel file only has headings of the table, we know no+2 will have the rquired data
        xlsx_path = './' + folder_name + xlsx_filename

        df = pd.read_excel(xlsx_path,header=None) #Reading the excel file            
        row_count = len(df) #Counting the number of rows in the file
        
        info=[None] * 19 #Creating a list which will later become the row to be inserted in the CSV file
        x=top=None #Will be requird lated in the for loop
        for element in data["elements"]:
            if ( "Text" in list(element) ):  #Checking if the element has text or not, as we only need the ones with text
    #We will check the bounds of most of the elements as the position of the text remains same across invoices of most of the data
                if(element["Path"].endswith("Sect/Title")):
                    info[3]=element["Text"] #Bussiness__Name
                elif(( element["Bounds"][0] == 76.72799682617188 ) and (element["Bounds"][3] == 717.5682373046875)):
                    info[1]=element["Text"] #"Bussiness__Country"
                elif(( element["Bounds"][0] == 76.72799682617188 ) and (element["Bounds"][3] == 704.2482452392578)):
                    info[5]=element["Text"] #"Bussiness__Zipcode"
                elif( (element["Bounds"][0] == 76.72799682617188 ) and (element["Bounds"][3] == 730.5582427978516) ): 
                    strin=element["Text"].split(",")
                    info[4]=strin[0] #"Bussiness__StreetAddress"
                    info[0]=strin[1] #"Bussiness__City"
                    if(element["Bounds"][1] == 708.1132049560547 ):
                        info[1]=strin[2]+ ", " +strin[3] #"Bussiness__Country"
                        

                elif(( element["Bounds"][0] != 76.72799682617188 ) and (element["Bounds"][3] == 730.5582427978516) ):
                    #Here we are assuming the issue date is included in the invoice number, in case it not true it will get overrided by the code written afterwards
                    info[16]=element["Text"][-11:] #"Invoice__IssueDate" We use the slicing to get rid of the word issue date written before the actual date
                    info[17]=element["Text"].split("Issue")[0][9:] #"Invoice__Number"


                elif((element["Bounds"][0] == 489.1699981689453 ) and (element["Bounds"][3] == 704.2482452392578)):
                    info[16]=element["Text"] #"Invoice__IssueDate"
                elif(( element["Bounds"][0] == 76.72799682617188 ) and (element["Bounds"][3] == 643.3882446289062)):
                    info[2]=element["Text"] #Bussiness__Description

                elif(( element["Bounds"][0] == 81.04800415039062 ) and (element["Bounds"][3] == 564.1382446289062)):
                    email=element["Text"].split(" ") #"Customer__Email"
                    info[8]=email[0]                 #"Customer__Email"
                    if(len(email)>=3 ): #Incase the email text extends to include the phone number as well 
                        if ('0' <= email[1][0] <= '9') : #Checking if its a phone number
                            info[10]=email[1]     #"Customer__PhoneNumber"
                        else: 
                            info[8]+=email[2]         #"Customer__Email addition : In case it is not the phone number it is the remaining part of the email, and will get added up
                            if(len(email)>=5): 
                                info[10]=email[4]  #"Customer__PhoneNumber"

                elif((element["Bounds"][0] == 412.8000030517578 ) and (element["Bounds"][3] == 577.1182403564453)):
                    info[15]=element["Text"][10:] #"Invoice__DueDate" Slicing required to get rid of the word "due date"

                elif(( element["Bounds"][0] == 240.25999450683594 ) and (element["Bounds"][3] == 577.1182403564453)):
                    info[14]=element["Text"] #"1 Invoice__Description"
                    #Incase the description runs across multiple lines it will get added
                elif(( element["Bounds"][0] == 240.25999450683594 ) and (element["Bounds"][3] == 564.1382446289062)):
                    info[14]+=element["Text"] #"2 Invoice__Description" 
                elif(( element["Bounds"][0] == 240.25999450683594 ) and (element["Bounds"][3] == 550.8182373046875)):
                    info[14]+=element["Text"] #"3 Invoice__Description"
                elif(( element["Bounds"][0] == 240.25999450683594 ) and (element["Bounds"][3] == 537.4982452392578)):
                    info[14]+=element["Text"] #"4 Invoice__Description"
                elif(( element["Bounds"][0] == 240.25999450683594 ) and (element["Bounds"][3] == 524.1782379150391)):
                    info[14]+=element["Text"] #"5 Invoice__Description"
                elif(( element["Bounds"][0] == 240.25999450683594 ) and (element["Bounds"][3] == 511.1982421875)):
                    info[14]+=element["Text"] #"6 Invoice__Description"

                elif(element["Text"].startswith("Tax %")) : #finding the position of Tax % which will be used later to extract the tax number
                    tax=element["Text"].split(" ")
                    if(len(tax)>3):
                        info[18]= tax[2]   #Invoice__Tax
                    else:
                        x=element["Bounds"][0]
                        top=element["Bounds"][3]
                elif(( element["Bounds"][0] != x ) and (element["Bounds"][3] == top) ):
                    info[18]=element["Text"] #Invoice__Tax 
                    
                #Passage under BILL TO Section
                elif(( element["Bounds"][0] == 81.04800415039062 ) and (element["Bounds"][3] == 577.1182403564453) ): #Line1****
                    name=element["Text"].split(" ") #Incase the the customer name extends across multiple lines to include other data as well 
                    info[9]= name[0] + " " + name[1] #"Customer__Name"
                    ln=len(name)
                    if(ln>3):
                        info[8]=name[2]          #"Customer__Email"
                    if(ln>4):
                        if('0'<=name[3][0]<='9'):
                            info[10]=name[3] #"Customer__PhoneNumber"
                        else:
                            info[8]+=name[3]
                    if(ln>5):
                        if('0'<=name[4][0]<='9'):
                            info[10]=name[4] #"Customer__PhoneNumber"
                            if(ln>6):
                                info[6]=name[5]+" " +name[6]+" " +name[7]  #"Customer__Address__line1"
                            if(ln>9):
                                info[7]=name[8]                            #"Customer__Address__line2"
                                if(ln>10):
                                    info[7]+=" "+name[9]
                        else:
                            info[6]=name[4]+" " +name[5]+" " +name[6]     #"Customer__Address__line1"


                elif(( element["Bounds"][0] == 81.04800415039062 ) and (element["Bounds"][3] ==564.1382446289062)):   #Line2****
                    mail=element["Text"].split(" ")        #"Customer__Email"
                    info[8]=mail[0]
                    ln=len(mail)
                    if(ln>2):
                        if ('0' <= mail[1] <='9') :      
                            info[10]=mail[1]          #"Customer__PhoneNumber"
                        else: 
                            info[8]+=mail[1]         #"Customer__Email"
                            if(ln>3):
                                info[10]=mail[2]         #"Customer__PhoneNumber"


                elif(( element["Bounds"][0] == 81.04800415039062 ) and (element["Bounds"][3] == 550.8182373046875)):  #Line3****
                    if ('0' <= element["Text"][0] <= '9') :      
                        info[10]=element["Text"]           #"Customer__PhoneNumber"
                    else: 
                        info[8]+=element["Text"]         #"Customer__Email"

                elif(( element["Bounds"][0] == 81.04800415039062 ) and (element["Bounds"][3] == 537.4982452392578) ): #Line4****
                    #Here, we are using -2 as the index to check if the last character is a number (and not -1) as the last character is a space
                    if( ('0' <= element["Text"][0] <= '9') and ('0' <= element["Text"][-2] <= '9')): #Checking if both first and last character is a number
                        info[10]=element["Text"] #"Customer__PhoneNumber"
                    else: 
                        ad=element["Text"].split(" ")
                        if(len(ad)>=4):
                            info[6]=ad[0] + " " + ad[1] + " " + ad[2] #"Customer__Address__line1"
                        if(len(ad)>=5):        
                            info[7]=ad[3]              #"Customer__Address__line2
                            if(len(ad)>=6):
                                info[7]+=" "+ ad[4]      #"Customer__Address__line2

                elif((element["Bounds"][0] == 81.04800415039062 ) and (element["Bounds"][3] == 524.1782379150391)): #Line5****
                    if('0' <= element["Text"][0] <= '9'):#Checking if its address line 1 or 2
                        addr=element["Text"].split(" ")
                        la=len(addr)
                        if(la>=3):
                            info[6]=addr[0]+" "+ addr[1]+" "+ addr[2]
                        if(la>=5):
                            info[7]=addr[3] #"Customer__Address__line2"
                            if(la>=6):
                                info[7]+=" "+addr[4]
                    else:
                        info[7]=element["Text"]
                elif((element["Bounds"][0] == 81.04800415039062 ) and (element["Bounds"][3] == 511.1982421875)): #Line6****
                    info[7]=element["Text"]

                
        for i in range(row_count): #We will run a loop to update the table info and keep inserting the updated info in the file in the loop as rest of the details remain same in a particular invoice
            #Completing the row information by taking the details of the table
            info[11] = df.iloc[i, 0][:-8] #"Invoice__BillDetails__Name
            info[12] = df.iloc[i, 1][:-8] #"Invoice__BillDetails__Quantity"
            info[13] = df.iloc[i, 2][:-8] #"Invoice__BillDetails__Rate"   

            file = open('ExtractedDataByMe_Final_Test_File.csv', 'a', newline='') #Opening the file to insert the row into CSV file
            writer = csv.writer(file)
            writer.writerow(info)
            file.close() #Closing the file
 

def browse_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        entry_pdf.delete(0, tk.END)  # Clear the current text in the Entry widget
        entry_pdf.insert(0, file_path)  # Insert the selected file path into the Entry widget
    process(file_path)

root = tk.Tk()
root.title("PDF File Name Input")

label_instruction = tk.Label(root, text="Enter the PDF file name:")
label_instruction.pack(pady=10)

entry_pdf = tk.Entry(root, width=50)
entry_pdf.pack(pady=5)

button_browse = tk.Button(root, text="Browse PDF", command=browse_pdf)
button_browse.pack(pady=10)




root.mainloop()
