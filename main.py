import openpyxl
from tika import parser
import os, sys
import re

                        #Cycle through each file in Tablet Scan Subfolder (PDF Versions Only)
path = "Z:/!USERS/Sean/Tablet-Test"
dirs = os.listdir(path)
lb = "------------------------------------------------------------------------------------------------------------------"

for file in dirs:
                        #Within this block, any code is applied to each file in Test Folder
    pdfContent = parser.from_file("Z:/!USERS/Sean/Tablet-Test/" + file)     #Parses PDF document
    tsContent = (pdfContent['content'])                                     #Returns only the CONTENT of the PDF

    #REGEX Searches ----------------------------------------------------------------------------------------------------
    productSku = re.search(r'(CD..[0-9][0-9][0-9][0-9])', tsContent)          #Matches any product that matches CDXX####
    if productSku == None:
        productSku = re.search(r'(CD..[0-9][0-9][0-9][0-9])\w*', tsContent)  # Matches any product that matches CDXX########

    productHeight = re.search(r'H\s([0-9]+)\s([0-9]+)', tsContent)  #Group 1 = Product Size, Group 2 = Box Size
    if productHeight == None:                                                #Error Loop, If No Box Size Found
        productHeight = re.search(r'H\s([0-9]+)', tsContent)

    productWidth = re.search(r'W\s([0-9]+)\s([0-9]+)', tsContent)   #Group 1 = Product Size, Group 2 = Box Size
    if productWidth == None:                                                 #Error Loop, If No Box Size Found
        productWidth = re.search(r'W\s([0-9]+)', tsContent)

    productDepth = re.search(r'D\s([0-9]+)\s([0-9]+)', tsContent)   #Group 1 = Product Size, Group 2 = Box Size
    if productDepth == None:                                                 #Error Loop, If No Box Size Found
        productDepth = re.search(r'D\s([0-9]+)', tsContent)

    productWeight = re.search(r'LBS\s([0-9]+.[0-9]\s)([0-9]+)', tsContent)  #Group 1 = Product Weight, Group 2 = Box Weight
    if productWeight == None:
        productWeight = re.search('LBS\s([0-9]+)\s([0-9])', tsContent)                 #Error Loop 1/2, if NO DECIMAL in weight.
        if productWeight == None:
            productWeight = re.search('LBS\s([0-9]+)', tsContent)                      #Error Loop 2/2, if NO Box Weight
            if productWeight == None:
                productWeight = "--"

    productDesc = re.search(r'ITEM\s#\n(\n[\w+\s]+)', tsContent)            #Find Product Description on Tablet Scan

    productPrice = re.search(r'TOTAL:\s([0-9]+.[0-9][0-9])\$*\s*', tsContent) #Find Product Total Price on Tablet Scan


    #REGEX Search Cleaning ---------------------------------------------------------------------------------------------
    try:
        cleanProductDesc = productDesc.group().split(productSku.group())        #Cut String off when SKU is found in string
        cleanProductDesc[0] = cleanProductDesc[0].lstrip("ITEM #")              #Remove 'ITEM #' from start of string
        cleanProductDesc[0] = cleanProductDesc[0].strip('\n')                   #Remove any '\n'  from string
        cleanProductDesc[0] = cleanProductDesc[0].replace('\n', "")             #Remove any '\n' from center of string
    except AttributeError:
        cleanProductDesc = "No Product Description"

    #Test Output -------------------------------------------------------------------------------------------------------
    #print(tsContent)                                                        #Prints PDF Content
    print(lb)                                                               #Prints Line Break
    if productSku:
        print(productSku.group(0))

    try:
        print(cleanProductDesc[0])
    except AttributeError:
        print(cleanProductDesc)

    print(lb)

    if productHeight:
        print(productHeight.group(1))

    if productWidth:
        print(productWidth.group(1))

    if productDepth:
        print(productDepth.group(1))

    print(lb)

    if productWeight:
        print(productWeight.group(1))

    #    productDesc and productHeight and productWidth and productDepth:
     #   print(productSku.group(0) + '\n' + cleanProductDesc[0] + '\n' + lb       #Prints First Block: SKU, Description
      #      + '\n' + productHeight.group(1) + '\n' + productWidth.group(1)    #Prints Second Block: Height/Width/Depth
       #     + '\n' + productDepth.group(1) + '\n' + productWeight.group(1))   #Weight
    #elif productSku and productHeight and productWidth and productDepth:
     #   print(productSku.group(0) + '\n' + cleanProductDesc + '\n' + lb  # Prints First Block: SKU, Description
      #      + '\n' + productHeight.group(1) + '\n' + productWidth.group(1)  # Prints Second Block: Height/Width/Depth
       #     + '\n' + productDepth.group(1) + '\n' + productWeight.group(1))  # Weight

    print(lb)                                                               #Prints Line Break

    try:
        print(productHeight.group(2))
    except IndexError:
        print('--')

    try:
        print(productWidth.group(2))
    except IndexError:
        print('--')

    try:
        print(productDepth.group(2))
    except IndexError:
        print('--')

    try:
        print(productWeight.group(2))
    except IndexError:
        print('--')
    except AttributeError:
        print('--')
    print(lb)                                                               #Prints Line Break

    if productPrice:
        print(productPrice.group(1))                                        #Prints Product Price
    else:
        productPrice = re.search(r'\n([0-9]+.[0-9]+)\$\s*\n\nI', tsContent) #FAILSAFE SEARCH TO BE ADDED TO REGEX SEARCH
        print(productPrice.group(1))
