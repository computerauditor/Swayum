import time
import pyautogui
import openpyxl 
import pyperclip
import tkinter as tk
from tkinter import messagebox

#-----------------WARING----------------------------
root = tk.Tk()
root.withdraw()  
message = "Please ensure you are at the SAP Menu.\nClose all running Excel files before proceeding.\nDon't touch the code untill you know what you are doing\n\n Made By Devyansh Rastogi"
messagebox.showwarning("Warning", message)
root.destroy()
#-----------------------------------------------------------------
# Function to read doc numbers from a the .txt file
def read_doc_numbers(file_path):
    with open(file_path, 'r') as file:
        doc_numbers = file.readlines()
    # Remove whitespace and newline characters
    doc_numbers = [doc.strip() for doc in doc_numbers]
    return doc_numbers
#-----------------------------------------------------------------
sap_menu_icon_location = (146,77)
pyautogui.click(sap_menu_icon_location)
time.sleep(3) 
pyautogui.typewrite('/nMIR4')
pyautogui.press('ENTER')
#-----------------------------------------------------------------
excel_file = "sap_details.xlsx"
Invoice_doc_no = (381,240)
pyautogui.click(sap_menu_icon_location)
time.sleep(0.5) 
doc_numbers = read_doc_numbers('docnum.txt')

wb = openpyxl.load_workbook("sap_details.xlsx")
ws = wb["Invoice Tracker"]

for i, doc_no in enumerate(doc_numbers, start=2):
    pyautogui.typewrite(doc_no)  # Corrected variable name
    pyautogui.press('ENTER')

    # Locate the area on the screen where Invoice Date is displayed
    
    # INVOICE DATE
    pyautogui.moveTo(376, 369, 1)
    time.sleep(1)
    pyautogui.click(376, 369)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    invoice_date = pyperclip.paste()
    time.sleep(0.2)
    
    # REFERENCE
    pyautogui.click(688, 368)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'c')
    invoice_number = pyperclip.paste()
    time.sleep(1)

    # VENDOR CODE
    pyautogui.click(316, 316) # GO TO DETAILS
    time.sleep(1)
    pyautogui.click(776, 440) # GO TO INV PARTY
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'c')
    vendor_number = pyperclip.paste()
    time.sleep(1)

    # VENDOR CODE
    pyautogui.click(789, 807) # GO TO PO NUMBER
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'c')
    PO_number = pyperclip.paste()
    time.sleep(1)
    
    # ITEM CODE
    pyautogui.doubleClick(789, 807)
    time.sleep(1)
    pyautogui.moveTo(402, 374)
    pyautogui.click(402, 374)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'c')
    item_Code = pyperclip.paste()

    # GRN QUANTITY

    pyautogui.moveTo(475,643)
    time.sleep(2)
    pyautogui.click(475,643)
    pyautogui.click(475,643)
    time.sleep(1)
    pyautogui.click(515,675)
    time.sleep(1)
    pyautogui.moveTo(229,379)
    pyautogui.click(229,379)
    time.sleep(1)
    pyautogui.click(366,643)
    time.sleep(1)
    pyautogui.click(406,741)
    time.sleep(8)
    pyautogui.click(709,604)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'f')
    pyautogui.typewrite('Tr./Ev. Goods receipt')
    pyautogui.press('ENTER')
    pyautogui.press('ENTER')
    for _ in range(10):
        pyautogui.press('TAB')
    pyautogui.press('ENTER')
    time.sleep(1)
    pyautogui.press('TAB')
    pyautogui.hotkey('ctrl', 'c')
    GRN_qty = pyperclip.paste()

    # INV_RECEIPT QUANTITY
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'f')
    pyautogui.typewrite('Tr./Ev. Invoice receipt')
    pyautogui.press('ENTER')
    pyautogui.press('ENTER')
    for _ in range(10):
        pyautogui.press('TAB')
    pyautogui.press('ENTER')
    time.sleep(1)
    pyautogui.press('TAB')
    pyautogui.hotkey('ctrl', 'c')
    InvoiceRec_qty = pyperclip.paste()    

    # PARK QUANTITY
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'f')
    pyautogui.typewrite('Tr./Ev. Parked invoice')
    pyautogui.press('ENTER')
    pyautogui.press('ENTER')
    for _ in range(10):
        pyautogui.press('TAB')
    pyautogui.press('ENTER')
    time.sleep(1)
    pyautogui.press('TAB')
    pyautogui.hotkey('ctrl', 'c')
    Park_qty = pyperclip.paste()   

    pyautogui.hotkey('alt', 'f4')
    pyautogui.press('TAB')
    time.sleep(1)
    pyautogui.press('ENTER')
    time.sleep(2)

    ws[f'D{i}'] = doc_no
    ws[f'F{i}'] = invoice_date
    ws[f'E{i}'] = invoice_number
    ws[f'H{i}'] = vendor_number
    ws[f'L{i}'] = PO_number
    ws[f'N{i}'] = item_Code
    ws[f'T{i}'] = GRN_qty
    ws[f'U{i}'] = InvoiceRec_qty
    ws[f'V{i}'] = Park_qty

    pyautogui.moveTo(168, 79)
    time.sleep(0.5)
    pyautogui.click(168, 79)    
    time.sleep(1)
    pyautogui.click(168, 79)
    pyautogui.typewrite('/nMIR4')
    pyautogui.press('ENTER')
    time.sleep(5)

# Save the Excel file
wb.save(excel_file)