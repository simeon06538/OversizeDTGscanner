import win32com.client
import pythoncom
import os 
import pandas as pd 
import re

def scan_barcode():
  print("Ready for the barcode scan:")
  barcode_data = input("Scan barcode:")
  return barcode_data.strip()

#

def extract_inkbay_id(options_text):
  if pd.isna(options_text) or not isinstance(options_text, str):
    return None
  
  match = re.search('inkybay customization id:\s*(\d+)', options_text, re.IGNORECASE)

  if match:
    return match.group(1)
  return None

#

def inkbay_csv(inkbay_id, csv_path):
  try:
    order = pd.read_csv(csv_path)
    order_data = order[
      (order['Item-Options'] == inkbay_id)
    ]
    matching_rows = []
    for index, row in order.iterrows():
      options_text = row['Item - Options']
      extracted_id = extract_inkbay_id(options_text)

      if extracted_id == inkbay_id:
        matching_rows.append(row)
        print(f"Found matching order!")

        return {
          'order-number' : row.get('Order - Number', 'N/A')
        }
        
    if not matching_rows:
      print(f"No order found with InkBayId: {inkbay_id}")  

  except Exception as e:
      print(f"unable to read CSV, error {e}")
  except FileNotFoundError:
    print(f"CSV file not found :")
    return None; 

def print_with_bpac(template_path, asset_data):
  try: 
    pythoncom.CoInitialize()
    ObjDoc = win32com.client.Dispatch("bpac.Document")
    bRet = ObjDoc.Open(template_path)

    if bRet:
      print("Template opened successfully")
      ObjDoc.GetObject("Name").Text=asset_data.get("Name", "")
      ObjDoc.GetObject("Section").Text= asset_data.get("Section", "")
      ObjDoc.GetObject("Number").Text= asset_data.get("Number", "")

      ObjDoc.StartPrint("",0)
      ObjDoc.PrintOut(1,0)
      ObjDoc.EndPrint()
      ObjDoc.Close()

      print("job sent successfully")
    
    else:
      print("Failed to open template: " + template_path)

  except Exception as e:
    print("Error " + e)
  finally:
    pythoncom.CoUninitialize()

def main():
  barcode_data = scan_barcode()
  csv_path = r'C:\Users\simeo\Downloads\oversizeDTGscanner\e7437335-152b-422e-b35a-eb26543b99f9(4).csv'
  order_data = inkbay_csv(barcode_data, csv_path)
  
#def inkbay_order(inkbay_id, order_data):

        


#bpac = win32com.client.Dispatch("bpac.Document")
#if bpac.Open(r"C:\Users\simeo\Downloads\Brother bPAC3 SDK\Templates\Item.lbx"):
  #obj = bpac.GetObject("objText")
  #if obj:
    #obj.Text = "Hello Brother!"

#bpac.StartPrint("",0)
#bpac.PrintOut(1,0)

