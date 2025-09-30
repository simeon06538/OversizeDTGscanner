import win32com.client
import pythoncom
import os 
import pandas as pd 
def scan_barcode():
  print("Ready for the barcode scan:")
  barcode_data = input("Scan barcode").strip()
  return barcode_data
def inkbay_csv(inkbay_id, csv_path):
  try:
    order = pd.read_csv(csv_path)
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
  print("yo")
#def inkbay_order(inkbay_id, order_data):

        


#bpac = win32com.client.Dispatch("bpac.Document")
#if bpac.Open(r"C:\Users\simeo\Downloads\Brother bPAC3 SDK\Templates\Item.lbx"):
  #obj = bpac.GetObject("objText")
  #if obj:
    #obj.Text = "Hello Brother!"

#bpac.StartPrint("",0)
#bpac.PrintOut(1,0)

