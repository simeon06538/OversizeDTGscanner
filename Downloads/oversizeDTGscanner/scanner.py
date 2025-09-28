import win32com.client

bpac = win32com.client.Dispatch("bpac.Document")
if bpac.Open(r"C:\Users\simeo\Downloads\Brother bPAC3 SDK\Templates\Item.lbx"):
  obj = bpac.GetObject("objText")
  if obj:
    obj.Text = "Hello Brother!"

bpac.StartPrint("",0)
bpac.PrintOut(1,0)
 
