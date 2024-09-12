
import win32com.client
import pythoncom
acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument
msp = doc.ModelSpace
psp = doc.PaperSpace
cl = win32com.client
def vtpnt(x,y,z=0):
    return cl.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,(x,y,z))

def vtobj(obj):
    return cl.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH,obj)

def vtFloat(list):
    return cl.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list)

def vtInt(list):
    return cl.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, list)
def vtVariant(list):
    return cl.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, list)
try:
    doc.SelectionSets.Item("SS1").Delete()
except:
    print("Delete selection failed")
slt = doc.SelectionSets.Add('SS1')
filterType=[-4,0,-4]
filterData=["<And","Line","And>"]
filterType=vtInt(filterType)
filterData=vtVariant(filterData)
slt.SelectOnScreen(filterType, filterData)
for obj in slt:
    print("Selected object:", obj.ObjectName)
