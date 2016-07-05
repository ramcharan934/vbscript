dim excel

set excel=createobject("excel.application")
msgbox typename(excel)
excel.visible=true
excel.workbooks.add()
excel.activeworkbook.saveas("c:\my\new1.xls")
'excel.visisble=false
set wb=excel.workbooks.open("c:\my\new1.xls")
msgbox typename(wb)
set sh=wb.sheets("sheet1")
msgbox typename(sh)
sh.cells(1,1).value="hello"
wb.save()
wb.close()
excel.visible=false
'remove the object reference from the varible
set sh=nothing
set wb=nothing
set excel=nothing