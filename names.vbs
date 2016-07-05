dim excel,m,n

set excel=createobject("excel.application")
msgbox typename(excel)
excel.visible=true
excel.workbooks.add()
excel.activeworkbook.saveas("c:\my\names.xls")
'excel.visisble=false
set wb=excel.workbooks.open("c:\my\names.xls")
msgbox typename(wb)
set sh=wb.sheets("sheet1")
msgbox typename(sh)
m=5
for n=1 to m
sh.cells(n,1).value=inputbox("enter name of person: "&n)
sh.cells(n,2).value=inputbox("enter mobile no.of person: "&n)	
next

wb.save()
wb.close()