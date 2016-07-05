dim excel,wb,sh,rc,opentxtfile,filename,fso,str
set fso=createobject("scripting.filesystemobject")
set excel=createobject("excel.application")
msgbox typename(excel)
'excel.visible=true
'excel.workbooks.add()
'excel.activeworkbook.saveas("c:\my\names.xls")
'excel.visisble=false
set wb=excel.workbooks.open("c:\my\names.xls")
msgbox typename(wb)
set sh=wb.sheets("sheet1")
rc=sh.usedrange.rows.count

filename="c:\my\note4.txt"
if(fso.fileexists(filename)) then
	set opentxtfile=fso.opentextfile(filename,8)
else
	fso.createtextfile(filename)
	msgbox "created a file"
	set opentxtfile=fso.opentextfile(filename,8)
end if
for i=1 to rc
str=sh.cells(i,1).value&", "&sh.cells(i,2).value
msgbox str
opentxtfile.writeline(str)
next
opentxtfile.close()
wb.save()
wb.close()