
dim fso,f,opentxtfile,temparray,tempstr,excel,path,wb,sh,filename,i,j
set fso = createobject("scripting.filesystemobject")

path="c:\my\names2.xls"
set excel=createobject("excel.application")
excel.workbooks.add()
excel.activeworkbook.saveas(path)
msgbox typename(excel)
set wb=excel.workbooks.open(path)
msgbox typename(wb)
set sh=wb.sheets("sheet1")
msgbox typename(sh)

filename="c:\my\note4.txt"
if(fso.fileexists(filename))then 
	set opentxtfile = fso.opentextfile(filename,1)
	msgbox "text opened"
else
	msgbox "no input txt fileexists"
end if

i=1
do
	tempstr= opentxtfile.readline()
	temparray=split(tempstr,",")
	j=1
	for each temp in temparray
		sh.cells(i,j).value=temp
		j=j+1
	next
	i=i+1	
loop until(opentxtfile.atendofstream)

opentxtfile.close()
wb.save()
wb.close()
	