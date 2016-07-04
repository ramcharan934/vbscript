option explicit
dim fso,n,s,i,j,mark,m,flag,avg,name,opentxtfile
set fso=createobject("scripting.filesystemobject")
n=cint(inputbox("enter no. of students: "))
s=cint(inputbox("enter no. of subjects: "))

if(fso.folderexists("c:\result")) then
else
	msgbox "no folder present"
	fso.createfolder("c:\result")
	msgbox "folder created"
	end if

for i=1 to n
	m=0
	flag=true
	name=inputbox("enter the student name: ")
	for j=1 to s 
		mark= cint(inputbox("enter the marks: "))
		m=m+mark
		call calfail(m,name)
		
	next
	avg=m/s
	if(flag)then
			call calavg(avg,name)
	else
			call writefail(name)
	end if 
next
function calfail(byval mark,byval name)
	if(mark<40) then 
		flag=false
	end if
end function
function calavg(byval avg,byval name)
	if(avg>=80) then
		msgbox "grade 1: "&avg&" "&name
		writegradeone name,avg
	elseif(avg<80 and avg>=65)then
		msgbox "grade 2: "&avg&" "&name
		writegradetwo name,avg
	elseif(avg>60 and avg<65) then
		msgbox "grade 3: "&avg&" "&name
		writegradethree name,avg
	else
	end if
end function
function writefail(byval name)
		dim filename
		filename="c:\result\failed.txt"
		if(fso.fileexists(filename))then
			set opentxtfile = fso.opentextfile(filename,8)
			opentxtfile.writeline(name)
			opentxtfile.close()
		else
			fso.createtextfile(filename)
			msgbox "created a file"
			set opentxtfile = fso.opentextfile(filename,8)
			opentxtfile.writeline(name)
			opentxtfile.close()
		end if
end function
function writegradeone(byval name,byval avg)
		dim filename
		filename="c:\result\grade1.txt"
	if(fso.fileexists(filename)) then
		set opentxtfile=fso.opentextfile(filename,8)
		opentxtfile.write(name)
		opentxtfile.write("- ")
		opentxtfile.writeline(avg)
		opentxtfile.close()
	else
		fso.createtextfile(filename)
		msgbox "created a file"
		set opentxtfile=fso.opentextfile(filename,8)
		opentxtfile.write(name)
		opentxtfile.write("-")
		opentxtfile.writeline(avg)
		opentxtfile.close()
	end if
end function
function writegradetwo(byval name,byval avg)
		dim filename
		filename="c:\result\grade2.txt"
	if(fso.fileexists(filename)) then
		set opentxtfile=fso.opentextfile(filename,8)
		opentxtfile.write(name)
		opentxtfile.write("- ")
		opentxtfile.writeline(avg)
		opentxtfile.close()
	else
		fso.createtextfile(filename)
		msgbox "created a file"
		set opentxtfile=fso.opentextfile(filename,8)
		opentxtfile.write(name)
		opentxtfile.write("-")
		opentxtfile.writeline(avg)
		opentxtfile.close()
	end if
end function
function writegradethree(byval name,byval avg)
		dim filename
		filename="c:\result\grade3.txt"
	if(fso.fileexists(filename)) then
		set opentxtfile=fso.opentextfile(filename,8)
		opentxtfile.write(name)
		opentxtfile.write("- ")
		opentxtfile.writeline(avg)
		opentxtfile.close()
	else
		fso.createtextfile(filename)
		msgbox "created a file"
		set opentxtfile=fso.opentextfile(filename,8)
		opentxtfile.write(name)
		opentxtfile.write("-")
		opentxtfile.writeline(avg)
		opentxtfile.close()
	end if
end function


	