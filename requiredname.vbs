option explicit
dim start,intake,e,str
start=inputbox("you like to enter type y")
intake=array()
while(start="y") 
	redim preserve intake(ubound(intake)+1)
	intake(ubound(intake))=inputbox("enter your value:")
	start=inputbox("you like to enter type y")
wend

for each e in intake
	str=str&" "&e
next
msgbox "you entered "&str