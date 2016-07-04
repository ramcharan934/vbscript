option explicit
dim a,b,count,op,start,n,c
start=inputbox("do you want to perform yes or no")
if (start="yes") then
	count=cint(inputbox("how many times"))
	n=1
	while(n <= count )
		a=cint(inputbox("enter value for a"))
		b=cint(inputbox("enter value for b"))
		op=inputbox("enter + for add " &vbcr&"enter - for subtraction"&vbcr&"/ for div")
		select case op
		case "+"
		c=a+b
		msgbox("add value is "&c)
		case "-"
		c=a-b
		msgbox("sub value is "&c)
		case "/"
		c=a/b
		msgbox("div value is "&c)
		end select
		n=n+1
	wend
	
end if