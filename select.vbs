option explicit
dim a,b,op,c

a=cint(inputbox("enter the value for a"))

b=cint(inputbox("enter the value for b"))

op=inputbox("enter op"&vbcr&"+ for add"&vbcr&"- for sub"&vbcr&"/ for div")
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

