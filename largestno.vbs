option explicit
dim a,b,c

a=cint(inputbox("enter the value of a"))
b=cint(inputbox("enter the value of b"))
c=cint(inputbox("enter the value of c"))

if(a>b and a>c)then
	msgbox(a&" : is the greatest")
elseif(b>a and b>c)then
	msgbox(b&" : is the greatest")
else
	msgbox(c&" : is the greatest")
end if