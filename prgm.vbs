option explicit

dim a,b,d

a=cint(inputbox("enter value for a"))
b=cint(inputbox("enter value for b"))
msgbox (a)
msgbox (b)
d=inputbox("enter + for add " &vbcr&"enter - for subtraction")

if(d="+") then 
msgbox ("value of +  "&(a+b))

elseif(d="-") then
msgbox ("value for - " &(a-b))

elseif(d="/") then
msgbox("value for / "&(a/b))

else
msgbox "enter valid d value"
end if