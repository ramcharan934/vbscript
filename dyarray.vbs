option explicit
dim a,l,u,i
a=array("a","b",3,"d")
msgbox typename(a)
l=lbound(a)
msgbox "lower bond value is"&l
u=ubound(a)
msgbox "upper bond value is"&u

for i=1 to u 
msgbox a(i)
next