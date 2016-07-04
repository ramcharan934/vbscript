option explicit
dim a,l,u,i,e
a=array("a","b",3,"d")
msgbox typename(a)
l=lbound(a)
msgbox "lower bond value is"&l
u=ubound(a)
msgbox "upper bond value is"&u

for i=0
	 to u 
	msgbox a(i)
	next
	redim preserve a(4)
	a(4)="g"
	for each e in a 
	msgbox e
next
