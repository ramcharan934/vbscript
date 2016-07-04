option explicit
dim pin,n,reqpin
pin="1234"
n=3

do
	reqpin=inputbox("enter you pin")
	n=n-1
	if(pin=reqpin) then
		msgbox("you have successfully entered correct pin")
		exit do
	elseif(n=0) then
		msgbox("your card is blocked")
	else
		msgbox("wrong pin you have "&n &" chances")
	end if
loop while(n>0)
