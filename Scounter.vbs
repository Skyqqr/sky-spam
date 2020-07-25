set shell = createobject ("wscript.shell")
strtimes = inputbox("Amount of messages to spam")
strtdelay = inputbox("Delay between messages")
cnum = inputbox("Set Cnum")
if not isnumeric(strtimes) then
wscript.quit
end if
msgbox "You have 5 seconds to get to your inputbox."
wscript.sleep( 5000 )
for i=1 to strtimes
cnum=cnum+1
shell.sendkeys(cnum)
shell.sendkeys("~")
wscript.sleep(strdelay)
next