' Skyqqr
' 2020/7
set shell = createobject ("wscript.shell")
strtimes = inputbox("Amount of messages to send")
strtdelay = inputbox("Delay between messages")
strtnum = inputbox("Set Number to count to")
if not isnumeric(strtimes) then
wscript.quit
end if
if not isnumeric(strtnum) then
wscript.quit
end if
if not isnumeric(strttimes) then
wscript.quit
end if
if not isnumeric(strtdelay) then
wscript.quit
end if
' Checks if input is valid
msgbox "You have 5 seconds to get to your inputbox."
wscript.sleep( 5000 )
for i=1 to strtimes
strtnum=strtnum+1
shell.sendkeys(strtnum)
shell.sendkeys("~")
' after writing the value of strtnum, sends an enter key
wscript.sleep(strdelay)
next