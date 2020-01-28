'Medi cat automated Sign-In VBScript'
'Winston-Salem State University OIT Department'
'Created by: Samuel Chance II 11/8/2019'

Option Explicit
Dim sh

Set sh = CreateObject("WScript.shell")

CreateObject("WScript.Shell").Popup "Please wait for automated login to complete before interacting with this PC.", 3, "Notice!"
WScript.Sleep 12000 'wait 12 Seconds'


WScript.Sleep 5000 'wait for application to load'

sh.SendKeys "******" 'Enter User Name : replace * with user name '
sh.SendKeys "{TAB}"

sh.SendKeys "******" 'Enter Password : replace * with Password '
sh.SendKeys "{TAB}"

sh.SendKeys "{DOWN}" 'Arrow input'
sh.SendKeys "{DOWN}"
sh.SendKeys "{TAB}"

sh.SendKeys "~" 'ENTER' 'select Login'
