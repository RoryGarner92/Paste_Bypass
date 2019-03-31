'VBScript ForcePaste
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objHTMl = CreateObject("htmlfile")

WshShell.SendKeys "%{TAB}"
WScript.Sleep 600

copyContent = objHTML.ParentWindow.ClipboardData.GetData("text")
WshShell.SendKeys copyContent 

