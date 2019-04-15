
Option Explicit

dim ComputerName
dim WMIService, NIC
dim ConnectionName



ComputerName = "."
ConnectionName = "Ethernet"


Set WshShell = WScript.CreateObject("WScript.Shell")
If WScript.Arguments.Length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe" _
    , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
  WScript.Quit
End if
if not IsElevated then 
    WScript.Echo "Please run this script with administrative rights!"
    WScript.Quit
end if


On Error Resume Next

Set WMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & ComputerName & "\root\cimv2")
Set NIC = WMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID = '" & ConnectionName & "'").ItemIndex(0)

if NIC is nothing then
    WScript.Echo "NIC not found!"
    WScript.quit
end if

WScript.Echo "NIC Current status: " & iif(NIC.NetEnabled, "Enabled", "Disabled") & vbcrlf

if NIC.NetEnabled then
    WScript.Echo "Disabling NIC..."
    NIC.Disable
    WScript.Sleep 1000
    WScript.Echo "Enabling NIC..."
    NIC.Enable
end if

function iif(cond, truepart, falsepart)
    if cond then iif=truepart else cond=falsepart
end function

function IsElevated()

    On Error Resume Next
    CreateObject("WScript.Shell").RegRead("HKEY_USERS\s-1-5-19\")
    IsElevated = (err.number = 0)

end function