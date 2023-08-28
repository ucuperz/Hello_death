Set WshShell = CreateObject("WScript.Shell")
'Run the hta.
WshShell.Run "Test.hta", 1, true
'Display the results.
MsgBox "Return Value = " & getReturn
Set WshShell = Nothing

Function getReturn
'Read the registry entry created by the hta.
On Error Resume Next
     Set WshShell = CreateObject("WScript.Shell")
    getReturn = WshShell.RegRead("HKEY_CURRENT_USER\Volatile Environment\MsgResp")
    If ERR.Number  0 Then
        'If the value does not exist return -1
         getReturn = -1
    Else
        'Otherwise return the value in the registry & delete the temperary entry.
        WshShell.RegDelete "HKEY_CURRENT_USER\Volatile Environment\MsgResp"
    End if
    Set WshShell = Nothing
End Function