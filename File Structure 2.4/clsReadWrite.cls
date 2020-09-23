VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReadWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private FileName As String

Public Function XSetFileName(File As String)
    FileName = File
End Function

Public Function XAppend(Data As String)
    Dim Num As Integer
    Num = FreeFile
    
    Open FileName For Append As Num
        Print #Num, Data
    Close Num
End Function

Public Function XDeleteLine(LineStartsWith As String) As Boolean
    On Error GoTo err1
    Dim MsgResult As String
    Dim TempInput As String
    Dim TheRest As String
    Dim Perm As String
    
    
    If Trim(FileName) = "" Then MsgBox "Please use the command 'XSetFilename' to specify a file to delete from before executing any other commands in this DLL.", vbCritical, "Error!": GoTo err1
    
    Open FileName For Input As #1
    Open App.Path & "TempX123.txt" For Output As #2
        Do Until EOF(1)
            Line Input #1, TempInput
            If LineStartsWith = Mid(TempInput, 1, Len(LineStartsWith)) Then GoTo FoundIt2
            If EOF(1) Then GoTo err1
            Print #2, TempInput
        Loop
FoundIt2:
    If EOF(1) = False Then Line Input #1, TheRest
        Do Until EOF(1)
            Line Input #1, TheRest
            Print #2, TheRest
        Loop
    Close #1
    Close #2
    Kill FileName
    Open FileName For Output As #1
    Open App.Path & "\TempX123.txt" For Input As #2
        Do Until EOF(2)
            Line Input #2, TempInput
            Print #1, TempInput
        Loop
    Close #1
    Close #2
    Kill App.Path & "\TempX123.txt"

    XDeleteLine = True
    Exit Function
err1:
    Err.Clear
    On Error Resume Next
    Close #1
    Close #2
    XDeleteLine = False
End Function

Public Function XAddToLine(LineStartsWith As String, Txt2Append As String) As Boolean
    On Error GoTo err1
    Dim MsgResult As String
    Dim TempInput As String
    Dim TheRest As String
    Dim Perm As String
    
    
    If Trim(FileName) = "" Then MsgBox "Please use the command 'XSetFilename' to specify a file to delete from before executing any other commands in this DLL.", vbCritical, "Error!": GoTo err1
    
    Open FileName For Input As #1
    Open App.Path & "TempX1234.txt" For Output As #2
        Do Until EOF(1)
            Line Input #1, TempInput
            If LineStartsWith = Mid(TempInput, 1, Len(LineStartsWith)) Then GoTo FoundIt2
            If EOF(1) Then GoTo err1
            Print #2, TempInput
        Loop
FoundIt2:
        Print #2, TempInput & Txt2Append
    If EOF(1) = False Then Line Input #1, TheRest
        Do Until EOF(1)
            Line Input #1, TheRest
            Print #2, TheRest
        Loop
    Close #1
    Close #2
    Kill FileName
    Open FileName For Output As #1
    Open App.Path & "\TempX1234.txt" For Input As #2
        Do Until EOF(2)
            Line Input #2, TempInput
            Print #1, TempInput
        Loop
    Close #1
    Close #2
    Kill App.Path & "\TempX1234.txt"

    XAddToLine = True
    Exit Function
err1:
    Err.Clear
    On Error Resume Next
    Close #1
    Close #2
    XAddToLine = False
End Function

Public Function XInfo()
    Dim Data As String
    
    Data = "File Structure DLL 2.4" & vbCrLf & vbCrLf
    Data = Data & "Programmed By: Harry Maugans" & vbCrLf
    Data = Data & "-----Current Commands-----" & vbCrLf
    Data = Data & "XInfo: This current function, displays a msgbox of the help and command list for File Structure 2.4" & vbCrLf
    Data = Data & "XSetFileName(File As String): Sets the path & filename to be edited." & vbCrLf
    Data = Data & "XAppend(Data As String): Add a line of text (data) to the end (the last line) of a file." & vbCrLf
    Data = Data & "XDeleteLine(LineStartsWith As String) As Boolean: Finds a line in the file that starts with the variable passed and deletes it." & vbCrLf
    Data = Data & "    It returns False if the line could not be found that starts with the text specified." & vbCrLf
    Data = Data & "XAddToLine(LineStartsWith As String, Text2Append As String) As Boolean:  Finds the line in the file that starts with the specified text" & vbCrLf
    Data = Data & "    and adds the value of the variable Text2Append to the end of that line."
    
    MsgBox Data, vbInformation, "File Structure 2.4 Info File"
    
End Function
