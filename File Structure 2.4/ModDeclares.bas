Attribute VB_Name = "ModDeclares"
Declare Function XSetFileName Lib "FileStructure.dll" (File As String) As Boolean
Declare Function XAppend Lib "FileStructure.dll" (Data As String) As Boolean
Declare Function XDeleteLine Lib "FileStructure.dll" (LineStartsWith As String) As Boolean
Declare Function XAddToLine Lib "FileStructure.dll" (LineStartsWith As String, Txt2Append As String) As Boolean
Declare Function XInfo Lib "FileStructure.dll" ()
