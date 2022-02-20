Attribute VB_Name = "RFL"
Sub RFL()
Dim filepath As String
Dim fileName As String
filepath = "\\qctdfsrt\prj\vlsi\pete\ptetools\prod\tss\runcommands\userFiles\"
outfldr = "\\qctdfsrt\prj\vlsi\pete\ptetools\prod\tss\runcommands\userOutputs\"
fileName = Environ$("Username") & "_RFL_cmd_" & Format(Now, "mmddyyyy_HHmmss")
SaveAsTxtFile filepath, fileName, "tssupdate3 -h #testing_prj\n#end_of_command"
 FContent = LoadFileStr(outfldr & fileName & ".out")
 MsgBox FContent
 
End Sub

Sub SaveAsTxtFile(filepath As String, fileName As String, lineText As String)
    Set fs = CreateObject("scripting.filesystemobject")
    If Not fs.FolderExists(filepath) Then MsgBox ("Unable to access the folder " & fileName)
    Dim myrng As Range
    Open filepath & fileName For Output As #1
    Print #1, lineText
    Close #1
    If Not fs.FileExists(filepath & fileName) Then MsgBox ("Unable to create the file " & fileName)
End Sub

Function LoadFileStr$(FN$)
    With CreateObject("Scripting.FileSystemObject")
          LoadFileStr = .OpenTextFile(FN, 1).readall
    End With

End Function
Function FileExists(fileName As String) As Boolean

        On Error GoTo NotExist
        If Not Len(Dir(fileName)) = 0 Then
            FileExists = True
            On Error GoTo 0
            Exit Function
        End If

NotExist:
        On Error GoTo 0
        FileExists = False
End Function ' FileExists

