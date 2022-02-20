Attribute VB_Name = "Savetotss"
 
 
'https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a
'parser using regex
Public Const RFLpath = "\\qctdfsrt\prj\vlsi\pete\ptetools\prod\tss\runcommands\"
Public Const RFLUnixPath = "/prj/vlsi/pete/ptetools/prod/tss/runcommands/"
Public Const RFLuser = "userFiles"
Public Const RFLSupport = "supportFiles"
Dim Data_file As String
Dim cmd_file As String
Dim block As String

Sub SendCToTSS()
    Application.ScreenUpdating = False
   DoEvents
    UserForm1.Show
    'Coverts JSON into string
    Dim jsonPath As String
    jsonPath = ThisWorkbook.Path & "\phasing_config.json"
    Dim jsonString As String
    jsonString = "{ phasing_status: { ReadOrder:Pattern,P0,P1,P2,P3, Pattern:pattern, P0: phase0, P1: phase1, P2: phase2, P3: phase3 } }"
    If FileExists(jsonPath) Then jsonString = LoadFileStr(jsonPath)
    progress 10
     
    'save the active sheet as csv in temp area
    savetocsv jsonString
     progress 50
     PrepareCommand
     progress 100
     UserForm1.Hide
    Application.ScreenUpdating = True
End Sub
 Sub progress(pctCompl As Single)

UserForm1.Text.Caption = pctCompl & "% Completed"
UserForm1.Bar.Width = pctCompl * 2

DoEvents

End Sub
Sub PrepareCommand()
Dim strCmd, strCmdFmt As String
strCmdFmt = "tssupdate3 -update_phasing input_file -chip testing_prj -rev r1_0 -block ATPG_POIPU -type pattern -foundry SEC,GF -platform SLT_GGC"
'Read the Prj/rev/Block and Platform ATE as default for the command
Dim Prj As String
Dim Rev As String
'Dim block As String defind as global variable due to block needs to be placed as first argument while passing the list file.
Dim Platform As String
Dim Foundry As String

Prj = "testing_prj"
Rev = "r1_0"
block = "ATPG_POIPU"
Platform = "ATE"
Foundry = "Sec"
'Read the file name and get the project information
Dim fileName As String
fileName = ActiveWorkbook.Name
Items = Split(fileName, "_")
On Error Resume Next
Prj = Items(0)
Rev = Items(1)
block = Items(2)

On Error GoTo 0

'Using HWECGI form the tssquery command.
PrjRevBlkStatus = False
If 0 Then 'Todo enable the validation by reading back the status.
    Dim HWECGIcmd As String
    HWECGIcmd = "https://hwecgi.qualcomm.com/Resources/Dad/perl/dev/temp/cmd.cgi?cmd="
    Dim tssqueryCmd As String
    tssqueryCmd = "/prj/vlsi/pete/scripts/ptetools/tss/utils/Query/1.1/tssquery.py --project " & Prj & " -rev " & Rev & " -block " & block & " --query" & "validate_project_rev_block.query --out_path ."
    ActiveWorkbook.FollowHyperlink Address:=HWECGIcmd & tssqueryCmd, NewWindow:=True
End If

If PrjRevBlkStatus = False Then
    'Get the project revision from the user
    Dim prjrevBlkJsonFolder As String
    Dim prjrevBlkJsonName As String
    Dim prjrevBlkJson As String
    prjrevBlkJsonFolder = Environ("temp") & "\"
    prjrevBlkJsonName = ActiveWorkbook.Name & ".json"
    prjrevBlkJson = prjrevBlkJsonFolder & prjrevBlkJsonName
    If Not FileExists(prjrevBlkJson) Then
       Prj = InputBox("Enter Project name")
       Rev = InputBox("Enter Revision name ")
       block = InputBox("Enter Block name")
       Platform = InputBox("Enter Platform name")
       Foundry = InputBox("Enter Foundry name")
        'Create  json file
        Dim strJsonData As String
        strJsonData = "{ """ & "Prj""" & ":""" & Prj & """" & ", """ & "Rev""" & ":""" & Rev & """, """ & "Block""" & ":""" & block & """ , """ & "Platform""" & ":""" & Platform & """ , """ & "Foundry""" & ":""" & Foundry & """ }"
        '{ "name":"John", "age":30, "car":null }"
        'write data in to a file
        SaveAsTxtFile prjrevBlkJsonFolder, prjrevBlkJsonName, strJsonData
    Else
        'Load the json containing Prj/Rev/Block
        'parses the JSON string into object
        Dim jsonString As String
        If FileExists(prjrevBlkJson) Then jsonString = LoadFileStr(prjrevBlkJson)
        Set JsonObject = ParseJSON(jsonString)
        Prj = JsonObject("obj.Prj")
        Rev = JsonObject("obj.Rev")
        block = JsonObject("obj.Block")
        Platform = JsonObject("obj.Platform")
        Foundry = JsonObject("obj.Foundry")
    End If
End If

'Todo validate the project infotmation using RFL
strCmd = "tssupdate3 -update_phasing " & RFLUnixPath & "/" & RFLSupport & "/" & Data_file & " -chip " & Prj & " -rev " & Rev & " -block " & block & " -foundry " & Foundry & " -platform " & Platform

Dim strFile_Path As String
textfile = FreeFile
cmd_file = Environ$("username") & "_RFL_cmd_" & Format(Now, "mmddyyyy_HHmmss") & ".txt" 'Change as per your test folder path
strFile_Path = Environ("temp") & "\" & cmd_file
Open strFile_Path For Output As #textfile

Print #textfile, strCmd & vbLf & "#end_of_command"
Close #textfile
'"tssupdate3 -h #testing_prj\n#end_of_command"

sbCopyingAFile strFile_Path, RFLpath & RFLuser & "\"
        Sleep (2000)

       If FileExists(RFLpath & RFLuser & "\" & cmd_file) Then MsgBox "Looks like RFL is not running,please start running RFL in your unix teminal and hit ok "

End Sub
'''''*cb example C:\Users\sundara\Downloads\VBA-JSON-2.3.1\VBA-JSON-2.3.1\specs
'https://github.com/VBA-tools/VBA-JSON
Sub savetocsv(jsonString As String)
  
   

            
    If Find_SAVE(jsonString) Then
    'MsgBox ("found")
        
    End If

End Sub

Function Find_SAVE(jsonString As String) As Boolean
 'parses the JSON string into object
    Set JsonObject = ParseJSON(jsonString)
    Dim Readorder As String
    On Error GoTo SetError
        jsonCount = 0
    
    While errorFlg = False
        RowsC = JsonObject("obj.phasing_status(" & jsonCount & ").ReadOrder")
        jsonCount = jsonCount + 1
        If errorFlg = True Or RowsC = "" Then
SetError: errorFlg = True
            jsonCount = jsonCount - 1

        End If
    Wend
 On Error GoTo 0

    For jsonItem = 0 To jsonCount
        Readorder = JsonObject("obj.phasing_status(" & jsonItem & ").ReadOrder")
        If Readorder = "" Then Exit For
        Dim parameter_header() As String
        parameter_header = Split(Readorder, ",")
       
       'Map the parameters using JSON
        Dim parameter_save() As String
        parameter_save = parameter_header
        For i = 0 To UBound(parameter_header)
            parameter_save(i) = JsonObject("obj.phasing_status(" & jsonItem & ")." & parameter_header(i))
        Next
        
        
        Find_SAVE = True
        Dim strNames() As String ', parameter_save() As String
        strNames = parameter_header
        Dim ArrayLength As Integer
        ArrayLength = UBound(strNames)
        Dim found_Columns() As Integer
        Dim Found_location As Integer
        
        'ReDim found_Columns(1 To ArrayLength)
        ReDim Preserve found_Columns(0)
        'declare a variant to hold the array element
        Dim Item As Variant
    
        Dim FindString As String
        'loop through the entire array
        For Each Item In strNames
            'show the element in the debug window.
             Dim Rng As Range
             
             If Trim(Item) <> "" Then
                 With ActiveSheet.UsedRange 'searches all of column A
                     Set Rng = .Find(What:=CStr(Item), _
                                     LookIn:=xlValues, _
                                     LookAt:=xlWhole, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlNext, _
                                     MatchCase:=False)
                     If Not Rng Is Nothing Then
                        ' Application.Goto Rng, True 'value found
                         found_Columns(UBound(found_Columns)) = Rng.Column
                         ReDim Preserve found_Columns(UBound(found_Columns) + 1)
                         Found_location = Rng.Row
                         'MsgBox "found"
                     Else
                         'MsgBox "Nothing found" 'value not found
                         collectMsg = collectMsg & Item & vbLf
                         Find_SAVE = False
                         
                     End If
                 End With
             End If
    
         Next
         'Save the found column to csv
        If Find_SAVE Then
            ReDim Preserve found_Columns(UBound(found_Columns) - 1)
           'Open a text file for edit in vba
            Dim strFile_Path As String
            Data_file = "phasing_data_" & Environ$("username") & "_" & Format(Now, "mmddyyyy_HHmmss") & ".txt" 'Change as per your test folder path
            strFile_Path = Environ("temp") & "\" & Data_file
            textfile = FreeFile
            Open strFile_Path For Output As #textfile
            Dim writeLine As String
            writeLine = Join(parameter_save, ",")
            Print #textfile, "#block," & writeLine
            For i = Found_location + 1 To ActiveSheet.UsedRange.Rows.Count
            writeLine = block
                'For loop found_Columns
                For Each Item In found_Columns
                    writeLine = writeLine & "," & ActiveSheet.Cells(i, Item)
                 Next
                If writeLine <> "" Then Print #textfile, writeLine
            Next
                Close #textfile
            'coppy the file to server
           sbCopyingAFile strFile_Path, RFLpath & RFLSupport & "\"
           Exit For
        End If
        collectMsg = collectMsg & "/" & vbLf
    Next
    
    collectMsg = Left(collectMsg, Len(collectMsg) - 2)
    If Not Find_SAVE Then
        MsgBox "None of the set of headers are found" & vbLf & "Recomended name are as below: " & vbLf & collectMsg, , "Please check the spelling for the headers"
        End
    End If
End Function
'In this Example I am Copying the File From "C:Temp" Folder to "D:Job" Folder
Sub sbCopyingAFile(sFile As String, sDFolder As String)
'Declare Variables
Dim FSO
'sfile is Your File Name which you want to Copy
'sDFolder destination folder path
'Create Object
Set FSO = CreateObject("Scripting.FileSystemObject")

'Checking If File Is Located in the Source Folder
If Not FSO.FileExists(sFile) Then
    MsgBox "Specified File Not Found", vbInformation, "Not Found"
    
'Copying If the Same File is Not Located in the Destination Folder
Else 'If Not FSO.FileExists(sDFolder & sFile) Then
    FSO.CopyFile (sFile), sDFolder, True
    msgStr = "Specified File " & sFile & " Copied Successfully to " & sDFolder
    Application.StatusBar = msgStr ', vbInformation, "Done!"
    Debug.Print msgStr
    
'Else
    'MsgBox "Specified File Already Exists In The Destination Folder", vbExclamation, "File Already Exists"
End If

End Sub
Sub t()
SendCToTSS
End Sub
Function IsWorkBookOpen2(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
    On Error GoTo 0
End Function
