Attribute VB_Name = "XTerm"
'https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a
'parser using regex
Public Const RFLpath = "\\qctdfsrt\prj\vlsi\pete\ptetools\prod\tss\runcommands\"
Public Const RFLUnixPath = "/prj/vlsi/pete/ptetools/prod/tss/runcommands/"
Public Const RFLuser = "userFiles"
Public Const RFLSupport = "supportFiles"
Public Const RFLuserOutputs = "userOutputs"
Public urlStr As String
Dim Data_file As String
Dim cmd_file As String
Dim block As String
Dim web

Public Sub XMLHTTPclient(urlStr As String)
Dim xmlhttp As New MSXML2.xmlhttp, myurl As String
myurl = "urlstr"
xmlhttp.Open "POST", myurl, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.Send "name=codingislove&email=admin@codingislove.com"
MsgBox (xmlhttp.responseText)
End Sub
Function hwecgiTransfer(Optional strCmd As String, Optional brows As String) As Boolean
    hwecgiTransfer = False
    If 1 Then 'Todo enable the validation by reading back the status.
        Dim HWECGIcmd As String
        If strCmd = "" Then strCmd = ActiveSheet.TextBox1.Text
        HWECGIcmd = "https://hwecgi.qualcomm.com/Resources/Dad/perl/dev/temp/cmd.cgi?cmd="
        urlStr = HWECGIcmd & strCmd
        'Dim tssqueryCmd As String
        'tssqueryCmd = "/prj/vlsi/pete/scripts/ptetools/tss/utils/Query/1.1/tssquery.py --project " & Prj & " -rev " & Rev & " -block " & block & " --query" & "validate_project_rev_block.query --out_path ."
        
        'If no browser is set then use user default browser is used.
         If Not IsWorkBookOpen("WebBrowser.xlam") Then
            Workbooks.Open Application.ActiveWorkbook.Path & "\WebBrowser.xlam"
            webbrowser.Show
            End If
        If test = "" Then ActiveWorkbook.FollowHyperlink Address:=urlStr, NewWindow:=True
        
       
        'If browser selected is webbrowser bilt in this project
        If browser = webbrowser Then
                     'webbrowser.Show
            
            ' General
           ' Unload webbrowser
       
    End If
        'If browser selected is internet explorer
        If browser = "IE" Then internetexplorer (urlStr)
         
        'If browser selected is native library
        If browser = "NativeXMLHTTP" Then XMLHTTPclient (urlStr)
        hwecgiTransfer = True
       
    End If
End Function
Function clearDisplay()
 With ThisWorkbook.Sheets("XtermRunner")
        .Range(.Cells(6, 1), .Cells(.UsedRange.Rows.Count, 1)).Clear
    End With
End Function
Function DisplayResult(Result As String)
    'Clear the output before printing new
    clearDisplay
    If Result = "" Then Exit Function
'Split the result in to multiple lines
'    ThisWorkbook.ActiveSheet.Cells(4, "a") = WebBrowser1.Document.body.innerHTML
    'ThisWorkbook.ActiveSheet.Cells(5, "a") = WebBrowser1.Document.body.innerText
    loadstr = Split(Result, Chr(10))
    
    'Check if Web browser needed
   
    On Error Resume Next
    If ThisWorkbook.Sheets("XtermRunner").Range("M2").Font.Bold = True And (Not UBound(loadstr) = -1) And UBound(loadstr) > 1 Then
        If (InStr(loadstr(2), "Navigation to the webpage was canceled") > 0) Then
            msgStr = "Please check your Connectivity "
            MsgBox msgStr
            Application.StatusBar = msgStr
            End
        End If
    End If
    If Not UBound(loadstr) = -1 Then
        'If InStr(loadstr(2), "Navigation to the webpage was canceled") > 0 Then
        '    MsgBox "Please check your Connectivity "
        '    Unload webbrowser
        '    End
        'End If
        For i = 0 To UBound(loadstr)
            'Check if the result is file or url and create a link to open
            ThisWorkbook.Sheets("XtermRunner").Cells(6 + i, 1).Value = loadstr(i)
        Next
    End If
    On Error GoTo 0

End Function
Public Function internetexplorer(urlStr As String) As String
    Dim ie As internetexplorer
    Dim idName As String
    Dim inputValue As String
    Set ie = New internetexplorer
    Url = urlStr
    idName = "htmlElementId"
    inputValue = "Value to be entered"
    ie.Visible = True
    ie.Navigate Url
    
    Exit Function
    While ie.Busy Or ie.ReadyState <> 4
       DoEvents
    Wend
    ie.Document.getElementById(idName).Value = inputValue
End Function
Function TransferCommond(strCmd As String) As Boolean

If strCmd = "" Then
    Application.StatusBar = "No command is entered. Please use the drop down or the below options to start."
    End
End If

Dim strFile_Path As String
textfile = FreeFile
cmd_file = Environ$("username") & "_RFL_cmd_" & Format(Now, "mmddyyyy_HHmmss") & ".txt" 'Change as per your test folder path
strFile_Path = Environ("temp") & "\" & cmd_file
Open strFile_Path For Output As #textfile

Print #textfile, strCmd & vbLf & "#end_of_command"
Close #textfile
'"tssupdate3 -h #testing_prj\n#end_of_command"
    If sbCopyingAFile(strFile_Path, RFLpath & RFLuser & "\") = False Then
        TransferCommond = False
        Exit Function
    End If
        Sleep (2000)
    If FileExists(RFLpath & RFLuser & "\" & cmd_file) Then
        MsgBox "Looks like RFL is not running,please start running RFL in your unix teminal and hit ok "
    End If
        ' Include the timer for maximum wait time
        Dim outputFile As String
        outputFile = RFLpath & RFLuserOutputs & "\" & cmd_file & ".out"
    If FileExists(outputFile) Then
         sbCopyingAFile outputFile, Environ("temp") & "\"
         loadstr = Split(LoadFileStr(Environ("temp") & "\" & cmd_file & ".out"), Chr(10))
         For i = 0 To UBound(loadstr)
            ThisWorkbook.Sheets("XtermRunner").Cells(4 + i, 1).Value = loadstr(i)
         Next
    End If
End Function
Function sbCopyingAFile(sFile As String, sDFolder As String) As Boolean
'Declare Variables
Dim FSO
'sfile is Your File Name which you want to Copy
'sDFolder destination folder path
'Create Object
Set FSO = CreateObject("Scripting.FileSystemObject")

'Checking If File Is Located in the Source Folder
If Not FSO.FileExists(sFile) Then
    MsgBox "Specified File Not Found", vbInformation, "Not Found"
    sbCopyingAFile = False
'Copying If the Same File is Not Located in the Destination Folder
Else 'If Not FSO.FileExists(sDFolder & sFile) Then
    FSO.CopyFile (sFile), sDFolder, True
    msgStr = "Specified File " & sFile & " Copied Successfully to " & sDFolder
    Application.StatusBar = msgStr ', vbInformation, "Done!"
    Debug.Print msgStr
    sbCopyingAFile = True
    
'Else
    'MsgBox "Specified File Already Exists In The Destination Folder", vbExclamation, "File Already Exists"
End If

End Function
Function Category_selected() As String
    Category_selected = ""
    If ThisWorkbook.Sheets("XtermRunner").Range("M1").Font.Bold = True Then Category_selected = "Documents"
    If ThisWorkbook.Sheets("XtermRunner").Range("M2").Font.Bold = True Then Category_selected = "Execute/Navigate"
    If ThisWorkbook.Sheets("XtermRunner").Range("M3").Font.Bold = True Then Category_selected = "Graphs"
End Function

Function IsWorkBookOpen(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
    On Error GoTo 0
End Function

Sub DisplayAddnote()
    ThisWorkbook.Sheets("XtermRunner").Range("B4").FormulaR1C1 = "Add note"
    ThisWorkbook.Sheets("XtermRunner").Range("I4").FormulaR1C1 = "Browse"
    ThisWorkbook.Sheets("XtermRunner").Range("J4").FormulaR1C1 = "Save"
    ThisWorkbook.Sheets("XtermRunner").Range("I4").Font.ThemeColor = xlThemeColorAccent1
    ThisWorkbook.Sheets("XtermRunner").Range("J4").Font.ThemeColor = xlThemeColorAccent1

    With ThisWorkbook.Sheets("XtermRunner").Range("C4:H4")
        .Merge
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
    End With
End Sub
'Here is the starting point of the function
Sub xtermmod()
    Dim strCmd As String
    Dim strResult As String
    strCmd = ActiveSheet.TextBox1.Text
    
    'Check the selected Category
    If Category_selected = "Documents" Then
        'AddSearchResult
        If Not IsWorkBookOpen("SearchAddin.xlam") Then
            Workbooks.Open Application.ActiveWorkbook.Path & "\SearchAddin.xlam"
        End If
        strResult = Application.Run("SearchAddin.xlam!Search", ActiveWorkbook.Sheets("XtermRunner").TextBox1.Text, Category_selected) 'Run "'" & WorkbookName & "!" & MacroName, argument1, argument2
        DisplayResult (strResult)
        'DisplayAddnote
    End If
    
    If Category_selected = "Graphs" Then
    End If
    
    If Category_selected = "Execute/Navigate" Then
        Application.StatusBar = "Sending command to unix to run"
        
            If hwecgiTransfer(strCmd, "webbrowser") = False Then TransferCommond (strCmd)
        Application.StatusBar = "Currently running the command" & Now
    End If
End Sub
