Attribute VB_Name = "PhasingUpdate"
'Uses this sheet to upload data to TSS
'https://qualcomm.sharepoint.com/:x:/r/teams/ptesimteam/eRoomDocs/Tools%20_%20Methodologies/TSS/PhasingAutomation/Waipio_V1_ATPG_Tracking_Sheet.xlsx?d=w3f0f6613a9fe433788689c5ef7af6a1a&csf=1&web=1&e=X9S6X0

'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const BANG_Server As String = "\\qctdfsrt\blr\prj\vlsi\pete\scripts\ptetools\tss_data\TSS_EXCEL\"
Public Const SINGAPORE_Server As String = "\\sing\pete_scripts\ptetools\tss_data\TSS_EXCEL\"
Public Const SD_Server As String = "\\qctdfsrt\prj\vlsi\pete\tss\TSS_EXCEL\"

Sub GetServerPath()
    LocalOffsetFromGMTVar = LocalOffsetFromGMT
    If LocalOffsetFromGMTVar = BDC Then
        download_server = BANG_Server
    ElseIf LocalOffsetFromGMTVar = SINGAPORE Then
        download_server = SINGAPORE_Server
    Else
        download_server = SD_Server
    End If
End Sub

Sub phasing_update()
'https://hwecgi.qualcomm.com/Resources/Dad/perl/dev/temp/cmd.cgi?cmd=/prj/vlsi/pete/scripts/ptetools/bin/tssupdate3%20-h
'navigate to above url
cmdStr = "https://hwecgi.qualcomm.com/Resources/Dad/perl/dev/temp/cmd.cgi?cmd=/prj/vlsi/pete/scripts/ptetools/bin/tssupdate3%20-h"
ActiveWorkbook.FollowHyperlink Address:=cmdStr, NewWindow:=True


    MsgBox ("Sending Current sheet to TSS will be available soon!")
End Sub

Sub SendAToTSS()
    MsgBox ("Sending All sheet to TSS will be available soon!")
End Sub


