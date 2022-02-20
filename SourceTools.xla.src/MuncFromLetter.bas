Attribute VB_Name = "MuncFromLetter"
'WNetGetConnection: Get UNC Path for Mapped Drive
'Author:     VBnet - Randy Birch
'http://vbnet.mvps.org/index.html?code/network/uncfrommappeddrive.htm

Private Const ERROR_SUCCESS As Long = 0
Private Const MAX_PATH As Long = 260
#If VBA7 Then
Private Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Private Declare PtrSafe Function PathIsNetworkPath Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Long
Private Declare PtrSafe Function PathIsUNC Lib "shlwapi.dll" Alias "PathIsUNCA" (ByVal pszPath As String) As Long
Private Declare PtrSafe Function PathStripToRoot Lib "shlwapi.dll" Alias "PathStripToRootA" (ByVal pPath As String) As Long
Private Declare PtrSafe Function PathSkipRoot Lib "shlwapi.dll" Alias "PathSkipRootA" (ByVal pPath As String) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare PtrSafe Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
#Else
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Private Declare Function PathIsNetworkPath Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Long
Private Declare Function PathIsUNC Lib "shlwapi.dll" Alias "PathIsUNCA" (ByVal pszPath As String) As Long
Private Declare Function PathStripToRoot Lib "shlwapi.dll" Alias "PathStripToRootA" (ByVal pPath As String) As Long
Private Declare Function PathSkipRoot Lib "shlwapi.dll" Alias "PathSkipRootA" (ByVal pPath As String) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
#End If

Public Function GetUncFullPathFromMappedDrive(sLocalName As String) As String
  Dim sLocalRoot As String
  Dim sRemoteName As String
  Dim cbRemoteName As Long

  sRemoteName = Space$(MAX_PATH)
  cbRemoteName = Len(sRemoteName)
  'get the drive letter
  sRemotePath = StripRootFromPath(sLocalName)
  sLocalRoot = StripPathToRoot(sLocalName)
  'if drive letter is a network share,
  'resolve the share UNC name
  If IsPathNetPath(sLocalRoot) Then
    If WNetGetConnection(sLocalRoot, _
                         sRemoteName, _
                         cbRemoteName) = ERROR_SUCCESS Then

      'this assures the retrieved name is in
      'fact a valid UNC path.
      sRemoteName = QualifyPath(TrimNull(sRemoteName)) & sRemotePath

      If IsUNCPathValid(sRemoteName) Then
        GetUncFullPathFromMappedDrive = TrimNull(sRemoteName)
      End If
    End If
  End If
End Function

Private Function QualifyPath(sPath As String) As String
'add trailing slash if required
  If Right$(sPath, 1) <> "\" Then
    QualifyPath = sPath & "\"
  Else
    QualifyPath = sPath
  End If
End Function

Public Function IsPathNetPath(ByVal sPath As String) As Boolean
'Determines whether a path represents network resource.
  IsPathNetPath = PathIsNetworkPath(sPath) = 1
End Function

Private Function IsUNCPathValid(ByVal sPath As String) As Boolean
'Determines if string is a valid UNC
  IsUNCPathValid = PathIsUNC(sPath) = 1
End Function

Private Function StripPathToRoot(ByVal sPath As String) As String
'Removes all of the path except for
'the root information (ie drive. Also
'removes any trailing slash.
  Dim pos As Integer
  Call PathStripToRoot(sPath)
  pos = InStr(sPath, Chr$(0))
  If pos Then
    StripPathToRoot = Left$(sPath, pos - 2)
  Else
    StripPathToRoot = sPath
  End If
End Function

Private Function TrimNull(startstr As String) As String
  TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
End Function

Private Function StripRootFromPath(ByVal sPath As String) As String
'Parses a path, ignoring the drive
'letter or UNC server/share path parts
  StripRootFromPath = TrimNull(GetStrFromPtrA(PathSkipRoot(sPath)))
End Function

Private Function GetStrFromPtrA(ByVal lpszA As Long) As String
'Given a pointer to a string, return the string
  GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
  Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function


