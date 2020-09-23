Attribute VB_Name = "mDlg"
Option Explicit
'       This file was added to allow the user to select the starting directory
'       for cDlg.Thanks to Ed Wilk for pointing this out and also for the link
'       to the code used to fix the problem: http://vbnet.mvps.org/code/callback/browsecallback.htm
'
'
'The two declarations below have nothing to do with this code but kept me from loading a new module
Public CD As cDlg 'declaration for the cDlg class
Public Loading As Boolean 'used to indicate to the user control not to draw images until all properties are updated
Public Const BIF_STATUSTEXT = 4
Private Const WM_USER = &H400
Private Const BFFM_SELCHANGED  As Long = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_ENABLEOK = (WM_USER + 101)
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Public lstPath As String 'the last path you browse with "cDlg"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Dim sBuffer As String

'I'm working on removing the file list box the commented code bellow is thje start of that
' change. Thusfar; however i havent been able to get the code to list files with any other
'pattern but all files "*.*"
''*************************************************************************************
'Private Fold As New Collection
'Private Files As New Collection
'Private Declare Function SendMessage2 Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function FindFirstFile Lib "KERNEL32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'Private Declare Function FindNextFile Lib "KERNEL32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'Private Declare Function GetFileAttributes Lib "KERNEL32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
'Private Declare Function FindClose Lib "KERNEL32" (ByVal hFindFile As Long) As Long
'Private Const MAX_PATH = 260
'Private Const MAXDWORD = &HFFFF
'Private Const INVALID_HANDLE_VALUE = -1
'Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
'Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
'Private Const FILE_ATTRIBUTE_HIDDEN = &H2
'Private Const FILE_ATTRIBUTE_NORMAL = &H80
'Private Const FILE_ATTRIBUTE_READONLY = &H1
'Private Const FILE_ATTRIBUTE_SYSTEM = &H4
'Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
'Private Const Error_NO_MORE_FILES = 0
'
'Private Type FILETIME
'    dwLowDateTime As Long
'    dwHighDateTime As Long
'End Type
'
'Private Type WIN32_FIND_DATA
'    dwFileAttributes As Long
'    ftCreationTime As FILETIME
'    ftLastAccessTime As FILETIME
'    ftLastWriteTime As FILETIME
'    nFileSizeHigh As Long
'    nFileSizeLow As Long
'    dwReserved0 As Long
'    dwReserved1 As Long
'    cFileName As String * MAX_PATH
'    cAlternate As String * 14
'End Type
'
'Private Function StripNulls(OriginalStr As String) As String
'        If (InStr(OriginalStr, Chr(0)) > 0) Then
'            OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
'        End If
'    StripNulls = OriginalStr
'End Function
'
'Private Function ListFoldersAndSubs(sStartDir As String, Optional optSubs As Boolean = False, Optional Pattern As String = "*.*") As String
'On Error Resume Next
'Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
'Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
'Dim strPath As String
'Dim eStartDir As String
'Dim lstFolders As New Collection
'Dim FilesSearched As Long
'Dim FoundInFiles As Long
'Dim NextLength As Long
'
'SearchNext:
'        If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
'    eStartDir = sStartDir
'    sStartDir = sStartDir & Pattern
'    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
'    lRet = lFileHdl
'        If lFileHdl <> -1 Then
'                Do Until lRet = Error_NO_MORE_FILES
'                    DoEvents
'                    strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
'                        If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = vbDirectory Then
'                            sTemp = StrConv(StripNulls(lpFindFileData.cFileName), vbProperCase)
'                                If sTemp <> "." And sTemp <> ".." Then
'                                    lstFolders.Add eStartDir & sTemp
'                                End If
'                        ElseIf (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
'                            sTemp = StrConv(StripNulls(lpFindFileData.cFileName), vbProperCase)
'                            Files.Add eStartDir & sTemp
'                        End If
'                    lRet = FindNextFile(lFileHdl, lpFindFileData)
'                Loop
'        End If
'    lRet = FindClose(lFileHdl)
'        If optSubs = True Then
'                If lstFolders.count > 0 Then
'                    sStartDir = lstFolders.item(1)
'                    lstFolders.Remove (1)
'                    GoTo SearchNext
'                End If
'        End If
'End Function
''***************************************************************************************

Public Function cDlg_CallBack(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
Dim ret As Long
        Select Case uMsg
            Case BFFM_INITIALIZED
                Call SendMessage(hWnd, BFFM_SETSELECTIONW, 0&, ByVal lpData)
            Case BFFM_SELCHANGED
                sBuffer = Space(255)
                ret = SHGetPathFromIDList(lParam, sBuffer)
                sBuffer = Left(sBuffer, InStr(1, sBuffer, Chr(0), vbBinaryCompare) - 1)
                    If Len(sBuffer) > 50 Then
                        sBuffer = "..." & Right(sBuffer, 50)
                    End If
                    If ret = 1 Then
                        SendMessage hWnd, BFFM_SETSTATUSTEXT, 0&, ByVal sBuffer
                    End If
        End Select
    cDlg_CallBack = 0
End Function
Public Function FarProc(pfn As Long) As Long
  FarProc = pfn
End Function

Public Function FileExists(myfilename As String, Attr As VBA.VbFileAttribute) As Boolean
    FileExists = LenB(Dir(myfilename, Attr))
End Function

