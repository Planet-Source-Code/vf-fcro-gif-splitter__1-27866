Attribute VB_Name = "OpenSave"

' ============================================================================
' GetOpen/SaveFileName

Public Const MAX_PATH = 260

Public Type OPENFILENAME  '  ofn
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As OFN_Flags
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

' File Open/Save Dialog Flags
Public Enum OFN_Flags
  OFN_READONLY = &H1
  OFN_OVERWRITEPROMPT = &H2
  OFN_HIDEREADONLY = &H4
  OFN_NOCHANGEDIR = &H8
  OFN_SHOWHELP = &H10
  OFN_ENABLEHOOK = &H20
  OFN_ENABLETEMPLATE = &H40
  OFN_ENABLETEMPLATEHANDLE = &H80
  OFN_NOVALIDATE = &H100
  OFN_ALLOWMULTISELECT = &H200
  OFN_EXTENSIONDIFFERENT = &H400
  OFN_PATHMUSTEXIST = &H800
  OFN_FILEMUSTEXIST = &H1000
  OFN_CREATEPROMPT = &H2000
  OFN_SHAREAWARE = &H4000
  OFN_NOREADONLYRETURN = &H8000&
  OFN_NOTESTFILECREATE = &H10000
  OFN_NONETWORKBUTTON = &H20000
  OFN_NOLONGNAMES = &H40000               ' force no long names for 4.x modules
  OFN_EXPLORER = &H80000                       ' new look commdlg
  OFN_NODEREFERENCELINKS = &H100000
  OFN_LONGNAMES = &H200000                 ' force long names for 3.x modules
  ' ===============================
  ' Win98/NT5 only...
  OFN_ENABLEINCLUDENOTIFY = &H400000           ' send include message to callback
  OFN_ENABLESIZING = &H800000
  ' ===============================
End Enum

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'

Public Function GetOpenFilePath(hWnd As Long, _
                                                      sFilter As String, _
                                                      iFilter As Integer, _
                                                      sFile As String, _
                                                      sInitDir As String, _
                                                      sTitle As String, _
                                                      sRtnPath As String) As Boolean
  Dim ofn As OPENFILENAME
  
  With ofn
    .lStructSize = Len(ofn)
    .hWndOwner = hWnd
    .lpstrFilter = sFilter & vbNullChar & vbNullChar
    .nFilterIndex = iFilter
    .lpstrFile = sFile & String$(MAX_PATH - Len(sFile), 0)
    .nMaxFile = MAX_PATH
    .lpstrInitialDir = sInitDir
    .lpstrTitle = sTitle & vbNullChar
    .Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or OFN_EXPLORER
  End With
  
  If GetOpenFileName(ofn) Then
    iFilter = ofn.nFilterIndex
    sFile = Mid$(ofn.lpstrFile, ofn.nFileOffset + 1, InStr(ofn.lpstrFile, vbNullChar) - (ofn.nFileOffset + 1))
    sRtnPath = GetStrFromBufferA(ofn.lpstrFile)
    GetOpenFilePath = True
  End If

End Function

Public Function GetSaveFilePath(hWnd As Long, _
                                                      sFilter As String, _
                                                      iFilter As Integer, _
                                                      sDefExt As String, _
                                                      sFile As String, _
                                                      sInitDir As String, _
                                                      sTitle As String, _
                                                      sRtnPath As String) As Boolean
  Dim ofn As OPENFILENAME
  With ofn
    .lStructSize = Len(ofn)
    .hWndOwner = hWnd
    .lpstrFilter = sFilter & vbNullChar & vbNullChar
    .lpstrFile = sFile & String$(MAX_PATH - Len(sFile), 0)
    .lpstrDefExt = sDefExt
    .nMaxFile = MAX_PATH
    .lpstrInitialDir = sInitDir
    .lpstrTitle = sTitle & vbNullChar
    .Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
  End With
  
  If GetSaveFileName(ofn) Then
    iFilter = ofn.nFilterIndex
    sFile = Mid$(ofn.lpstrFile, ofn.nFileOffset + 1, InStr(ofn.lpstrFile, vbNullChar) - (ofn.nFileOffset + 1))
    sRtnPath = GetStrFromBufferA(ofn.lpstrFile)
    GetSaveFilePath = True
  End If

End Function

' Returns the string before first null char (if any) in an ANSII string.

Public Function GetStrFromBufferA(szA As String) As String
  If InStr(szA, vbNullChar) Then
    GetStrFromBufferA = Left$(szA, InStr(szA, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would rtn a zero length string ("").
    GetStrFromBufferA = szA
  End If
End Function
