VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Splitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Type GifSize
 X As Long
 Y As Long
End Type
Private TPic() As New Pic
Private FramesCount As Long
Public Function LoadGif(filename As String) As Boolean
Dim sImgHeader    As String
Dim filenameHeader   As String
Dim sBuff         As String
Dim sPicsBuff     As String
Dim ImageCNT     As Long
Dim i             As Long
Dim j             As Long
Dim xOff          As Long
Dim yOff          As Long
Dim TimeWait      As Long
Dim PointGIF     As String
If Dir$(filename) = "" Or filename = "" Then
LoadGif = False
Exit Function
End If
PointGIF = Chr$(0) & Chr$(33) & Chr$(249)
Open filename For Binary Access Read As #1
sBuff = String(LOF(1), Chr(0))
Get #1, , sBuff
Close #1
i = 1
FramesCount = 0
j = InStr(1, sBuff, PointGIF) + 1
filenameHeader = Left(sBuff, j)
If Left$(filenameHeader, 3) <> "GIF" Then
LoadGif = False
Exit Function
End If
i = j + 2
Open App.Path & "\temp.gif" For Binary As #1
ReDim Preserve TPic(0)
Do
j = InStr(i, sBuff, PointGIF) + 3
If j > Len(PointGIF) Then
sPicsBuff = String(Len(filenameHeader) + j - i, Chr$(0))
sPicsBuff = filenameHeader & Mid(sBuff, i - 1, j - i)
Put #1, 1, sPicsBuff
TPic(UBound(TPic)).SetFilePic = sPicsBuff
sImgHeader = Left(Mid(sBuff, i - 1, j - i), 16)
TPic(UBound(TPic)).SetFrameWait = ((Asc(Mid(sImgHeader, 4, 1))) + (Asc(Mid(sImgHeader, 5, 1)) * 256&)) * 10&
Set TPic(UBound(TPic)).SetPicture = LoadPicture(App.Path + "\temp.gif")
i = j
End If
If j = 3 Then Exit Do
FramesCount = FramesCount + 1
ReDim Preserve TPic(UBound(TPic) + 1)
Loop
On Error Resume Next
ReDim Preserve TPic(UBound(TPic) - 1)
On Error GoTo 0
Close #1
Kill App.Path + "\temp.gif"
LoadGif = True
End Function
Public Property Get GetFrame(ByVal index As Long) As StdPicture
If index < 0 Or index > FramesCount - 1 Then Exit Property
Set GetFrame = TPic(index).GetPicture
End Property
Public Property Get GetFramesCount() As Long
GetFramesCount = FramesCount
End Property
Public Property Get GetFrameDimenzion(ByVal index As Long) As GifSize
If index < 0 Or index > FramesCount - 1 Then Exit Property
GetFrameDimenzion = TPic(index).GetDimenzion
End Property
Public Property Get FrameWait(ByVal index As Long) As Long
If index < 0 Or index > FramesCount - 1 Then Exit Property
FrameWait = TPic(index).GetFrameWait
End Property
Public Function SaveFrame(ByVal filename As String, ByVal index As Long) As Boolean
If filename = "" Then SaveFrame = False: Exit Function
If Dir(filename) <> "" Then Kill filename
If index < 0 Or index > FramesCount - 1 Then SaveFrame = False: Exit Function
Open filename For Binary As #1
Put #1, , TPic(index).GetFilePic
Close #1
SaveFrame = True
End Function
