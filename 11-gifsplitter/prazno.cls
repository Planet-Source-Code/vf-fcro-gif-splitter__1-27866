VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Picture1 As StdPicture
Private Xlen As Long
Private Ylen As Long
Private FilePic As String
Private FrameWait As Long
Public Property Get GetDimenzion() As GifSize
GetDimenzion.X = CLng(Picture1.Width * (567 / 1000) / 15)
GetDimenzion.Y = CLng(Picture1.Height * (567 / 1000) / 15)
End Property
Public Property Set SetPicture(newpic As StdPicture)
Set Picture1 = newpic
End Property
Public Property Get GetFilePic() As String
GetFilePic = FilePic
End Property
Public Property Let SetFilePic(newstr As String)
FilePic = newstr
End Property
Public Property Get GetPicture() As StdPicture
Set GetPicture = Picture1
End Property
Public Property Get GetFrameWait() As Long
GetFrameWait = FrameWait
End Property
Public Property Let SetFrameWait(newwait As Long)
FrameWait = newwait
End Property
