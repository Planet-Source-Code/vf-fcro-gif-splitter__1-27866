Attribute VB_Name = "Module1"
Public TotalFrames As Long
Public Function LoadGif(filename As String, aImg As Variant) As Long
    
   Dim hFile         As Long
   Dim sImgHeader    As String
   Dim filenameHeader   As String
   Dim sBuff         As String
   Dim sPicsBuff     As String
   Dim nImgCount     As Long
   Dim i             As Long
   Dim j             As Long
   Dim xOff          As Long
   Dim yOff          As Long
   Dim TimeWait      As Long
   Dim PointGIF     As String
   
If Dir$(filename) = "" Or filename = "" Then
Exit Function
End If
   
PointGIF = Chr$(0) & Chr$(33) & Chr$(249)

Open filename For Binary Access Read As #1
sBuff = String(LOF(hFile), Chr(0))
Get #1, , sBuff
Close #1
    
i = 1
nImgCount = 0
j = InStr(1, sBuff, PointGIF) + 1
filenameHeader = Left(sBuff, j)
    

If Left$(filenameHeader, 3) <> "GIF" Then
Exit Function
End If
    
LoadGif = True
    
   'set pointer ahead 2 bytes from the
   'end of the gif magic number
    i = j + 2
    
   'if the fileheader size was greater than
   '127, the info on how many individual
   'frames the gif has is located within the header.
    If Len(filenameHeader) >= 127 Then
        RepeatTimes& = Asc(Mid(filenameHeader, 126, 1)) + (Asc(Mid(filenameHeader, 127, 1)) * 256&)
    Else
        RepeatTimes = 0
    End If
    
  'create a temporary file in the current directory
'   hFile = FreeFile
   Open App.Path & "\temp.gif" For Binary As hFile
            
  'split out each frame of the gif, and
  'write each the frame to the temporary file.
  'Then load an image control for the frame,
  'and load the temp file into that control.
   Do
      
     'increment counter
      nImgCount = nImgCount + 1
      
     'locate next frame end
      j = InStr(i, sBuff, PointGIF) + 3
        
     'another check
      If j > Len(PointGIF) Then
        
        'pad an output string, fill with the
        'frame info, and write to disk. A header
        'needs to be added as well, to assure
        'LoadPicture recognizes it as a gif.
        'Since VB's LoadPicture command ignores
        'header info and loads animated gifs as
        'static, we can safely reuse the header
        'extracted above.
         sPicsBuff = String(Len(filenameHeader) + j - i, Chr$(0))
         sPicsBuff = filenameHeader & Mid(sBuff, i - 1, j - i)
         Put #hFile, 1, sPicsBuff
         
        'The first part of the
        'extracted data is frame info
         sImgHeader = Left(Mid(sBuff, i - 1, j - i), 16)
        
        'embedded in the frame info is a
        'field that represents the frame delay
         TimeWait = ((Asc(Mid(sImgHeader, 4, 1))) + (Asc(Mid(sImgHeader, 5, 1)) * 256&)) * 10&
            
        'assign the data.
         If nImgCount > 1 Then
         
           'if this is the second or later
           'frame, load an image control
           'for the frame
            Load aImg(nImgCount - 1)
            
           'the frame header also contains
           'the x and y offsets of the image
           'in relation to the first (0) image.
            xOff = Asc(Mid(sImgHeader, 9, 1)) + (Asc(Mid(sImgHeader, 10, 1)) * 256&)
            yOff = Asc(Mid(sImgHeader, 11, 1)) + (Asc(Mid(sImgHeader, 12, 1)) * 256&)
            
           'position the image controls at
           'the required position
            aImg(nImgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
            aImg(nImgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
            
         End If
         
           'use each control's .Tag property to
           'store the frame delay period, and
           'load the picture into the image control.
            aImg(nImgCount - 1).Tag = TimeWait
            aImg(nImgCount - 1).Picture = LoadPicture(App.Path + "\temp.gif")
            
           'update pointer
            i = j
        End If
    
   'when the j = Instr() command above returns 0,
   '3 is added, so if j = 3 there was no more
   'data in the header. We're done.
    Loop Until j = 3
      
  'close and nuke the temp file
   Close #hFile
   Kill App.Path + "\temp.gif"

   TotalFrames = aImg.Count - 1
   
   LoadGif = TotalFrames
   Exit Function
    
ErrHandler:

   MsgBox "Error No. " & Err.Number & " when reading file", vbCritical
   LoadGif = False
   On Error GoTo 0
    
End Function

