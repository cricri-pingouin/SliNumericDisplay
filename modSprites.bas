Attribute VB_Name = "modSprites"
Option Explicit

'Windows GDI Bitmap API
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Integer

'Windows GDI API constants and Functions for HDCs
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'Arrays holding HDCs
Dim MemHdc() As Long
Dim BitmapHdc() As Long
Dim TrashBmpHdc() As Long
Dim NumOfDcs As Byte 'Integer 'Byte is enough here

Function CreateMemHdc(ScreenHdc As Long, Width As Integer, Height As Integer) As Long
   'This function will create a temporary Hdc to blit in and out of
   'ScreenHdc = the display DC that we will be compatible
   'Width = width of needed bitmap
   'Height = height of needed bitmap
   ReDim Preserve MemHdc(NumOfDcs)
   ReDim Preserve BitmapHdc(NumOfDcs)
   ReDim Preserve TrashBmpHdc(NumOfDcs)

   MemHdc(NumOfDcs) = CreateCompatibleDC(ScreenHdc)
   If MemHdc(NumOfDcs) Then
      BitmapHdc(NumOfDcs) = CreateCompatibleBitmap(ScreenHdc, Width, Height)
      If BitmapHdc(NumOfDcs) Then
         TrashBmpHdc(NumOfDcs) = SelectObject(MemHdc(NumOfDcs), BitmapHdc(NumOfDcs))
         CreateMemHdc = MemHdc(NumOfDcs)
      End If
   End If
   NumOfDcs = NumOfDcs + 1
End Function

Sub DestroyHdcs()
   'Subroutine to free all Dc's
   Dim RetVal As Long
   Dim i As Byte 'Integer 'Byte is enough here
   
   For i = 0 To NumOfDcs - 1
      BitmapHdc(i) = SelectObject(MemHdc(i), TrashBmpHdc(i))
      RetVal = DeleteObject(BitmapHdc(i))
      RetVal = DeleteDC(MemHdc(i))
   Next i
End Sub

Sub LoadBmpToHdc(MHdc As Long, FileN As String)
   'Load a bitmap picture to hdc... I bet you would have guess :)
   Dim OrgBmp As Long

   'You can use App.Path if you wish, but it's not necessary in this example
   OrgBmp = SelectObject(MHdc, LoadPicture(FileN)) '(App.Path & "\" & FileN))
   If OrgBmp Then
      DeleteObject (OrgBmp)
   End If
End Sub
