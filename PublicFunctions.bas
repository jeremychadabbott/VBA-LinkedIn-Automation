Attribute VB_Name = "PublicFunctions"
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long

Private Const srccopy = &HCC0020
'Module I got from the internet for getting screen pixel color
' https://www.reddit.com/r/vba/comments/993y2o/easy_way_to_get_the_pixel_color_at_any_screen/

Public Function getScreenPixel(x As Long, y As Long) As Variant
 Dim desktopDC As LongPtr: desktopDC = GetDC(0)
 Dim memDC     As LongPtr: memDC = CreateCompatibleDC(desktopDC)
 Dim memBMP    As LongPtr: memBMP = CreateCompatibleBitmap(desktopDC, 1, 1)
 If SelectObject(memDC, memBMP) <> 0 And BitBlt(memDC, 0, 0, 1, 1, desktopDC, x, y, srccopy) <> 0 Then
  getScreenPixel = GetPixel(memDC, 0, 0)
 End If
 DeleteObject memBMP
 DeleteDC memDC
End Function

