Attribute VB_Name = "CaptureDesktop"
Public Declare Function GetDesktopWindow Lib "user32" () As Long 'declare API
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt% Lib "gdi32" (ByVal hdc%, ByVal X%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal XSrc%, ByVal YSrc%, ByVal nSrcWidth%, ByVal nSrcHeight%, ByVal dwRop&)
Public Const SRCINVERT = &H660046 'set constants
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCERASE = &H440328
Public Const SRCPAINT = &HEE0086
Public Sub DumpToWindow(TargetBox As Control, change, fliph As Boolean, flipv As Boolean)
If fliph = True And flipv = True Then MsgBox "Cannot flip horizontal and vertical at the same time!": Exit Sub 'just in case you copy and paste and try to flip both ways at the same time, this doesn't really matter but I just put it for no preticular reason, so take this out if you want
Dim Desktop As Long 'this will be set the hDc of the desktop
Desktop = GetDC(GetDesktopWindow) 'get the hDc of the desktop and put it in the variable 'desktop'
ww = Screen.Width / Screen.TwipsPerPixelX 'get screen size in pixels
hh = Screen.Height / Screen.TwipsPerPixelY
BitBlt TargetBox.hdc, 0, 0, ww, hh, Desktop, 0, 0, change 'copy to form2.picture1
If fliph = True Then Call FlipPictureHorizontal(Form2.Picture1, Form2.Picture1) 'if requested, flip
If flipv = True Then Call FlipPictureVertical(Form2.Picture1, Form2.Picture1) 'if requested, flip
End Sub
Sub FlipPictureHorizontal(pic1 As PictureBox, pic2 As PictureBox)
pic1.ScaleMode = 3 'set scale modes
pic2.ScaleMode = 3
Dim px%
Dim py%
Dim retval%
px% = pic1.ScaleWidth
py% = pic1.ScaleHeight
retval% = StretchBlt(pic2.hdc, px%, 0, -px%, py%, pic1.hdc, 0, 0, px%, py%, SRCCOPY)
End Sub
Sub FlipPictureVertical(pic1 As PictureBox, pic2 As PictureBox)
pic1.ScaleMode = 3 'set scale modes
pic2.ScaleMode = 3
Dim px%
Dim py%
Dim retval%
px% = pic1.ScaleWidth
py% = pic1.ScaleHeight
retval% = StretchBlt(pic2.hdc, 0, py%, px%, -py%, pic1.hdc, 0, 0, px%, py%, SRCCOPY)
End Sub
