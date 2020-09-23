VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Gdi Clock"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   120
      Pattern         =   "*.png"
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const bytMaxSize As Byte = 26
Private Const bytMinSize As Byte = 26
Private Const gdiBicubic = 1
Dim tm(0 To 7) As String
Dim funcBlend32bpp As BLENDFUNCTION
Dim bmpInfo As BITMAPINFO
Dim dcMemory As Long, bmpMemory, Counter As Long
Dim lngHeight As Long, lngWidth As Long, lngBitmap As Long, lngImage As Long, lngGDI As Long, lngReturn As Long, lngCursor As Long
Dim sngIndex As Single, sngUBound As Single, sngStep As Single, sngStartTop As Single, sngStartLeft As Single
Dim sngHeight As Single, sngWidth As Single, sngLeft As Single, sngTop As Single, sngFrom() As Single
Dim gdipInit As GDIPLUS_STARTINPUT
Dim apiWindow As POINTAPI, apiPoint As POINTAPI, apiMouse As POINTAPI, apiMouseTmr As POINTAPI
Dim strOldTime As String

Private Sub Form_Click()
Unload Me: End
End Sub

Private Sub Form_Initialize()
strOldTime = Time
gdipInit.GDIPlusVersion = 1
If GdiplusStartup(lngGDI, gdipInit, ByVal 0&) <> 0 Then
    MsgBox "Error loading GDI+!", vbCritical
    Unload Me
End If
Me.Height = Screen.Height
Me.Width = Screen.Width
File1.Path = App.Path & "\Clock_Images\"
ReDim Preserve sngFrom(File1.ListCount)
sngStartTop = (Me.Height / Screen.TwipsPerPixelX) - bytMaxSize - 30
sngStartLeft = ((Me.Width / Screen.TwipsPerPixelX) - (8 * bytMinSize)) / 2

bmpInfo.bmpHeader.Size = Len(bmpInfo.bmpHeader)
bmpInfo.bmpHeader.BitCount = 32
bmpInfo.bmpHeader.Height = Me.ScaleHeight
bmpInfo.bmpHeader.Width = Me.ScaleWidth
bmpInfo.bmpHeader.Planes = 1
bmpInfo.bmpHeader.SizeImage = bmpInfo.bmpHeader.Width * bmpInfo.bmpHeader.Height * (bmpInfo.bmpHeader.BitCount / 8)

dcMemory = CreateCompatibleDC(Me.hdc)
bmpMemory = CreateDIBSection(dcMemory, bmpInfo, DIB_RGB_COLORS, ByVal 0, 0, 0)
SelectObject dcMemory, bmpMemory
GdipCreateFromHDC dcMemory, lngImage
draw_time
End Sub

Sub draw_time()
RestoreGDIPlus

  tm(0) = Mid$(Time$, 1, 1)
  tm(1) = Mid$(Time$, 2, 1)
  tm(2) = Mid$(Time$, 3, 1)
  tm(3) = Mid$(Time$, 4, 1)
  tm(4) = Mid$(Time$, 5, 1)
  tm(5) = Mid$(Time$, 6, 1)
  tm(6) = Mid$(Time$, 7, 1)
  tm(7) = Mid$(Time$, 8, 1)
sngLeft = sngStartLeft
For a = 0 To 7
    sngHeight = bytMinSize
    sngWidth = bytMinSize
    sngTop = 0 'sngStartTop + bytMaxSize - bytMinSize
If tm(a) = ":" Then
    LoadPictureGDIPlus App.Path & "\Clock_Images\" & "space.png", CLng(sngLeft), CLng(sngTop), CLng(sngWidth), CLng(sngHeight)
Else
    LoadPictureGDIPlus App.Path & "\Clock_Images\" & File1.List(tm(a)), CLng(sngLeft), CLng(sngTop), CLng(sngWidth), CLng(sngHeight)
End If
    sngFrom(a) = sngLeft * Screen.TwipsPerPixelX
    sngLeft = sngLeft + sngWidth
Next a
sngFrom(UBound(sngFrom)) = sngLeft * Screen.TwipsPerPixelX
UpdateGDIPlus

End Sub

Private Function RestoreGDIPlus()
DeleteObject bmpMemory
bmpMemory = CreateDIBSection(dcMemory, bmpInfo, DIB_RGB_COLORS, ByVal 0, 0, 0)
SelectObject dcMemory, bmpMemory
GdipCreateFromHDC dcMemory, lngImage
End Function

Private Function UpdateGDIPlus()
lngReturn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
SetWindowLong Me.hwnd, GWL_EXSTYLE, lngReturn Or WS_EX_LAYERED
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE

apiPoint.X = 0
apiPoint.Y = 0
apiWindow.X = Screen.Width / Screen.TwipsPerPixelX
apiWindow.Y = Screen.Height / Screen.TwipsPerPixelY

funcBlend32bpp.AlphaFormat = AC_SRC_ALPHA
funcBlend32bpp.BlendFlags = 0
funcBlend32bpp.BlendOp = AC_SRC_OVER
funcBlend32bpp.SourceConstantAlpha = 255

GdipDisposeImage lngBitmap

UpdateLayeredWindow Me.hwnd, Me.hdc, ByVal 0&, apiWindow, dcMemory, apiPoint, 0, funcBlend32bpp, ULW_ALPHA

GdipReleaseDC lngImage, dcMemory
GdipDeleteGraphics lngImage
GdipDisposeImage lngBitmap
GdiplusShutdown lngGDI

gdipInit.GDIPlusVersion = 1
If GdiplusStartup(lngGDI, gdipInit, ByVal 0&) <> 0 Then
    MsgBox "Error loading GDI+!", vbCritical
    Unload Me
End If

SelectObject dcMemory, bmpMemory
GdipCreateFromHDC dcMemory, lngImage
End Function

Private Function LoadPictureGDIPlus(strFilename As String, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
GdipLoadImageFromFile StrPtr(strFilename), lngBitmap
If Width = -1 Or Height = -1 Then
    GdipGetImageHeight lngBitmap, Height
    GdipGetImageWidth lngBitmap, Width
End If
' For better quality use:
' GdipSetInterpolationMode lngImage, gdiBicubic
GdipDrawImageRectI lngImage, lngBitmap, Left, Top, Width, Height
GdipDisposeImage lngBitmap
End Function

Private Sub Form_Unload(Cancel As Integer)
If lngImage Then
    GdipReleaseDC lngImage, dcMemory
    GdipDeleteGraphics lngImage
End If
If lngBitmap Then GdipDisposeImage lngBitmap
If lngGDI Then GdiplusShutdown lngGDI
End Sub

Private Sub Timer1_Timer()
    If Not Time = strOldTime Then
        draw_time
        strOldTime = Time
    End If
End Sub
