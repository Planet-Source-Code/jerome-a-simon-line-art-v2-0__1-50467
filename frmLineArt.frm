VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLineArt 
   Caption         =   "Line Art"
   ClientHeight    =   4950
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   2220
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox pctScrollDisplay 
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   240
      Width           =   1755
      Begin VB.PictureBox pctLineArt 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1155
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   1155
         ScaleWidth      =   1035
         TabIndex        =   3
         Top             =   0
         Width           =   1035
      End
      Begin VB.PictureBox pctBackBuffer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1155
         Left            =   180
         ScaleHeight     =   1155
         ScaleWidth      =   1035
         TabIndex        =   4
         Top             =   180
         Width           =   1035
      End
      Begin VB.PictureBox pctUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1155
         Left            =   360
         ScaleHeight     =   1155
         ScaleWidth      =   1035
         TabIndex        =   5
         Top             =   360
         Width           =   1035
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Line Art"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuDivide 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuImageClear 
         Caption         =   "&Clear"
      End
   End
   Begin VB.Menu mnuLine 
      Caption         =   "&Option"
      Begin VB.Menu mnuLineOption 
         Caption         =   "Point A - Line &BC"
         Index           =   0
      End
      Begin VB.Menu mnuLineOption 
         Caption         =   "Line AB - Line B&C"
         Index           =   1
      End
      Begin VB.Menu mnuLineOption 
         Caption         =   "Line AB - Line C&D"
         Index           =   2
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "&Image"
      Begin VB.Menu mnuSize 
         Caption         =   "&Size"
      End
      Begin VB.Menu mnuLines 
         Caption         =   "&Lines"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmLineArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StartX As Single
Dim StartY As Single

Private Sub Form_Load()
 ' Initialize Point Counter
 cPoint = 0
 
 ' Initialize Max Line Count
 MaxLines = defLineCount
 
 ' Initial Canvas Position
 Units = 0       ' default to TWIPS
 StartX = 0
 StartY = 0
 MaxX = Screen.TwipsPerPixelX * MaxPixelX
 MaxY = Screen.TwipsPerPixelY * MaxPixelY
 
 ' Initial Line Type
 LineType PointA_LineBC
 
 ' Initial Scroll Min Values
 VScroll1.SmallChange = SnapUnit
 HScroll1.SmallChange = SnapUnit
 
End Sub

Private Sub Form_Resize()
 Dim mx As Single
 Dim my As Single
 
 ' Determine display size of PictureBox (minus scroll bars)
 mx = Me.ScaleWidth
 If mx > VScroll1.Width Then mx = mx - VScroll1.Width
 my = Me.ScaleHeight
 If my > HScroll1.Height Then my = my - HScroll1.Height
 
 ' resize pctScrollDisplay
 pctScrollDisplay.Move 0, 0, mx, my
 
 ' resize (and position) scroll bars
 VScroll1.Move mx, 0, VScroll1.Width, my
 HScroll1.Move 0, my, mx, HScroll1.Height
 
 If MaxX > mx Then
  ' Canvas Larger than Scroll Area
  HScroll1.Max = MaxX - mx
  HScroll1.LargeChange = mx
  HScroll1.Enabled = True
 Else
  ' Canvas Smaller than Scroll Area
  HScroll1.Max = 0
  HScroll1.Enabled = False
 End If
 
 If MaxY > my Then
  ' Canvas Larger than Scroll Area
  VScroll1.Max = MaxY - my
  VScroll1.LargeChange = my
  VScroll1.Enabled = True
 Else
  ' Canvas Smaller than Scroll Area
  VScroll1.Max = 0
  VScroll1.Enabled = False
 End If
 
 ' position
 PositionCanvas StartX, StartY
 
End Sub

Private Sub HScroll1_Change()
 StartX = -HScroll1.Value
 PositionCanvas StartX, StartY

End Sub

Private Sub mnuAbout_Click()
 frmAbout.Show vbModal
 
End Sub

Private Sub mnuCopy_Click()
 Clipboard.SetData pctBackBuffer.Image, vbCFBitmap

End Sub

Private Sub mnuImageClear_Click()
 pctBackBuffer.Cls

End Sub

Private Sub mnuLineOption_Click(Index As Integer)
 LineType Index
 cPoint = 0

End Sub

Private Sub mnuLines_Click()
 frmImageLines.Show vbModal

End Sub

Private Sub mnuLoad_Click()
  Dim msg As String
  
  ' Set CancelError is True
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  ' Set flags
  CommonDialog1.Flags = cdlOFNHideReadOnly
  ' Set filters
  CommonDialog1.Filter = "Bitmap (*.bmp)|*.bmp"
  ' Specify default filter
  CommonDialog1.FilterIndex = 1
  ' Display the Open dialog box
  CommonDialog1.ShowOpen
  ' Display name of selected file
  With pctBackBuffer
   .AutoSize = True    ' allow loading image to st picture size
   .Picture = LoadPicture(CommonDialog1.FileName)
   .AutoSize = True    ' turn it off now... picture box resized
   MaxX = .Width       ' record values
   MaxY = .Height
  End With
  Form_Resize          ' resize display
  
  Exit Sub
  
ErrHandler:
   If Err.Number = 32755 Then Exit Sub   ' Cancel Pressed (Not Error)
   
   msg = "Error # " & Str(Err.Number) & " was generated by " _
         & Err.Source & Chr(13) & Err.Description
   MsgBox msg, , "Error", Err.HelpFile, Err.HelpContext
  'User pressed the Cancel button
  Exit Sub

End Sub

Private Sub mnuSave_Click()
  Dim msg As String
  
  ' Set CancelError is True
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  ' Set flags
  CommonDialog1.Flags = cdlOFNHideReadOnly
  ' Set filters
  CommonDialog1.Filter = "Bitmap (*.bmp)|*.bmp"
  ' Specify default filter
  CommonDialog1.FilterIndex = 1
  ' Display the Open dialog box
  CommonDialog1.ShowSave
  ' Display name of selected file
  SavePicture pctBackBuffer.Image, CommonDialog1.FileName
  Exit Sub
  
ErrHandler:
   If Err.Number = 32755 Then Exit Sub   ' Cancel Pressed (Not Error)
   
   msg = "Error # " & Str(Err.Number) & " was generated by " _
         & Err.Source & Chr(13) & Err.Description
   MsgBox msg, , "Error", Err.HelpFile, Err.HelpContext
  'User pressed the Cancel button
  Exit Sub

End Sub

Private Sub mnuSize_Click()
 frmImageAttrib.Show vbModal   ' show Image Attribute Window
 
 ' Wait for return
 StartX = 0
 StartY = 0
 Form_Resize
 
End Sub

Private Sub mnuUndo_Click()
 ' pctBackBuffer.Picture = pctUndo.Image
 BlitPic pctBackBuffer, pctUndo
 
End Sub

Private Sub pctLineArt_KeyPress(KeyAscii As Integer)
 If KeyAscii = VK_ESCAPE Then
  If cPoint > 0 Then
   cPoint = cPoint - 1
  End If
 End If

End Sub

Private Sub pctLineArt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim t As Integer
 
 x = Snap(SnapUnit, x)
 y = Snap(SnapUnit, y)
 
 ' Record Mouse Coords
 cPoint = cPoint + 1
 xPoint(cPoint) = x
 yPoint(cPoint) = y
 
 If mnuLineOption(LineAB_LineBC).Checked And cPoint = 2 Then
  cPoint = 3
  xPoint(cPoint) = xPoint(2)
  yPoint(cPoint) = yPoint(2)
 End If
 If mnuLineOption(PointA_LineBC).Checked And cPoint = 1 Then
  cPoint = 2
  xPoint(cPoint) = xPoint(1)
  yPoint(cPoint) = yPoint(1)
 End If
 
 If Not cPoint < MaxPoints Then
  ' Clear pctLineArt
  'pctUndo.Picture = pctBackBuffer.Image
  BlitPic pctUndo, pctBackBuffer
  DrawLineArt pctBackBuffer, cPoint - 3, cPoint - 2, cPoint - 1, cPoint
  cPoint = 0
 End If

End Sub

Private Sub pctLineArt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim tBar As String
 Dim t As Integer
 
 ' Clear pctLineArt
 'pctLineArt.Picture = pctBackBuffer.Image
 BlitPic pctLineArt, pctBackBuffer
 
 ' Snap Cursor to Points
 x = Snap(SnapUnit, x)
 y = Snap(SnapUnit, y)
 pctLineArt.Line (x, 0)-(x, pctLineArt.ScaleHeight), RGB(0, 150, 0)
 pctLineArt.Line (0, y)-(pctLineArt.ScaleWidth, y), RGB(0, 150, 0)
 pctLineArt.Circle (x, y), SnapUnit, RGB(150, 0, 0)
 
 If cPoint = 0 Then
  ' No Points Defined - Nothing to do!
  Exit Sub
 End If
 
 ' Atleast One Point Defined Set first "Line" point
 t = 1
 pctLineArt.Line (xPoint(t), yPoint(t))-(xPoint(t), yPoint(t))
    
 ' Draw "Other" lines (if any)
 Do While t < cPoint
  t = t + 1
  pctLineArt.Line (xPoint(t - 1), yPoint(t - 1))-(xPoint(t), yPoint(t)), LineColor(t)
 Loop
 
 ' Fallow Mouse Pointer - from last point
 pctLineArt.Line -(x, y), LineColor(t + 1)
 If cPoint < MaxPoints Then
  If cPoint = 3 Then
   xPoint(4) = x
   yPoint(4) = y
   DrawLineArt pctLineArt, cPoint - 2, cPoint - 1, cPoint, cPoint + 1
  End If
 Else
  DrawLineArt pctBackBuffer, cPoint - 3, cPoint - 2, cPoint - 1, cPoint
 End If
 
End Sub

Private Sub LineType(lType As Integer)
 Dim t As Integer
 
 ' Reset Point Values (to be new Line formation)
 cPoint = 0
 For t = 0 To LineAB_LineCD
  mnuLineOption(t).Checked = False
 Next t
 mnuLineOption(lType).Checked = True
 
End Sub

Private Sub PositionCanvas(x As Single, y As Single)
 pctLineArt.Move x, y, MaxX, MaxY
 pctBackBuffer.Move x, y, MaxX, MaxY
 pctUndo.Move x, y, MaxX, MaxY

End Sub

Private Sub VScroll1_Change()
 StartY = -VScroll1.Value
 PositionCanvas StartX, StartY
 
End Sub
