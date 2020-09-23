VERSION 5.00
Begin VB.Form frmImageLines 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lines"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctExample 
      BackColor       =   &H00FFFFFF&
      Height          =   2235
      Left            =   1500
      ScaleHeight     =   2175
      ScaleWidth      =   2655
      TabIndex        =   4
      Top             =   60
      Width           =   2715
   End
   Begin VB.CommandButton cmdLessLines 
      Caption         =   "Less Lines"
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   1395
   End
   Begin VB.CommandButton cmdMoreLines 
      Caption         =   "More Lines"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1920
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1500
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmImageLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Spacing As Integer = 50
Dim imageLines As Single

Private Sub cmdCancel_Click()
 ' Discard any changes
 Unload Me
 
End Sub

Private Sub cmdLessLines_Click()
 If imageLines > 2 Then
  imageLines = imageLines / 2
 End If

 Form_Resize
End Sub

Private Sub cmdMoreLines_Click()
 If imageLines < MaxLineCount Then
  imageLines = imageLines * 2
 End If

 Form_Resize
End Sub

Private Sub cmdOK_Click()
 ' Set NEW Image Properties
 If imageLines > 1 And imageLines < 33 Then
  ' Proper Value checked
  MaxLines = imageLines
 End If
 
 If imageLines > 0 Then
  ' Larger than 0 Twips
  MaxLines = imageLines
 End If
 
 ' Remove this form
 Unload Me
 
End Sub

Private Sub Form_Load()
 ' Initalize Form Values
 imageLines = MaxLines
 
 cPoint = 4
 With pctExample
  xPoint(1) = Spacing
  yPoint(1) = Spacing
  
  xPoint(2) = Spacing
  yPoint(2) = .ScaleHeight - Spacing
  
  xPoint(3) = .ScaleWidth / 2
  yPoint(3) = .ScaleHeight - Spacing
  
  xPoint(4) = .ScaleWidth - Spacing
  yPoint(4) = Spacing
 End With
 
 ' Adjust Example to look like possible drawing
 If frmLineArt.mnuLineOption(LineAB_LineBC).Checked Then
  xPoint(3) = xPoint(2)
  yPoint(3) = yPoint(2)
 End If
 If frmLineArt.mnuLineOption(PointA_LineBC).Checked Then
  xPoint(2) = xPoint(1)
  yPoint(2) = yPoint(1)
 End If
 
 ' Center on Screen
 Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 
End Sub

Private Sub Form_Resize()
 pctExample.Cls
 DrawLineArt pctExample, 1, 2, 3, 4, (imageLines)
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim t As Integer
 
 For t = 0 To MaxPoints
  xPoint(t) = 0
  yPoint(t) = 0
 Next t
 
 cPoint = 0
End Sub
