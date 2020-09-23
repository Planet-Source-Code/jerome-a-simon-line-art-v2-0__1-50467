VERSION 5.00
Begin VB.Form frmImageAttrib 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attributes"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Units"
      Height          =   735
      Left            =   60
      TabIndex        =   4
      Top             =   480
      Width           =   2775
      Begin VB.OptionButton optUnits 
         Caption         =   "Pixel"
         Height          =   350
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   800
      End
      Begin VB.OptionButton optUnits 
         Caption         =   "Twip"
         Height          =   350
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   800
      End
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   60
      Width           =   615
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Height:"
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Width:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmImageAttrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim imageHeight As Single
Dim imageWidth As Single
Dim imageUnits As Single

Private Sub cmdCancel_Click()
 ' Discard any changes
 Unload Me
 
End Sub

Private Sub cmdOK_Click()
 ' Set NEW Image Properties
 If imageUnits = 0 Or imageUnits = 1 Then
  ' Proper Value checked
  Units = imageUnits
 End If
 
 If imageWidth > 0 Then
  ' Larger than 0 Twips
  MaxX = imageWidth
 End If
 
 If imageHeight > 0 Then
  ' Larger than 0 Twips
  MaxY = imageHeight
 End If
 
 ' Remove this form
 Unload Me
 
End Sub

Private Sub Form_Load()
 ' Initalize Form Values
 imageUnits = Units
 imageWidth = MaxX
 imageHeight = MaxY
 
 ' Update Display
 optUnits(imageUnits).Value = True
 UnitConversion
 
 ' Center on Screen
 Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
  
End Sub
 
Private Sub optUnits_Click(Index As Integer)
 imageUnits = Index
 UnitConversion
 
End Sub

Private Sub UnitConversion()
 Select Case imageUnits
  Case 0:  ' Twips
   txtSize(0).Text = imageWidth
   txtSize(1).Text = imageHeight
   
  Case 1: ' Pixels
   txtSize(0).Text = imageWidth / Screen.TwipsPerPixelX
   txtSize(1).Text = imageHeight / Screen.TwipsPerPixelY
   
 End Select
 
End Sub

Private Sub txtSize_Change(Index As Integer)
 Dim theSize As Single
 
 Select Case Index
  Case 0: ' width
   theSize = Val(txtSize(Index))
   If imageUnits = 1 Then ' Pixels
    ' Convert
    theSize = theSize * Screen.TwipsPerPixelX
   End If
   imageWidth = theSize
   
  Case 1: ' height
   theSize = Val(txtSize(Index))
   If imageUnits = 1 Then ' Pixels
    ' Convert
    theSize = theSize * Screen.TwipsPerPixelY
   End If
   imageHeight = theSize
   
 End Select
 
End Sub

Private Sub txtSize_GotFocus(Index As Integer)
 txtSize(Index).SelStart = 0
 txtSize(Index).SelLength = Len(txtSize(Index).Text)
 
End Sub
