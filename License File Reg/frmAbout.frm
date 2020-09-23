VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About License File Registration"
   ClientHeight    =   4365
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6000
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3012.801
   ScaleMode       =   0  'User
   ScaleWidth      =   5634.31
   ShowInTaskbar   =   0   'False
   Begin LicenseFileRegistration.XpBs XpBs1 
      Height          =   735
      Left            =   4320
      TabIndex        =   2
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      Caption         =   "Email"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      URL             =   "mailto:sriharish@msn.com"
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmAbout.frx":000C
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "VOTE FOR ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   33
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   2415
      Left            =   4320
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":0494
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Did you vote?", vbInformation + vbYesNo, ":( VOTE FOR ME") = vbYes Then
Unload Me
Else
MsgBox "Then Please vote", vbExclamation, ":-} VOTE FOR ME"

End If

End Sub
