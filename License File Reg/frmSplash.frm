VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3765
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   3765
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin LicenseFileRegistration.XpBs XpBs3 
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   2760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "About Author"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      Enabled         =   0   'False
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
   End
   Begin LicenseFileRegistration.XpBs XpBs2 
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Continue"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      Enabled         =   0   'False
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
   End
   Begin LicenseFileRegistration.XpBs XpBs1 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "License File Maker"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      Enabled         =   0   'False
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
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   720
      Top             =   120
   End
   Begin LicenseFileRegistration.Xp_ProgressBar Xp_Pro 
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
   End
   Begin VB.Label chkserial 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label chkreg 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.mysite.com"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Website :"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 19XX-200X"
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MYSOFTWARE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'This will enable Progressbar
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
'Progress bar operations
Xp_Pro.Value = Xp_Pro.Value + 1
If Xp_Pro.Value > 99 Then
Timer1.Enabled = False
XpBs1.Enabled = True 'enable continue
XpBs2.Enabled = True 'enable license file maker
XpBs3.Enabled = True 'enable Author Dialog
Label6.Caption = "Loding Complete"
End If
End Sub

Private Sub XpBs1_Click()
'license file maker
Lmaker.Show
Unload Me

End Sub

Private Sub XpBs2_Click()

'if there is _check.ini file then verify it
Dim chekreg As String
Dim chekserial

Close #1
On Error GoTo errors

Open App.Path & "\" & "_check.ini" For Input As #1

Line Input #1, chekreg
Line Input #1, chekserial
chkreg.Caption = chekreg
chkserial.Caption = chekserial
Close #1
' if _check.ini present then verify
Dim i
Dim zip
Dim final
Dim code1 As Single

For i = 1 To Len(chkreg.Caption) - 1
    code1 = Format(Asc(Right(chkreg.Caption, Len(chkreg.Caption) - i)) * 2 + (39 / i) + (i + 3 / 7), "#.#")
    zip = zip & code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    code1 = Format(Asc(Right(zip, Len(zip) - i)) * 2 + (1 / i) + (i + 1 / 7), "#00")
    final = final & code1
Next i
final = Right(final, Len(final) - 4)
final = final & Asc(chkreg.Caption)
'If _check.ini  is correct the directly switch to main software
If chkserial.Caption = final Then
showsoftware.Show
Unload Me
Exit Sub
Unload Form1
Else
    'invalid _check.ini show form1
Form1.Show


Unload Me
End If
'or if _check.ini is not found then show form1
errors: Form1.Show
Unload Me
Exit Sub

End Sub

Private Sub XpBs3_Click()
frmAbout.Show
Unload Me
End Sub
