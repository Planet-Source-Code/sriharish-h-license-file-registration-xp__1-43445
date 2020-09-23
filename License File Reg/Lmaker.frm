VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Lmaker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "License File Maker"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   Icon            =   "Lmaker.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin LicenseFileRegistration.XpBs XpBs1 
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Create LIC File"
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
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   3855
      Begin VB.Label Label5 
         Caption         =   "EXAMPLE- 8546854- MAJOR KEY- 34681"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      MaxLength       =   17
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Use KEY GENERATOR to create correct License File"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "MAJOR KEY :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Email :"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Company :"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "REG ID :"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Lmaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub XpBs1_Click()
Dim enc1 As String
If Len(Text1.Text) > 0 And Len(Text2.Text) > 0 And Len(Text3.Text) > 0 And Len(Text4.Text) > 0 Then

CommonDialog1.Filter = "Lic File|*.lic|" 'Filter Extension
CommonDialog1.ShowSave
'================
If Len(CommonDialog1.FileTitle) > 0 Then
enc1 = EnigmaEncrypt(Text1.Text) 'registration name
Open CommonDialog1.FileName For Append As #1
Print #1, enc1 'print encrypted registration name
Close #1
'===============
enc1 = EnigmaEncrypt(Text2.Text) 'Major serial Key
Open CommonDialog1.FileName For Append As #1
Print #1, enc1 ' print encrypted serial key
Close #1
'==============
enc1 = EnigmaEncrypt(Text3.Text) 'Company name
Open CommonDialog1.FileName For Append As #1
Print #1, enc1 'print encrypted company name
Close #1
'==============
enc1 = EnigmaEncrypt(Text4.Text) 'Email Address
Open CommonDialog1.FileName For Append As #1
Print #1, enc1 'print encrypted Email ID
Close #1
MsgBox "License File Created at :" & CommonDialog1.FileName, vbInformation, "LIC File Created"

End If
Else
MsgBox "Fill all the information first", vbExclamation, "Knock-Knock"
End If
End Sub
