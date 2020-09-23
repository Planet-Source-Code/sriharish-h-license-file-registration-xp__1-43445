VERSION 5.00
Begin VB.Form showsoftware 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mysoftware"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4920
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4920
   StartUpPosition =   1  'CenterOwner
   Begin LicenseFileRegistration.XpBs XpBs1 
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Caption         =   "Email Me : Sriharis@msn.com"
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
   Begin VB.Label Label2 
      Caption         =   "Dont Forget to Vote For ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "MY SOFTWARE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu file 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "showsoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Unload Form1
End Sub
