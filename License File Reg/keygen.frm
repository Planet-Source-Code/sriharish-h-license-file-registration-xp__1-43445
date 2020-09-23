VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Key Generator"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "keygen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Generate Key"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   2385
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   2
      Top             =   840
      Width           =   1725
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text4 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "License File Maker if provided in Main application."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "VOTE FOR ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "REG ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Code1 As Single

If Len(Text1.Text) < 4 Then
    MsgBox "The Name must be more than 4 characters.", vbInformation + vbOKOnly, "Ooops"
    Exit Sub
End If

For i = 1 To Len(Text1.Text) - 1
    Code1 = Format(Asc(Right(Text1.Text, Len(Text1.Text) - i)) * 2 + (39 / i) + (i + 3 / 7), "#.#")
    zip = zip & Code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    Code1 = Format(Asc(Right(zip, Len(zip) - i)) * 2 + (1 / i) + (i + 1 / 7), "#00")
    final = final & Code1
Next i
Text3.Text = "8546854"
Text4.Text = "64381"
final = Right(final, Len(final) - 4)
final = final & Asc(Text1)
Text2 = final



End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Please Vote for Me.", vbOKCancel, ":( VOTE FOR ME"
End Sub
