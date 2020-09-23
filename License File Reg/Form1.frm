VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lisence File Registration"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Enter &Registration Information"
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin LicenseFileRegistration.XpBs XpBs4 
         Height          =   375
         Left            =   3720
         TabIndex        =   19
         Top             =   2640
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Set Trial To Zero"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         BackColor       =   255
         ForeColor       =   16777215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Locate License File"
         Enabled         =   0   'False
         Height          =   2535
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   5775
         Begin LicenseFileRegistration.XpBs XpBs3 
            Height          =   375
            Left            =   4200
            TabIndex        =   18
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "&Browse"
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
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   480
            Width           =   3975
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   5040
            Top             =   1080
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label capserial 
            Caption         =   "-"
            Height          =   255
            Left            =   1080
            TabIndex        =   27
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label capemail 
            Caption         =   "-"
            Height          =   255
            Left            =   1080
            TabIndex        =   26
            Top             =   1680
            Width           =   3255
         End
         Begin VB.Label capcompany 
            Caption         =   "-"
            Height          =   255
            Left            =   1080
            TabIndex        =   25
            Top             =   1320
            Width           =   3135
         End
         Begin VB.Label capid 
            Caption         =   "-"
            Height          =   255
            Left            =   1080
            TabIndex        =   24
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label9 
            Caption         =   "Serial :"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Email :"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Company :"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Reg ID :"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   735
         End
      End
      Begin LicenseFileRegistration.XpBs XpBs2 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&Continue"
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   49152
         ForeColor       =   16777215
      End
      Begin LicenseFileRegistration.XpBs XpBs1 
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&Validate"
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   33023
         ForeColor       =   16777215
      End
      Begin LicenseFileRegistration.Xp_ProgressBar Xps 
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4560
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         MaxLength       =   17
         TabIndex        =   1
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "License File Registration Maker By Sri Harish"
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
         Left            =   120
         TabIndex        =   29
         Top             =   5640
         Width           =   3975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fear No Copy-Write"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   5880
         Width           =   3375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "out of 10 Executions"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Registration :"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail :"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "&Company :"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "&Reg ID :"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function TrialTime(TheForm As Form, TrialOverMSG As String, TrialOverMSGTitle As String, TrialOverMSGType As String, TrialCount As Integer, Work As Boolean)

    If Not Work Then SaveSetting TheForm.Name, "Trial", "TimesOpen", ".": End
'If Work = False then reset trial to 0 if Work = True then Count up the Trial

    SaveSetting TheForm.Name, "Trial", "TimesOpen", Val(GetSetting(TheForm.Name, "Trial", "TimesOpen")) + 1
'Write + 1 to the last to the last time opened

    If GetSetting(TheForm.Name, "Trial", "TimesOpen") > TrialCount Then SaveSetting TheForm.Name, "Trial", "TimesOpen", TrialCount: MsgBox TrialOverMSG, TrialOverMSGType, TrialOverMSGTitle: End
'If the amount of times open is > then the TrialCount..
'Reset it to the number in TrialCount specified
'Display a message and terminate the program
End Function



Private Sub Form_Load()
Label1.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
'Progress bar status
If Val(Label1.Caption) = 0 Then
Xps.Value = 0 'Xps is a progressbar name
End If
If Val(Label1.Caption) = 1 Then
Xps.Value = 10
End If
If Val(Label1.Caption) = 2 Then
Xps.Value = 20
End If
If Val(Label1.Caption) = 3 Then
Xps.Value = 30
End If
If Val(Label1.Caption) = 4 Then
Xps.Value = 40
End If
If Val(Label1.Caption) = 5 Then
Xps.Value = 50
End If
If Val(Label1.Caption) = 6 Then
Xps.Value = 60
End If
If Val(Label1.Caption) = 7 Then
Xps.Value = 70
End If
If Val(Label1.Caption) = 8 Then
Xps.Value = 80
End If
If Val(Label1.Caption) = 9 Then
Xps.Value = 90
End If
If Val(Label1.Caption) = 10 Then
Xps.Value = 100
End If
End Sub

Private Sub XpBs1_Click()
'Registration  Code format
Dim i
Dim zip
Dim final
Dim code1 As Single
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
MsgBox ("Please Fill In All The Information!"), vbInformation, ("Registration")
Exit Sub
End If


If Len(Text1.Text) < 4 Then
    MsgBox "The Name must be more than 4 characters.", vbInformation + vbOKOnly, "Ooops"
    Exit Sub
End If

If Text5.Text = ("8546854") And Text6.Text = "64381" Then


Else
    MsgBox "Registration Failed. Please check your information", vbCritical, ("Registration")
Exit Sub
End If


For i = 1 To Len(Text1.Text) - 1
    code1 = Format(Asc(Right(Text1.Text, Len(Text1.Text) - i)) * 2 + (39 / i) + (i + 3 / 7), "#.#")
    zip = zip & code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    code1 = Format(Asc(Right(zip, Len(zip) - i)) * 2 + (1 / i) + (i + 1 / 7), "#00")
    final = final & code1
Next i
final = Right(final, Len(final) - 4)
final = final & Asc(Text1)
'If reg code is correct
If Text2.Text = final Then
'Enable License file Frame
Frame2.Enabled = True
    MsgBox "Registration inforamtion correct please locate your license file.", vbInformation + vbOKOnly, "Registered"
Else
    MsgBox "Registration Failed. Please check your information", vbCritical, ("Registration")
End If

End Sub

Private Sub XpBs2_Click()
TrialTime Me, "The trial of Mysoftware" & " has expired. Please register this product to get the full version.", "Trial Expired", vbCritical, 10, True
'Activates the trial counter. True to count up and False to reset the Trial count
    Label1.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
'Display times open
showsoftware.Show
Unload Me
End Sub

Private Sub XpBs3_Click()
' This section decrypts the Lic File
' and tries to match information which is
'Typed in the Validation Key section
'=====================================
'Declare file inputs
Dim regid, majorkey, companyname, emailaddress
'Declare decrypion inputs
Dim deregid, dekey, decompanyname, deemailaddress

CommonDialog1.Filter = "Lic File|*.lic| 'change filter to Lic"
CommonDialog1.ShowOpen
If Len(CommonDialog1.FileTitle) > 0 Then

Open CommonDialog1.FileName For Input As #1
'Input all the information line by line
On Error Resume Next 'if there is no 4 lines in LIC file
Line Input #1, regid
Line Input #1, majorkey
Line Input #1, companyname
Line Input #1, emailaddress
'Begin decryption process
deregid = EnigmaDecrypt(regid) 'decrypt Line One:Reg ID
dekey = EnigmaDecrypt(majorkey) 'decrypt line two:Major Key
decompanyname = EnigmaDecrypt(companyname) 'decrypt line three:Company name
deemailaddress = EnigmaDecrypt(emailaddress) 'decrypt line 4 : email ID
'==================================
'After Decryption Begin comparison
'if all the decrypted information are matched
'with the information typed in the text fields
'then trial will be unlocked
'create a check file which is verified
'every time the program starts up
Close #1
Open App.Path & "\" & "_check.ini" For Output As #1
Print #1, deregid
Print #1, dekey
Print #1, decompanyname
Print #1, deemailaddress
Close #1
'verify the check.ini file
Dim checkreg, checkserial, checkcompany, checkemail
Open App.Path & "\" & "_check.ini" For Input As #1
Line Input #1, checkreg
Line Input #1, checkserial
Line Input #1, checkcompany
Line Input #1, checkemail
'fill all the captions
capid.Caption = checkreg
capserial.Caption = checkserial
capcompany.Caption = checkcompany
capemail.Caption = checkemail
Close #1
'if information in captions match with typed text then registered
If capid.Caption = Text1.Text And capserial.Caption = Text2.Text And capcompany.Caption = Text3.Text And capemail.Caption = Text4.Text Then
MsgBox "Thank you for registering and supporting shareware.Make sure you don't lose your liscense file and Registration information", vbInformation, "Thank you-Program Registered"
Else
'Decrypted information didn't match with information in text box
MsgBox "Invalid Registration information found in License file. If you have obtained Serial Key and other information legally then please contact your customer support at http://www.mycompanysite.com", vbCritical, "Registration Failed"
Kill App.Path & "\" & "_check.ini"
Exit Sub
End If
End If
'error handler


End Sub

Private Sub XpBs4_Click()
    SaveSetting Me.Name, "Trial", "TimesOpen", 0
'Resets the trial
    Label1.Caption = 0
'Resets the Label
End Sub
