VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1950
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6630
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   ForeColor       =   &H00000000&
   Icon            =   "frmAccessLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808080&
      Caption         =   "show password text"
      BeginProperty Font 
         Name            =   "Neuropolitical"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "BankGothic RUSS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "ENTER THE PASSWORD"
      BeginProperty Font 
         Name            =   "BankGothic RUSS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'System Access Login v1.01
'Programmed by Bambang Abdi Setiawan
'contact@servomedia.net
'http://www.servomedia.net
'Mobile Phone +6281521566776 (SMS me please... :))

Private Sub Check1_Click()
If Check1.Value = 0 Then
Text1.PasswordChar = "*"
Check1.Caption = "show password text"
Text1.SetFocus
Else
Text1.PasswordChar = ""
Check1.Caption = "hide password text"
Text1.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0 'Pass a null char to the text box
    If Text1.Text = "hacker" Then
    Pesan = MsgBox("Access Granted!", vbOKOnly + vbInformation, "Correct Password")
    End
    Else
    Pesan = MsgBox("Access Denied!", vbOKOnly + vbCritical, "Wrong Password!")
    Text1.Text = ""
    End If
End If
End Sub
