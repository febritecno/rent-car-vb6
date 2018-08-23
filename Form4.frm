VERSION 5.00
Begin VB.Form Form5_login 
   BorderStyle     =   0  'None
   Caption         =   "LOGIN"
   ClientHeight    =   2715
   ClientLeft      =   -60
   ClientTop       =   -75
   ClientWidth     =   3390
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   1800
      Picture         =   "Form4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   360
      Picture         =   "Form4.frx":18F02
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox tpass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fira Code Retina"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox tuser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1_login"
      BeginProperty Font 
         Name            =   "Fira Code Retina"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   480
      TabIndex        =   0
      Top             =   435
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1680
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form5_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Me.Hide
End Sub

Private Sub Command1_Click()
If tuser.Text = "" Then
    MsgBox "Username Masih Kosong!", vbCritical + vbOKOnly, "Error"
    tuser.SetFocus
ElseIf tpass.Text = "" Then
    MsgBox "Password masih Kosong", vbCritical + vbOKOnly, "Error"
    tpass.SetFocus
Else
    koneksi
    rslogin.Open " select * from login " & " where username = '" & tuser & "' " & " and pass = '" & tpass & " ' ", conn
    If rslogin.EOF Then
        MsgBox "Login Salah!", vbCritical + vbOKOnly, "Error"
        tuser.Text = ""
        tpass.Text = ""
        tuser.SetFocus
    ElseIf rslogin!UserName = "ADMIN" Then
        Form1_home.Show
        Unload Me
    Else
        With Form1_home
            .Show
            .cmd_pengaturan.Enabled = False
            Unload Me
        End With
    End If
End If
End Sub

Private Sub Command2_Click()
If MsgBox("Anda Akan Keluar Dari Program?", vbYesNo + vbInformation, "Konfirmasi") = vbYes Then End
End Sub

Private Sub Form_Load()
koneksi
End Sub

Private Sub tpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub tuser_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    tpass.SetFocus
End If
End Sub
