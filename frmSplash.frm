VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   3  'Dash-Dot
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      ScaleHeight     =   4155
      ScaleWidth      =   12795
      TabIndex        =   1
      Top             =   -840
      Width           =   12855
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "RENTAL MOBIL"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1935
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   10695
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   480
      Top             =   3720
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Complex"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   7560
      TabIndex        =   3
      Top             =   3600
      Width           =   885
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub



Private Sub Timer1_Timer()

If Label2.Visible = True Then
Label2.Visible = False
Else
Label2.Visible = True
End If

a = a + 1
Label1.Caption = CStr(a) & "% "
ProgressBar1.Value = a
If a = 100 Then
Unload Me
Form5_login.Show
End If

End Sub
