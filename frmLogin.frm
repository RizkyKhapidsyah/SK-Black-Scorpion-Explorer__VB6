VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Black Scorpion ""Login"""
   ClientHeight    =   3075
   ClientLeft      =   2775
   ClientTop       =   3465
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3075
   ScaleWidth      =   6105
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      TabIndex        =   2
      Text            =   "Screen Name"
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogin_Click()

If txtUser.Text = ("Password") Then
frmLogin.Visible = False
MDImain.Show
Else
MsgBox "Please try again", vbCritical, "Error"
End If

End Sub

Private Sub Form_Load()

frmLogin.Caption = ("Login")

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End
    
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
cmdLogin_Click
End If

End Sub


