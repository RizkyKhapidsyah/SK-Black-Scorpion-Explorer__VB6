VERSION 5.00
Begin VB.Form frmIEport 
   BackColor       =   &H8000000E&
   Caption         =   "Import /Export Wizard..."
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   FillColor       =   &H00404040&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "frmIEport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame fraCommand 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Click Next to continue or Cancel to exit this wizard."
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   $"frmIEport.frx":058A
      Height          =   975
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label lblHeading 
      BackColor       =   &H8000000E&
      Caption         =   "Welcome to the Import/Export Wizard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblTag 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmIEport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub cmdNext_Click()

frmIEport2.Show
Unload Me

End Sub


