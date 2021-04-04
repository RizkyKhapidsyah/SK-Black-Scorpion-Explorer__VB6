VERSION 5.00
Begin VB.Form frmIEport2 
   BackColor       =   &H8000000A&
   Caption         =   "Import/Export Wizard..."
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIEport2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   3600
      Width           =   855
   End
   Begin VB.Frame fraDesc 
      Caption         =   "Description"
      Height          =   2295
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   2655
      Begin VB.Label lblDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   3360
      Width           =   5775
      Begin VB.CommandButton cmdBack 
         Caption         =   "< &Back"
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ListBox List1 
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label lblTag 
      Caption         =   "Choose an action to perform."
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000C&
      Caption         =   "      You can select what to import/export."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label lblHeading 
      BackColor       =   &H8000000C&
      Caption         =   "Import/Export Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmIEport2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim slhelp As ShellUIHelper

Private Sub cmdBack_Click()

frmIEport.Show
Unload Me

End Sub

Private Sub cmdCanc_Click()

Unload Me

End Sub

Private Sub cmdFinish_Click()

'On Error GoTo ecomm
If List1.ListIndex = 0 Then
slhelp.ImportExportFavorites True, "c:\windows\desktop\Favorites.htm"
Unload Me
End If
If List1.ListIndex = 1 Then
slhelp.ImportExportFavorites False, "c:\windows\desktop\Favorites.htm"
Unload Me
End If
'ecomm:

End Sub

Private Sub Form_Load()

Set slhelp = New ShellUIHelper

End Sub

Private Sub List1_Click()

If List1.ListIndex = 0 Then
lblDesc.Caption = "Import Favorites From File File must reside in c:\windows\desktop\Favorites.htm  Click Finish To Import it. "
End If
If List1.ListIndex = 1 Then
lblDesc.Caption = "Export Favorites to another File File will be created in c:\windows\desktop\Favorites.htm  Click Finish To Export it.  "
End If

End Sub


