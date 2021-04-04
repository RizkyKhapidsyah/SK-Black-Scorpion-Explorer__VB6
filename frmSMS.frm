VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmSMS 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send SMS Message via ICQ...."
   ClientHeight    =   4800
   ClientLeft      =   3000
   ClientTop       =   2490
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "frmSMS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5070
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   1800
      Width           =   735
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4320
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox user 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1005
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "135133274"
      Top             =   2880
      Width           =   1890
   End
   Begin VB.TextBox pass 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1005
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3330
      Width           =   1920
   End
   Begin VB.TextBox prefix 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   990
      TabIndex        =   5
      Top             =   3735
      Width           =   540
   End
   Begin VB.TextBox number 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   2550
      TabIndex        =   4
      Top             =   3735
      Width           =   2325
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   345
      Left            =   3480
      TabIndex        =   1
      Top             =   840
      Width           =   1260
   End
   Begin VB.TextBox msg 
      BackColor       =   &H00E0E0E0&
      Height          =   1860
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   3195
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label copyrigth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copyrigth 2001 - Iceberg Tip"
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   2310
   End
   Begin VB.Label email 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dllegs2k@hotmail.com"
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   4335
      Width           =   1740
   End
   Begin VB.Label lblIcq 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icq # :"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   435
      TabIndex        =   11
      Top             =   2910
      Width           =   465
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3390
      Width           =   780
   End
   Begin VB.Label lblPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prefix :"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   420
      TabIndex        =   9
      Top             =   3795
      Width           =   480
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number :"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1785
      TabIndex        =   8
      Top             =   3795
      Width           =   645
   End
   Begin VB.Label words 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   3
      Top             =   2295
      Width           =   690
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You have :"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3360
      TabIndex        =   2
      Top             =   2340
      Width           =   780
   End
End
Attribute VB_Name = "frmSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdEnd_Click()

MsgBox "Stop the SMS program?", vbInformation, "SMS Disconnected.."

Unload Me

End Sub

Private Sub cmdSend_Click()

Status.Caption = "opening registry page and say you are online"
'opens the registry page and say your are online
Inet1.OpenURL "http://web.icq.com/karma/dologin/1,,,00.html?uService=1&uLogin=" + user.Text + "&uPassword=" + pass.Text
Status.Caption = "Sending the message to the phone number"
'send the message to the phone number you want
Inet1.OpenURL "http://web.icq.com/sms/send_history/1,,,00.html?target=msghistory&prefix=+" + prefix.Text + "&carrier=aaa&tophone=" + Number.Text + "&msg=" + msg.Text
MsgBox "Message sent with success !", vbInformation, "SMS text confirmed.."
Status.Caption = ""

End Sub

Private Sub email_Click()

'isTemp = "mailto:dllegs2k@hotmail.com"
'lRet = ShellExecute(hwnd, "open", isTemp, vbNull, vbNull, 1)

End Sub

Private Sub email_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

email.ForeColor = QBColor(9)
email.FontUnderline = True

End Sub

Private Sub Form_Load()

Me.Height = 5175
    Me.Width = 5160

words.Caption = 150

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

email.ForeColor = QBColor(0)
email.FontUnderline = False

End Sub

Private Sub msg_Change()

words.Caption = 150 - Len(msg)

End Sub
