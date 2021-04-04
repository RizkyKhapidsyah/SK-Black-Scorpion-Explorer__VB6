VERSION 5.00
Begin VB.Form frmCalculator 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   2880
   ClientLeft      =   2835
   ClientTop       =   3990
   ClientWidth     =   2520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "frmCalculator.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2880
   ScaleWidth      =   2520
   Begin VB.CommandButton cmdVat 
      Caption         =   "VAT"
      Height          =   225
      Left            =   720
      TabIndex        =   24
      Top             =   600
      Width           =   525
   End
   Begin VB.CommandButton Percent 
      Caption         =   "%"
      Height          =   360
      Left            =   1080
      TabIndex        =   22
      Top             =   2400
      Width           =   360
   End
   Begin VB.CommandButton Operator 
      Caption         =   "="
      Height          =   360
      Index           =   4
      Left            =   1560
      TabIndex        =   21
      Top             =   2400
      Width           =   360
   End
   Begin VB.CommandButton Dmal 
      Caption         =   "."
      Height          =   360
      Left            =   600
      TabIndex        =   20
      Top             =   2400
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   360
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      Height          =   360
      Index           =   0
      Left            =   1560
      TabIndex        =   18
      Top             =   960
      Width           =   360
   End
   Begin VB.CommandButton Operator 
      Caption         =   "X"
      Height          =   360
      Index           =   2
      Left            =   2040
      TabIndex        =   17
      Top             =   960
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      Height          =   360
      Index           =   3
      Left            =   1080
      TabIndex        =   16
      Top             =   1920
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      Height          =   360
      Index           =   2
      Left            =   600
      TabIndex        =   15
      Top             =   1920
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      Height          =   360
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   360
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      Height          =   360
      Index           =   3
      Left            =   2040
      TabIndex        =   13
      Top             =   1440
      Width           =   360
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      Height          =   840
      Index           =   1
      Left            =   1560
      TabIndex        =   12
      Top             =   1440
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      Height          =   360
      Index           =   6
      Left            =   1080
      TabIndex        =   11
      Top             =   1440
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      Height          =   360
      Index           =   5
      Left            =   600
      TabIndex        =   10
      Top             =   1440
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      Height          =   360
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   360
   End
   Begin VB.CommandButton CancelEntry 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1920
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "C"
      Height          =   240
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      Height          =   360
      Index           =   9
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      Height          =   360
      Index           =   8
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Cancel          =   -1  'True
      Caption         =   "7"
      Height          =   360
      Index           =   7
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   360
   End
   Begin VB.CommandButton CmdOFF 
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   3
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "Copy"
      Height          =   240
      Left            =   1800
      TabIndex        =   2
      Top             =   3360
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   2880
   End
   Begin VB.CommandButton OneByXCmd 
      Caption         =   "1/X"
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   360
   End
   Begin VB.CommandButton SqrCmd 
      Caption         =   "Sqr"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   2280
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wfalgs As Long) As Long

Const hwnd_Topmost = -1
Const swp_Showwindow = &H40
Const swp_Drawframe = &H20

Option Explicit
Dim Retval
Dim Op1, Op2                ' Previously input operand.
Dim DmalFlag As Integer  ' Decimal point present yet?
Dim NumOps As Integer       ' Number of operands.
Dim LastInput               ' Indicate type of last keypress event.
Dim OpFlag                  ' Indicate pending operation.
Dim TempReadout

Private Sub Cancel_Click()

' Click event procedure for C (cancel) key.
' Reset the display and initializes variables.
    Readout = "0."
    Op1 = 0
    Op2 = 0
    ChkPlse
    
End Sub

Private Sub CancelEntry_Click()

' Click event procedure for CE (cancel entry) key.
frmCalculator.Caption = "Calculator & VAT"
    Readout = "0."
    DmalFlag = False
    LastInput = "CE"
    
End Sub

Private Sub CmdOFF_Click()

Readout.BackColor = &H0&
Timer1.Enabled = True

End Sub

Private Sub cmdVat_Click()

Readout = Round((Readout * 0.175), 2)

End Sub

Private Sub Dmal_Click()

' Click event procedure for decimal point (.) key.
' If last keypress was an operator, initialize
' readout to "0." Otherwise, append a decimal
' point to the display.
    If LastInput = "NEG" Then
        Readout = "-0."
    ElseIf LastInput <> "NUMS" Then
        Readout = "0."
    End If
    DmalFlag = True
    LastInput = "NUMS"
    FrmCaption
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 48 Then
    Number(0).SetFocus
        Number_Click (0)
ElseIf KeyAscii = 49 Then
Number(1).SetFocus
    Number_Click (1)
ElseIf KeyAscii = 50 Then
Number(2).SetFocus
    Number_Click (2)
ElseIf KeyAscii = 51 Then
    Number(3).SetFocus
    Number_Click (3)
ElseIf KeyAscii = 52 Then
    Number(4).SetFocus
    Number_Click (4)
ElseIf KeyAscii = 53 Then
    Number(5).SetFocus
    Number_Click (5)
ElseIf KeyAscii = 54 Then
    Number(6).SetFocus
    Number_Click (6)
ElseIf KeyAscii = 55 Then
    Number(7).SetFocus
    Number_Click (7)
ElseIf KeyAscii = 56 Then
    Number(8).SetFocus
    Number_Click (8)
ElseIf KeyAscii = 57 Then
    Number(9).SetFocus
    Number_Click (9)
ElseIf KeyAscii = 61 Then
    Operator(4).SetFocus
     Operator_Click (4)
                
ElseIf KeyAscii = 47 Then
Operator(0).SetFocus
   Operator_Click (0)
       
ElseIf KeyAscii = 42 Then
    Operator(2).SetFocus
    Operator_Click (2)
     
ElseIf KeyAscii = 45 Then
    Operator(3).SetFocus
    Operator_Click (3)
     
ElseIf KeyAscii = 43 Then
    Operator(1).SetFocus
    Operator_Click (1)
        
ElseIf KeyAscii = 99 Then
    Cancel.SetFocus
    Cancel_Click
ElseIf KeyAscii = 111 Then
    CmdOFF.SetFocus
    CmdOFF_Click
ElseIf KeyAscii = 37 Then
    Percent.SetFocus
    Percent_Click
     
ElseIf KeyAscii = 46 Then
   Dmal.SetFocus
    Dmal_Click
   
End If

FrmCaption

End Sub

Private Sub Form_Load()

' Initialization routine for the form.
' Set all variables to initial values.
Retval = SetWindowPos(Me.hwnd, hwnd_Topmost, 100, 100, 170, 220, swp_Drawframe Or swp_Showwindow)
ChkPlse
    
End Sub

Private Sub Number_Click(Index As Integer)

' Click event procedure for number keys (0-9).
' Append new number to the number in the display.

    If LastInput <> "NUMS" Then
        Readout = "."
        DmalFlag = False
    End If
    If DmalFlag Then
        Readout = Readout + Number(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, ".") - 1) + Number(Index).Caption + "."
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
    
FrmCaption

End Sub

Private Sub OneByXCmd_Click()

Readout.Caption = 1 / Val(Readout.Caption)

End Sub

Private Sub Operator_Click(Index As Integer)

' Click event procedure for operator keys (+, -, x, /, =).
' If the immediately preceeding keypress was part of a
' number, increments NumOps. If one operand is present,
' set Op1. If two are present, set Op1 equal to the
' result of the operation on Op1 and the current
' input string, and display the result.
TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
        If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
            Readout = "-" & Readout
            LastInput = "NEG"
        End If
        Case 1
        Op1 = Readout
        If Operator(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            Readout = "-"
            LastInput = "NEG"
        End If
        Case 2
        Op2 = TempReadout
        Select Case OpFlag
            Case "+"
                Op1 = Val(Op1) + Val(Op2)
            Case "-"
                Op1 = Op1 - Op2
            Case "X"
                Op1 = Op1 * Op2
            Case "/"
                If Op2 = 0 Then
                   MsgBox "Can't divide by zero", 48, "Calculator & VAT"
                Else
                   Op1 = Op1 / Op2
                End If
            Case "="
                Op1 = Op2
            Case "%"
                Op1 = Op1 * Op2
            End Select
        Readout = Op1
        NumOps = 1
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
    
    FrmCaption
    
End Sub

Private Sub Percent_Click()

' Click event procedure for percent key (%).
' Compute and display a percentage of the first operand.
Readout = Readout / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DmalFlag = True
    FrmCaption
    
End Sub

Sub FrmCaption()

If Readout.Caption = "0." Or Readout.Caption = "0" Then
frmCalculator.Caption = "Calculator & VAT"
Else
'frmCalculator.Caption = Readout.Caption
End If

End Sub

Sub ChkPlse()

DmalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    frmCalculator.Caption = "Calculator & VAT"

End Sub

Private Sub SqrCmd_Click()

Readout.Caption = Val(Readout.Caption) * Val(Readout.Caption)

End Sub

Private Sub Timer1_Timer()

frmCalculator.WindowState = 1
Unload Me

End Sub

