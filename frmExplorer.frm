VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmExplorer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   5160
   ClientLeft      =   5175
   ClientTop       =   4020
   ClientWidth     =   6705
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExplorer.frx":0000
   LinkTopic       =   "MDImain"
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   6705
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser wWeb 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      ExtentX         =   9128
      ExtentY         =   8493
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()

    On Error Resume Next
    
    'this method is not acurate for real resizing...:P
    'wWeb.Move 50, 50, Me.ScaleWidth - 100, Me.ScaleHeight - 100

    wWeb.Left = 0
    wWeb.Top = 0
    wWeb.Width = ScaleWidth '- 90
    wWeb.Height = ScaleHeight '- 70

End Sub

Private Sub wWeb_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

MDImain.cboURL.AddItem URL

End Sub

Private Sub wWeb_DocumentComplete(ByVal pDisp As Object, URL As Variant)

MDImain.cboURL.Text = wWeb.LocationURL
Me.Caption = wWeb.LocationName

End Sub

Private Sub wWeb_DownloadBegin()

    'Start working animation
    MDImain.tmrProgress.Enabled = True
    
End Sub

Private Sub wWeb_DownloadComplete()

    'Stop working animation
    MDImain.tmrProgress.Enabled = False
    
End Sub

Private Sub wWeb_NavigateComplete2(ByVal pDisp As Object, URL As Variant)

    frmExplorer.Caption = wWeb.LocationName
    MDImain.StatusBar1.Panels(1).Text = wWeb.LocationURL
    MDImain.cboURL.Text = URL

End Sub

Private Sub wWeb_NewWindow2(ppDisp As Object, Cancel As Boolean)

On Error GoTo NewWindow2_Error:
    
    If MDImain.mnuDisable.Checked = True Then
                        
                        'if checked then don't
        Cancel = True   'allow new window, else
                        'allow them to open...
                      
    ElseIf MDImain.mnuDisable.Checked = False Then
    
        Dim CapturedURL As String
        Dim WBrowser As New frmExplorer
        Set WBrowser = New frmExplorer
        Set ppDisp = WBrowser.wWeb.object
                   
        CapturedURL = MDImain.StatusBar1.Panels(1).Text

        WBrowser.Caption = wWeb.LocationName
        WBrowser.Show
    
    End If
    
        Set WBrowser = Nothing
        
    Exit Sub
    
NewWindow2_Error:
End Sub

Private Sub wWeb_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)

On Error Resume Next
MDImain.ProgressBar1.Max = ProgressMax
MDImain.ProgressBar1.Value = Progress
MDImain.ProgressBar1.Refresh

End Sub

Private Sub wWeb_StatusTextChange(ByVal Text As String)

MDImain.StatusBar1.Panels(1).Text = Text

End Sub

Private Sub wWeb_TitleChange(ByVal Text As String)

Me.Caption = Text

End Sub


