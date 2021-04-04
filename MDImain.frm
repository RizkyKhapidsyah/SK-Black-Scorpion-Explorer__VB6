VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.MDIForm MDImain 
   BackColor       =   &H8000000C&
   Caption         =   "Black Scorpion Explorer..."
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11880
   Icon            =   "MDImain.frx":0000
   LinkTopic       =   "MDImain"
   WindowState     =   2  'Maximized
   Begin PicClip.PictureClip pcProgress 
      Left            =   7320
      Top             =   3960
      _ExtentX        =   21114
      _ExtentY        =   1005
      _Version        =   393216
      Picture         =   "MDImain.frx":08CA
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   9960
      Top             =   6480
   End
   Begin VB.Timer tmrProgress 
      Interval        =   200
      Left            =   10320
      Top             =   6480
   End
   Begin VB.Timer tmrMedia1 
      Interval        =   1000
      Left            =   9960
      Top             =   6120
   End
   Begin VB.Timer tmrMedia2 
      Interval        =   1000
      Left            =   10320
      Top             =   6120
   End
   Begin VB.PictureBox picStatusbar 
      Align           =   2  'Align Bottom
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11880
      TabIndex        =   3
      Top             =   8010
      Width           =   11880
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   6840
         TabIndex        =   5
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   6
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   12012
               Picture         =   "MDImain.frx":16CC4
               Key             =   "URL"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   4471
               MinWidth        =   4480
               Key             =   "Progress"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               AutoSize        =   2
               Object.Width           =   1508
               MinWidth        =   1499
               TextSave        =   "31/03/02"
               Key             =   "Date"
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               Alignment       =   1
               AutoSize        =   2
               Object.Width           =   1058
               MinWidth        =   1058
               TextSave        =   "15:29"
               Key             =   "Time"
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   1
               AutoSize        =   2
               Enabled         =   0   'False
               Object.Width           =   873
               MinWidth        =   882
               TextSave        =   "CAPS"
               Key             =   "Caps Lock"
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   2
               AutoSize        =   2
               Object.Width           =   873
               MinWidth        =   882
               TextSave        =   "NUM"
               Key             =   "Nums Lock"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7020
      Left            =   0
      ScaleHeight     =   6960
      ScaleWidth      =   3015
      TabIndex        =   2
      Top             =   990
      Visible         =   0   'False
      Width           =   3075
      Begin MSComctlLib.Toolbar tbrHistory 
         Height          =   330
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         ButtonWidth     =   1799
         ButtonHeight    =   582
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "History..."
               Key             =   "tbrHistory"
               Description     =   "History"
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin SHDocVwCtl.WebBrowser webHistory 
         Height          =   6735
         Left            =   0
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
         ExtentX         =   5318
         ExtentY         =   11880
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
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
      Begin VB.ListBox LstHistory 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6690
         Left            =   0
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.ComboBox CboHistory 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   6600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton cmdOrgFav 
         Caption         =   "Organize"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddFav 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   14
         Top             =   2400
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ComboBox cboSearchBox 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtSearchBox 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   2745
      End
      Begin MSComctlLib.Toolbar tbrFavorites 
         Height          =   330
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         ButtonWidth     =   2090
         ButtonHeight    =   582
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Favorites..."
               Key             =   "tbrFavorites"
               Description     =   "Favorites"
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin MSComctlLib.Toolbar tbrSearch 
         Height          =   330
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         ButtonWidth     =   1852
         ButtonHeight    =   582
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Search..."
               Key             =   "tbrSearch"
               Description     =   "Search"
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6615
         Left            =   0
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   11668
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImageList2"
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame fraMedia 
         BackColor       =   &H80000014&
         Height          =   2295
         Left            =   0
         TabIndex        =   23
         Top             =   2520
         Visible         =   0   'False
         Width           =   3015
         Begin VB.TextBox txtMedia 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   480
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "Next"
            Height          =   255
            Left            =   960
            TabIndex        =   30
            Top             =   1800
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdPrev 
            Caption         =   "Prev"
            Height          =   255
            Left            =   1560
            TabIndex        =   29
            Top             =   1800
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "Stop"
            Height          =   255
            Left            =   1440
            TabIndex        =   28
            Top             =   1560
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdPause 
            Caption         =   "Pause"
            Height          =   255
            Left            =   2040
            TabIndex        =   27
            Top             =   1560
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdPlay 
            Caption         =   "Play"
            Height          =   255
            Left            =   840
            TabIndex        =   26
            Top             =   1560
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "Open"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1560
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ListBox List1 
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   1800
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSComctlLib.Slider Slider2 
            Height          =   255
            Left            =   480
            TabIndex        =   32
            Top             =   1200
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Max             =   2500
            TickStyle       =   3
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            TickStyle       =   3
         End
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   0
            TabIndex        =   35
            Top             =   120
            Visible         =   0   'False
            Width           =   2775
         End
         Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
            Height          =   255
            Left            =   2160
            TabIndex        =   34
            Top             =   1800
            Visible         =   0   'False
            Width           =   495
            AudioStream     =   -1
            AutoSize        =   0   'False
            AutoStart       =   -1  'True
            AnimationAtStart=   -1  'True
            AllowScan       =   -1  'True
            AllowChangeDisplaySize=   -1  'True
            AutoRewind      =   0   'False
            Balance         =   0
            BaseURL         =   ""
            BufferingTime   =   5
            CaptioningID    =   ""
            ClickToPlay     =   -1  'True
            CursorType      =   0
            CurrentPosition =   -1
            CurrentMarker   =   0
            DefaultFrame    =   ""
            DisplayBackColor=   0
            DisplayForeColor=   16777215
            DisplayMode     =   0
            DisplaySize     =   4
            Enabled         =   -1  'True
            EnableContextMenu=   -1  'True
            EnablePositionControls=   -1  'True
            EnableFullScreenControls=   0   'False
            EnableTracker   =   -1  'True
            Filename        =   ""
            InvokeURLs      =   -1  'True
            Language        =   -1
            Mute            =   0   'False
            PlayCount       =   1
            PreviewMode     =   0   'False
            Rate            =   1
            SAMILang        =   ""
            SAMIStyle       =   ""
            SAMIFileName    =   ""
            SelectionStart  =   -1
            SelectionEnd    =   -1
            SendOpenStateChangeEvents=   -1  'True
            SendWarningEvents=   -1  'True
            SendErrorEvents =   -1  'True
            SendKeyboardEvents=   0   'False
            SendMouseClickEvents=   0   'False
            SendMouseMoveEvents=   0   'False
            SendPlayStateChangeEvents=   -1  'True
            ShowCaptioning  =   0   'False
            ShowControls    =   -1  'True
            ShowAudioControls=   -1  'True
            ShowDisplay     =   0   'False
            ShowGotoBar     =   0   'False
            ShowPositionControls=   -1  'True
            ShowStatusBar   =   0   'False
            ShowTracker     =   -1  'True
            TransparentAtStart=   0   'False
            VideoBorderWidth=   0
            VideoBorderColor=   0
            VideoBorder3D   =   0   'False
            Volume          =   0
            WindowlessVideo =   0   'False
         End
      End
      Begin MSComctlLib.Toolbar tbrMedia 
         Height          =   330
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         ButtonWidth     =   1720
         ButtonHeight    =   582
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Media..."
               Key             =   "tbrMedia"
               Description     =   "Media"
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.Label lblSearch 
         BackStyle       =   0  'Transparent
         Caption         =   "Find a web page:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblSEngine 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose a Search Engine: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   2037
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Address :"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H0000FF00&
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   435
      End
      Begin VB.ComboBox cboURL 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   0
         Width           =   6405
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1111
      ButtonWidth     =   1323
      ButtonHeight    =   1058
      Appearance      =   1
      Style           =   1
      ImageList       =   "tbrImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "Back"
            Description     =   "Back"
            ImageIndex      =   1
            Style           =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            Description     =   "Forward"
            ImageIndex      =   2
            Style           =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "Stop"
            Description     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            Description     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Key             =   "Home"
            Description     =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            Description     =   "Search"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Favorites"
            Key             =   "Favorites"
            Description     =   "Favorites"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "History"
            Key             =   "History"
            Description     =   "History"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mail"
            Key             =   "Mail"
            Description     =   "Mail"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MC Mail"
                  Text            =   "Check Mail"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MS Mail"
                  Text            =   "Send Mail"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Media"
            Key             =   "Media"
            Description     =   "Media"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            Description     =   "Print"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   1
      Begin VB.PictureBox pbProgress 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   11280
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   36
         Top             =   0
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CDialog3 
      Left            =   10920
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList tbrImageList 
      Left            =   10680
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":171C8
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":17724
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":17C80
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":181DC
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":18738
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":18C94
            Key             =   "History"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1915C
            Key             =   "Favorites"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":195AC
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":19B08
            Key             =   "Mail"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1A00C
            Key             =   "Media"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1A134
            Key             =   "Print"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDialog2 
      Left            =   10440
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   10680
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1A5C0
            Key             =   "CloseFolder"
            Object.Tag             =   "CloseFolder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1AA14
            Key             =   "OpenFolder"
            Object.Tag             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1AE68
            Key             =   "URL"
            Object.Tag             =   "URL"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11160
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1B2BC
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1B738
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1BBBC
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1C084
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1C4B4
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1C9B8
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1CEBC
            Key             =   "Favorites"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1D30C
            Key             =   "History"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1D7D4
            Key             =   "Mail"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1DCD8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1E170
            Key             =   "Media"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1E2A0
            Key             =   "Media1"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDialog1 
      Left            =   9960
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin PicClip.PictureClip pcProgress1 
      Left            =   7320
      Top             =   4560
      _ExtentX        =   4445
      _ExtentY        =   4763
      _Version        =   393216
      Rows            =   6
      Cols            =   6
      Picture         =   "MDImain.frx":1E388
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNWindows 
         Caption         =   "&New Windows"
         Shortcut        =   ^N
      End
      Begin VB.Menu F1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu F2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveA 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu F3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPrintP 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu F4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuProp 
         Caption         =   "P&roperties"
      End
      Begin VB.Menu F5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIEport 
         Caption         =   "&Import and  Export"
      End
      Begin VB.Menu F6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWorkOffline 
         Caption         =   "Work Offline"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu E1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuABook 
         Caption         =   "Address &Book"
      End
      Begin VB.Menu mnuTBar 
         Caption         =   "&Toolbars"
         Begin VB.Menu mnuSButtons 
            Caption         =   "&Standard Buttons"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuUrlToolbar 
            Caption         =   "&URL Toolbar"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuDisable 
            Caption         =   "&Disable Pop-Ups"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuExplorerB 
         Caption         =   "&Explorer Bar"
         Begin VB.Menu mnuESearch 
            Caption         =   "&Search"
            Shortcut        =   ^E
         End
         Begin VB.Menu mnuFav 
            Caption         =   "&Favorites"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuEHistory 
            Caption         =   "&History"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuMPlayer 
            Caption         =   "&Media Player"
         End
      End
      Begin VB.Menu V1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoTo 
         Caption         =   "G&o To"
         Begin VB.Menu mnuBack 
            Caption         =   "&Back"
         End
         Begin VB.Menu mnuForward 
            Caption         =   "&Forward"
         End
         Begin VB.Menu V2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHpage 
            Caption         =   "Home Page"
         End
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Sto&p"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu V4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextsize 
         Caption         =   "Te&xt Size"
         Begin VB.Menu mnuLargest 
            Caption         =   "Lar&gest"
         End
         Begin VB.Menu mnuLarger 
            Caption         =   "&Larger"
         End
         Begin VB.Menu mnuMedium 
            Caption         =   "&Medium"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSmaller 
            Caption         =   "&Smaller"
         End
         Begin VB.Menu mnuSmallest 
            Caption         =   "&Smallest"
         End
      End
      Begin VB.Menu V5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSource 
         Caption         =   "Sour&ce"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuWUpdate 
         Caption         =   "&Windows Update"
      End
      Begin VB.Menu mnuIEUpdate 
         Caption         =   "Internet Explorer &Update"
      End
      Begin VB.Menu T1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIOptions 
         Caption         =   "Internet &Options"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuExcite 
         Caption         =   "&Excite"
      End
      Begin VB.Menu mnuYahoo 
         Caption         =   "&Yahoo!"
      End
      Begin VB.Menu mnuAVista 
         Caption         =   "&Alta Vista"
      End
      Begin VB.Menu mnuYPages 
         Caption         =   "Yellow &Pages"
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZNet 
         Caption         =   "&ZNet"
      End
      Begin VB.Menu mnuTucows 
         Caption         =   "&Tucows"
      End
   End
   Begin VB.Menu mnuIce 
      Caption         =   "Ice&berg Tip"
      Begin VB.Menu mnuCalcul 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu mnuSMS 
         Caption         =   "&SMS"
      End
      Begin VB.Menu IT1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBBank 
         Caption         =   "&Barclays Bank"
      End
      Begin VB.Menu mnuAbbeyN 
         Caption         =   "Abbey &National"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help..."
      Begin VB.Menu mnuConIndex 
         Caption         =   "&Contents and Index"
      End
      Begin VB.Menu mnuTDay 
         Caption         =   "Tip of the &Day"
      End
      Begin VB.Menu mnuNetscape 
         Caption         =   "&Netscape User"
      End
      Begin VB.Menu mnuTour 
         Caption         =   "&Tours"
      End
      Begin VB.Menu mnuOnlineSupport 
         Caption         =   "Online &Support"
      End
      Begin VB.Menu mnuSendfeedback 
         Caption         =   "Send Feedbac&k"
      End
      Begin VB.Menu H1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About!"
      End
   End
End
Attribute VB_Name = "MDImain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strURL As String
Dim dblReturn As Double
Public AllowPopup As Boolean
Public Ice As Boolean
Dim a
Dim iRotate As Integer

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim FP As FILE_PARAMS
Dim Itm As Node
Dim favpath As String
Dim sLastFolder As String
Dim sRoot As String
Dim bSubItem As Boolean
Dim nCount As Long
Dim bCancel As Boolean

Private Sub setFontSize(Size As Integer)
    
    Dim Range As Variant
    
    On Error Resume Next
    
    Range = CLng(Size)
    frmExplorer.wWeb.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, Range, Null
    
End Sub

Sub PercentBar(Shape As Control, Done As Variant, Total As Variant)

On Error Resume Next
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "Fixedsys"
Shape.FontSize = 8.25
Shape.FontBold = False
StatusBar1.Panels(1).Text = Done / Total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), &HC0C0C0, BF
Shape.Line (0, 0)-(StatusBar1.Panels(1).Text - 10, Shape.Height), &H800000, BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.ForeColor = RGB(0, 0, 0)

End Sub

Private Function GetFileInformation(FP As FILE_PARAMS) As Long

  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim pos As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   Dim sURL As String
   Dim sShortcut As String
   Dim itmX As ListItem
         
  'FP.sFileNameExt (assigned to sPath) contains
  'the full path and filespec.
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt
   
  'obtain handle to the first filespec match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then

      Do
      
        'remove trailing nulls
         sTmp = TrimNull(WFD.cFileName)

         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) _
            = FILE_ATTRIBUTE_DIRECTORY Then
           
           'determine the link name by removing
           'the .url extension
            pos = InStr(sTmp, ".url")
            
            If pos > 0 Then
            
                sShortcut = Left$(sTmp, pos - 1)
           
                'extract the URL
                sURL = ProfileGetItem("InternetShortcut", "URL", "", sRoot & sTmp)
                If sLastFolder = "" Then
                    'In The Root
                    Call LoadTreeView(sShortcut, False, False, "R", sURL)
                Else
                    Call LoadTreeView(sShortcut, False, False, sLastFolder, sURL)
                End If

            End If
         
         End If
         
      Loop While FindNextFile(hFile, WFD)
      
     'close the handle
      hFile = FindClose(hFile)
   
   End If
   
  'clean up
   Set itmX = Nothing
   
End Function

Private Function GetFolderPath(CSIDL As Long) As String

   Dim sPath As String
   Dim sTmp As String
  
  'fill pidl with the specified folder item
   sPath = Space$(MAX_LENGTH)
   
   If SHGetFolderPath(Me.hwnd, CSIDL, 0&, SHGFP_TYPE_CURRENT, sPath) = S_OK Then
       sTmp = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
   End If
   
   GetFolderPath = sTmp
   
End Function

Private Function InRoot(ByVal sPath As String) As Boolean
Dim sTmp As String

    InRoot = False
    
    sTmp = favpath & "\" & sPath
    
    If Dir(sTmp, vbDirectory) <> "" Then
        InRoot = True
    End If
    
End Function

Private Sub LoadTreeView(ItemName As String, bFolder As Boolean, bRoot As Boolean, _
    Optional SubItem As String, Optional sURL As String)
        
    If bRoot Then
        Set Itm = TreeView1.Nodes.Add(, , "R", ItemName, 3)
        Itm.Tag = FP.sFileRoot
        Exit Sub
    End If
    
    If bFolder Then
        If Len(SubItem) > 0 Then
            Set Itm = TreeView1.Nodes.Add(SubItem, tvwChild, ItemName, ItemName, 1, 2)
            
        Else
            Set Itm = TreeView1.Nodes.Add("R", tvwChild, ItemName, ItemName, 1, 2)
            
        End If
        
        Itm.Tag = FP.sFileRoot
        
    Else
        Set Itm = TreeView1.Nodes.Add(SubItem, tvwChild, ItemName & "_URL", ItemName, 3)
        Itm.Tag = sURL
        
    End If
       
End Sub

Public Function ProfileGetItem(lpSectionName As String, _
                               lpKeyName As String, _
                               defaultValue As String, _
                               inifile As String) As String
        
   Dim success As Long
   Dim nSize As Long
   Dim ret As String
  
  'Pad a string large enough to hold the data.
   ret = Space$(2048)
   nSize = Len(ret)
   success = GetPrivateProfileString(lpSectionName, lpKeyName, _
                                     defaultValue, ret, nSize, inifile)
   
   If success Then
      ProfileGetItem = Left$(ret, success)
   End If
   
End Function

Private Function QualifyPath(sPath As String) As String

  'assures that a passed path ends in a slash
   If Right$(sPath, 1) <> "\" Then
         QualifyPath = sPath & "\"
   Else: QualifyPath = sPath
   End If
      
End Function

Private Sub SearchForFilesArray(FP As FILE_PARAMS)
'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
      
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & "*.*"
   
  'obtain handle to the first match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then

      Call GetFileInformation(FP)

      Do
      
        'if the returned item is a folder...
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            
           'remove trailing nulls
            sTmp = TrimNull(WFD.cFileName)
            
           'and if the folder is not the default
           'self and parent folders...
            If sTmp <> "." And sTmp <> ".." Then

               FP.sFileRoot = sRoot & sTmp
              
              If InRoot(sTmp) Then
                Call LoadTreeView(sTmp, True, False)
                sLastFolder = sTmp
                  
              Else
                Call LoadTreeView(sTmp, True, False, sLastFolder)
                sLastFolder = sTmp
              End If

               If FP.sFileNameExt = "*.*" Then
                  
               End If
               
              'call again
               Call SearchForFilesArray(FP)
            
            End If
                  
         End If
         
     'continue looping until FindNextFile returns
     '0 (no more matches)
      Loop While FindNextFile(hFile, WFD)
      
     'close the find handle
      hFile = FindClose(hFile)
   
   End If
   
End Sub

Private Sub cboURL_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
frmExplorer.wWeb.Navigate cboURL.Text

End If

End Sub

Private Sub cboURL_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
frmExplorer.wWeb.Navigate cboURL.Text
cboURL.AddItem cboURL.Text

End If

End Sub

Private Sub cmdAddFav_Click()

    Dim shellHelper As New ShellUIHelper
    Dim strLocationName, strLocationURL As String

    strLocationName = frmExplorer.wWeb.LocationName
    strLocationURL = frmExplorer.wWeb.LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
    
End Sub

Private Sub cmdGo_Click()

If cboURL.Text = "" Then
        'do nothing
    Else
        frmExplorer.wWeb.Navigate (cboURL.Text)
    End If
    strURL = cboURL.Text
    If Left(LCase(strURL), 7) = "http://" Or Left(LCase(strURL), 6) = "ftp://" Then
        cboURL.Text = strURL
    Else
        If Left(strURL, 7) <> "http://" Then
            cboURL.Text = "http://" & strURL
        Else
            If Left(strURL, 6) = "ftp://" Then
                cboURL.Text = "ftp://" & strURL
            End If
        End If
    End If
    cboURL.SelStart = 0
    cboURL.SelLength = Len(cboURL.Text)
    
End Sub

Private Sub cmdNext_Click()

On Error Resume Next
List1.ListIndex = List1.ListIndex + 1
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
txtMedia.Text = MediaPlayer1.FileName
End Sub

Private Sub cmdOpen_Click()

On Error Resume Next
CDialog2.Filter = "Audio Files|*.wav;*.mid;*.mp3;mp2;*.mod|"
CDialog2.Flags = cdlOFNHideReadOnly
CDialog2.CancelError = True
CDialog2.DialogTitle = "Choose an mediafile to open"
CDialog2.FileName = ""
CDialog2.ShowOpen

List1.AddItem CDialog2.FileName
List1.ListIndex = List1.ListIndex + 1
MediaPlayer1.FileName = CDialog2.FileName
txtMedia.Text = CDialog2.FileName

End Sub

Private Sub cmdOrgFav_Click()

'Organize Favorites
   Dim lpszRootFolder As String
   Dim success As Long
  
   lpszRootFolder = GetFolderPath(&H6)
   success = DoOrganizeFavDlg(hwnd, lpszRootFolder)
   
   cboURL.Refresh
   
End Sub

Private Sub cmdPause_Click()

On Error Resume Next
If MediaPlayer1.PlayState = mpPlaying Then
MediaPlayer1.Pause
Else
MediaPlayer1.Play
End If

End Sub

Private Sub cmdPlay_Click()

On Error Resume Next
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
txtMedia.Text = MediaPlayer1.FileName

End Sub
Private Sub cmdPrev_Click()

On Error Resume Next
List1.ListIndex = List1.ListIndex - 1
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
txtMedia.Text = MediaPlayer1.FileName
End Sub

Private Sub cmdSearch_Click()

If cboSearchBox.Text = "Select a Search Engine" Then
MsgBox "Please select a search engine to search", vbInformation, "Select Engine"
End If
If txtSearchBox.Text = "" Then
MsgBox "Please enter at least 1 word to search for", vbInformation, "Enter Word"
End If

On Error Resume Next
Select Case cboSearchBox.Text
    
Case "MSN"
    frmExplorer.wWeb.Navigate ("http://search.msn.com/results.asp?RS=CHECKED&FORM=MSNH&v=1&q=" & txtSearchBox.Text)
Case "Excite"
    frmExplorer.wWeb.Navigate ("http://search.excite.com/search.gw?search=" & txtSearchBox.Text)
Case "Google"
    frmExplorer.wWeb.Navigate ("http://www.google.com/search?q=" & txtSearchBox.Text & "&meta=lr%3D%26hl%3Den&btnG=Google+Search")
Case "Yahoo"
    frmExplorer.wWeb.Navigate ("http://ink.yahoo.com/bin/query?p=" & txtSearchBox.Text & "&z=2&hc=0&hs=0")
Case "Altavista"
    frmExplorer.wWeb.Navigate ("http://www.altavista.com/cgi-bin/query?pg=q&kl=XX&stype=stext&q=" & txtSearchBox.Text)
Case "Lycos"
    frmExplorer.wWeb.Navigate ("http://www.lycos.com/srch/?lpv=1&loc=searchhp&query=" & txtSearchBox.Text)
Case "About.COM"
    frmExplorer.wWeb.Navigate ("http://search.about.com/fullsearch.htm?terms=" & "&PM=59_0100_S&Action.x=9&Action.y=7 ")

End Select

End Sub

Private Sub cmdStop_Click()

On Error Resume Next
MediaPlayer1.Stop

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

Dim Tip As Integer
    Tip = MsgBox("Do you really want to Quit the Black Scorpion Internet Explorer?", vbYesNo + vbCritical, "About Internet Browser Program...")
If Tip = vbYes Then End
If Tip = vbNo Then Cancel = 1

Exit Sub
End Sub

Private Sub mnuAbbeyN_Click()

frmExplorer.wWeb.Navigate "http://www.abbeynational.co.uk/index_flash.html"

End Sub

Private Sub mnuABook_Click()

On Error Resume Next
Dim X
X = Shell("C:\Program Files\Outlook Express\wab.exe", 1)

End Sub

Private Sub mnuAbout_Click()

    frmIntro.Show
    
End Sub

Private Sub mnuAVista_Click()

    frmExplorer.wWeb.Navigate "http://www.altavista.com/cgi-bin/query?pg=q&what=web&fmt=.&q="
    
End Sub

Private Sub mnuBack_Click()

On Error Resume Next
Ice = False
 For a = 1 To Toolbar1.Buttons("Forward").ButtonMenus.Count
            If Toolbar1.Buttons("Forward").ButtonMenus.Item(a).Text = frmExplorer.wWeb.LocationURL Then
                Ice = True
            End If
        Next a
        If Ice = False Then Toolbar1.Buttons("Forward").ButtonMenus.Add Text:=frmExplorer.wWeb.LocationURL
        frmExplorer.wWeb.GoBack
        
End Sub

Private Sub mnuBBank_Click()

    frmExplorer.wWeb.Navigate "https://ibankon.barclays.co.uk/"
End Sub

Private Sub mnuCalcul_Click()

frmCalculator.Show

End Sub

Private Sub mnuClose_Click()

    On Error Resume Next
    Unload ActiveForm
    
    If ActiveForm Is Nothing Then
        mnuClose.Enabled = False
    Else
        mnuClose.Enabled = True
    End If
    Exit Sub

End Sub

Private Sub mnuConIndex_Click()

CDialog3.HelpFile = "Iexplore.HLP"
CDialog3.HelpCommand = cdlHelpContents
CDialog3.ShowHelp

End Sub

Private Sub mnuCopy_Click()

    frmExplorer.wWeb.SetFocus
    On Error Resume Next
    frmExplorer.wWeb.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuCut_Click()

    frmExplorer.wWeb.SetFocus
    On Error Resume Next
    frmExplorer.wWeb.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuDelete_Click()

    frmExplorer.wWeb.SetFocus
    On Error Resume Next
    frmExplorer.wWeb.ExecWB OLECMDID_DELETE, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuDisable_Click()

    On Error Resume Next
    mnuDisable.Checked = Not mnuDisable.Checked

        If mnuDisable.Checked = True Then
            SaveSetting App.Title, "Settings", "Popups", mnuDisable.Checked
        ElseIf mnuDisable.Checked = False Then
            SaveSetting App.Title, "Settings", "Popups", mnuDisable.Checked
        End If
End Sub

Private Sub mnuEHistory_Click()

            tbrHistory.Visible = False
            Picture1.Visible = False
            LstHistory.Visible = False
            CboHistory.Visible = False
            webHistory.Visible = False
    Toolbar1.Buttons(12).Value = tbrUnpressed
    
End Sub

Private Sub mnuESearch_Click()

            tbrSearch.Visible = False
            Picture1.Visible = False
            txtSearchBox.Visible = False
            cboSearchBox.Visible = False
            cmdSearch.Visible = False
            lblSearch.Visible = False
            lblSEngine.Visible = False
    Toolbar1.Buttons(9).Value = tbrUnpressed
    
End Sub

Private Sub mnuExcite_Click()

frmExplorer.wWeb.Navigate "http://search.excite.com/search/"

End Sub

Private Sub mnuExit_Click()

Dim Iceberg As Integer
    Iceberg = MsgBox("Do you really want to Exit from Iceberg Tip Internet Explorer?", vbYesNo + vbQuestion, "About This Program...")
If Iceberg = vbYes Then
   End
Else
    Iceberg = vbNo
End If

End Sub

Private Sub mnuFav_Click()

    Picture1.Visible = False
    TreeView1.Visible = False
    tbrFavorites.Visible = False
    cmdAddFav.Visible = False
    cmdOrgFav.Visible = False
    Toolbar1.Buttons(11).Value = tbrUnpressed
End Sub

Private Sub mnuForward_Click()

        For a = 1 To Toolbar1.Buttons("Back").ButtonMenus.Count
        If Toolbar1.Buttons("Back").ButtonMenus.Item(a).Text = frmExplorer.wWeb.LocationURL Then
            Ice = True
        End If
    Next a
    If Ice = False Then Toolbar1.Buttons("Back").ButtonMenus.Add Text:=frmExplorer.wWeb.LocationURL
       frmExplorer.wWeb.GoForward
End Sub

Private Sub mnuHpage_Click()

frmExplorer.wWeb.SetFocus
    
End Sub

Private Sub mnuIEport_Click()

frmIEport.Show

End Sub

Private Sub mnuIEUpdate_Click()

    frmExplorer.wWeb.Navigate "http://www.microsoft.com/windows/ie/downloads/recommended/128bit/default.asp"
    
End Sub

Private Sub mnuIOptions_Click()

dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)

End Sub

Private Sub mnuLarger_Click()

    setFontSize (3)
    mnuLargest.Checked = False
    mnuLarger.Checked = True
    mnuMedium.Checked = False
    mnuSmaller.Checked = False
    mnuSmallest.Checked = False
    
End Sub

Private Sub mnuLargest_Click()
    
    setFontSize (4)
    mnuLargest.Checked = True
    mnuLarger.Checked = False
    mnuMedium.Checked = False
    mnuSmaller.Checked = False
    mnuSmallest.Checked = False
    
End Sub

Private Sub mnuMedium_Click()

    setFontSize (2)
    mnuLargest.Checked = False
    mnuLarger.Checked = False
    mnuMedium.Checked = True
    mnuSmaller.Checked = False
    mnuSmallest.Checked = False
    
End Sub

Private Sub mnuMPlayer_Click()

    tbrMedia.Visible = False
    Picture1.Visible = False
    fraMedia.Visible = False
    lblTime.Visible = False
    txtMedia.Visible = False
    Picture1.Visible = False
    Slider1.Visible = False
    Slider2.Visible = False
    cmdOpen.Visible = False
    cmdPlay.Visible = False
    cmdStop.Visible = False
    cmdPause.Visible = False
    cmdNext.Visible = False
    cmdPrev.Visible = False
    Toolbar1.Buttons(15).Value = tbrUnpressed
           
End Sub

Private Sub mnuNetscape_Click()

CDialog3.HelpFile = "Iexplore.HLP"
CDialog3.HelpCommand = cdlHelpContents
CDialog3.ShowHelp

End Sub

Private Sub mnuNWindows_Click()

    On Error Resume Next
    Static lDocumentCount As Long
    Dim IB2 As Form
    lDocumentCount = lDocumentCount + 1
    Set IB2 = New frmExplorer
    IB2.Show
    IB2.SetFocus

End Sub

Private Sub mnuOnlineSupport_Click()

frmExplorer.wWeb.Navigate "http://support.microsoft.com/default.aspx?LN=EN-US"

End Sub

Private Sub mnuOpen_Click()

frmExplorer.wWeb.SetFocus
    On Error Resume Next
    CDialog1.Filter = "All Internet Files (*.hmt,*.html,*.asp,*.shtml,*.js,*.dhtml) | *.htm;*.html;*.asp;*.shtml;*.js;*.dhtml"
    CDialog1.ShowOpen
    If CDialog1.FileName = "" Then
        Exit Sub
        Else
    frmExplorer.wWeb.Navigate (CDialog1.FileName)
        End If

End Sub

Private Sub mnuPaste_Click()

frmExplorer.wWeb.SetFocus
On Error Resume Next
frmExplorer.wWeb.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuPrint_Click()

    frmExplorer.wWeb.SetFocus
    On Error Resume Next
    frmExplorer.wWeb.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuPrintP_Click()

    frmExplorer.wWeb.SetFocus
    On Error Resume Next
    frmExplorer.wWeb.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuProp_Click()

    frmExplorer.wWeb.SetFocus
    On Error Resume Next
    frmExplorer.wWeb.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuPSetup_Click()

    frmExplorer.wWeb.SetFocus
    On Error Resume Next
    frmExplorer.wWeb.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuRefresh_Click()

frmExplorer.wWeb.SetFocus
frmExplorer.wWeb.Refresh

End Sub

Private Sub mnuSAll_Click()

    frmExplorer.wWeb.SetFocus
    On Error Resume Next
    frmExplorer.wWeb.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuSave_Click()

frmExplorer.wWeb.SetFocus
    On Error Resume Next
    frmExplorer.wWeb.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuSaveA_Click()

frmExplorer.wWeb.SetFocus
    On Error Resume Next
    CDialog1.Filter = "htm (*.htm) | *.htm"
    CDialog1.ShowSave
    If CDialog1.FileName = "" Then
        Exit Sub
        Else
    Open CDialog1.FileName For Output As #1
        End If

End Sub

Private Sub mnuSButtons_Click()

mnuSButtons.Checked = Not mnuSButtons.Checked
        
If mnuSButtons.Checked = True Then
       Toolbar1.Visible = mnuSButtons.Checked
ElseIf mnuSButtons.Checked = False Then
        Toolbar1.Visible = mnuSButtons.Checked
End If

End Sub

Private Sub mnuSendfeedback_Click()

frmExplorer.wWeb.Navigate "http://register.microsoft.com/contactus30/contactus.asp?domain=generic"

End Sub

Private Sub mnuSmaller_Click()

    setFontSize (1)
    mnuLargest.Checked = False
    mnuLarger.Checked = False
    mnuMedium.Checked = False
    mnuSmaller.Checked = True
    mnuSmallest.Checked = False
End Sub

Private Sub mnuSmallest_Click()

    setFontSize (0)
    mnuLargest.Checked = False
    mnuLarger.Checked = False
    mnuMedium.Checked = False
    mnuSmaller.Checked = False
    mnuSmallest.Checked = True
    
End Sub

Private Sub mnuSMS_Click()

frmSMS.Show

End Sub

Private Sub mnuSource_Click()

    On Error Resume Next
    
    If Len(frmExplorer.wWeb.Document.documentElement.innerHTML) > 0 Then
        frmExplorer.wWeb.Navigate frmExplorer.wWeb.Document.documentElement.innerHTML
    End If
    
End Sub

Private Sub mnuStop_Click()

frmExplorer.wWeb.Stop

End Sub

Private Sub mnuTDay_Click()

frmTip.Show

End Sub

Private Sub mnuTour_Click()

frmExplorer.wWeb.Navigate "http://www.microsoft.com/insider/internet/default.htm"

End Sub

Private Sub mnuTucows_Click()

    frmExplorer.wWeb.Navigate "http://www.tucows.com"
    
End Sub

Private Sub mnuUrlToolbar_Click()

mnuUrlToolbar.Checked = Not mnuUrlToolbar.Checked
        
If mnuUrlToolbar.Checked = True Then
       Toolbar2.Visible = mnuUrlToolbar.Checked
ElseIf mnuUrlToolbar.Checked = False Then
        Toolbar2.Visible = mnuUrlToolbar.Checked
End If

End Sub

Private Sub mnuWorkOffline_Click()

On Error Resume Next
    
    'If checked then uncheck, vice versa
    mnuWorkOffline.Checked = Not mnuWorkOffline.Checked

If mnuWorkOffline.Checked = True Then
        
        mnuWorkOffline.Caption = "Work Offline"
        frmExplorer.wWeb.Offline = True
        Timer2.Enabled = True
        
ElseIf mnuWorkOffline.Checked = False Then
        
        mnuWorkOffline.Caption = "Go Online"
        frmExplorer.wWeb.Offline = False
        Timer2 = True
    
End If

End Sub

Private Sub mnuWUpdate_Click()

    frmExplorer.wWeb.Navigate "http://windowsupdate.microsoft.com/"
    
End Sub

Private Sub mnuYahoo_Click()

    frmExplorer.wWeb.Navigate "http://search.yahoo.com/bin/search?p="
    
End Sub

Private Sub mnuYPages_Click()

frmExplorer.wWeb.Navigate "http://www.yell.co.uk/"

End Sub

Private Sub mnuZNet_Click()

frmExplorer.wWeb.Navigate "http://www.hotfiles.com"

End Sub

Private Sub Slider1_Click()

On Error Resume Next
MediaPlayer1.CurrentPosition = Slider1.Value

End Sub

Private Sub Slider2_Click()

Dim a As Integer, b As Integer
Dim d, c
c = Slider2.Value - 2500
MediaPlayer1.Volume = c
b = Slider2.Min
a = Slider2.Value

End Sub

Private Sub tmrMedia1_Timer()

If MediaPlayer1.PlayState = mpPlaying Then
lblTime.Caption = ConvertTime(Round(MediaPlayer1.CurrentPosition, 0)) & " / " & ConvertTime(Round(MediaPlayer1.Duration, 0))
Else
lblTime.Caption = "00:00:00 / 0:00:00"
End If

End Sub

Private Sub tmrMedia2_Timer()

On Error Resume Next
Slider1.Max = MediaPlayer1.Duration
Slider1.Value = MediaPlayer1.CurrentPosition

End Sub

Function ConvertTime(i As Integer)
Dim Secs As Integer
Dim Mins As Integer
Dim Hours As Integer
Secs = i Mod 60
Mins = Int(i / 60) Mod 60
Hours = Int(i / 3600)
If Secs < 10 Then Secs = "0" & Secs
If Mins < 10 Then Mins = "0" & Mins
ConvertTime = Hours & ":" & Mins & ":" & Secs
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim starting
On Error Resume Next

Select Case Button.Key
    Case "Back"
        mnuBack_Click
    Case "Forward"
        mnuForward_Click
    Case "Stop"
        frmExplorer.wWeb.Stop
    Case "Refresh"
        frmExplorer.wWeb.Refresh
    Case "Home"
        frmExplorer.wWeb.SetFocus
        frmExplorer.wWeb.GoHome
    Case "Search"
       If tbrSearch.Visible = True Then
        Toolbar1.Buttons(9).Value = tbrUnpressed
            tbrSearch.Visible = False
            Picture1.Visible = False
            txtSearchBox.Visible = False
            cboSearchBox.Visible = False
            cmdSearch.Visible = False
            lblSearch.Visible = False
            lblSEngine.Visible = False
        Else
        Toolbar1.Buttons(9).Value = tbrPressed
            tbrSearch.Visible = True
            Picture1.Visible = True
            txtSearchBox.Visible = True
            cboSearchBox.Visible = True
            cmdSearch.Visible = True
            lblSearch.Visible = True
            lblSEngine.Visible = True
        End If
    Case "Favorites"
        If TreeView1.Visible = True Then
        Toolbar1.Buttons(11).Value = tbrUnpressed
            TreeView1.Visible = False
            Picture1.Visible = False
            tbrFavorites.Visible = False
            cmdAddFav.Visible = False
            cmdOrgFav.Visible = False
        Else
        Toolbar1.Buttons(11).Value = tbrPressed
            TreeView1.Visible = True
            Picture1.Visible = True
            tbrFavorites.Visible = True
            cmdAddFav.Visible = True
            cmdOrgFav.Visible = True
        End If
    Case "History"
        If tbrHistory.Visible = True Then
        Toolbar1.Buttons(12).Value = tbrUnpressed
            tbrHistory.Visible = False
            Picture1.Visible = False
            LstHistory.Visible = False
            CboHistory.Visible = False
            webHistory.Visible = False
        Else
        Toolbar1.Buttons(12).Value = tbrPressed
            tbrHistory.Visible = True
            Picture1.Visible = True
            LstHistory.Visible = True
            CboHistory.Visible = True
            webHistory.Visible = True
            CboHistory = ("C:\WINDOWS\History\")
            webHistory.Navigate CboHistory.Text
        End If
    
    Case "Media"
        If tbrMedia.Visible = True Then
        Toolbar1.Buttons(15).Value = tbrUnpressed
            tbrMedia.Visible = False
            Picture1.Visible = False
            fraMedia.Visible = False
            lblTime.Visible = False
            txtMedia.Visible = False
            Picture1.Visible = False
            Slider1.Visible = False
            Slider2.Visible = False
            cmdOpen.Visible = False
            cmdPlay.Visible = False
            cmdStop.Visible = False
            cmdPause.Visible = False
            cmdNext.Visible = False
            cmdPrev.Visible = False
        Else
        Toolbar1.Buttons(15).Value = tbrPressed
            tbrMedia.Visible = True
            Picture1.Visible = True
            fraMedia.Visible = True
            lblTime.Visible = True
            txtMedia.Visible = True
            Picture1.Visible = True
            Slider1.Visible = True
            Slider2.Visible = True
            cmdOpen.Visible = True
            cmdPlay.Visible = True
            cmdStop.Visible = True
            cmdPause.Visible = True
            cmdNext.Visible = True
            cmdPrev.Visible = True
        End If
    
    Case "Print"
        frmExplorer.wWeb.SetFocus
        On Error Resume Next
        frmExplorer.wWeb.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
        
End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

On Error Resume Next
Select Case ButtonMenu.Key

'Check Mail
Case "MC Mail"
    Shell "C:\Program Files\Outlook Express\MSIMN.EXE", vbNormalFocus
   
'Send Mail
Case "MS Mail"
    Dim subject, person
        person = InputBox("Enter email address", "email")
        subject = InputBox("Enter subject for email", "subject")
        frmExplorer.wWeb.Navigate ("mailto:" & person & "?subject=" & subject)
    
End Select

End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyF5 Then
         TreeView1.Nodes.Clear
         TreeView1.Refresh
         
         'retrieve the special folder path
         'to the internet favorites
         favpath = GetFolderPath(CSIDL_FAVORITES)
         
         'Initializes the Root Item in the TreeView
         Call LoadTreeView("Internet Favorites", True, True)
        
         If Len(favpath) > 0 Then
        
          'set up the search UDT
           With FP
              .sFileRoot = favpath
              .sFileNameExt = "*.url"
              .bRecurse = True
           End With
           
          'get the files
           Call SearchForFilesArray(FP)
           TreeView1.Nodes("R").Expanded = True
         Else
         
            MsgBox " Could not locate favorites folder! " & _
                "This program requires Microsoft's Internet " & _
                "Explorer to be installed. Program will shutdown now!", _
                vbCritical + vbOKOnly, "FavMenu Error"
            End
         End If
    End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

'Navigate current Tab\Browser to the selected URL
    If Right(Node.Key, 4) = "_URL" Then
        Set Itm = Node
        frmExplorer.wWeb.Navigate Itm.Tag
    End If

End Sub


Private Sub MDIForm_Load()

    frmExplorer.wWeb.Navigate "http://www.msn.co.uk"
    cboSearchBox.Text = "MSN"

    'Set up working animaiton.....
    pcProgress.Cols = 21
    pcProgress.Rows = 1
    pbProgress.ScaleMode = vbPixels
    pbProgress.Picture = pcProgress.GraphicCell(0)
    
    'Searching.....
    cboSearchBox.AddItem "Google"
    cboSearchBox.AddItem "Yahoo"
    cboSearchBox.AddItem "Lycos"
    cboSearchBox.AddItem "Altavista"
    cboSearchBox.AddItem "MSN"
    cboSearchBox.AddItem "About.COM"
    cboSearchBox.AddItem "Excite"
    
    TreeView1.Nodes.Clear
    TreeView1.Refresh
    
    'retrieve the special folder path
    'to the internet favorites
    favpath = GetFolderPath(CSIDL_FAVORITES)
    
    'Initializes the Root Item in the TreeView
    Call LoadTreeView("Internet Favorites", True, True)
   
    If Len(favpath) > 0 Then
   
     'set up the search UDT
      With FP
         .sFileRoot = favpath
         .sFileNameExt = "*.url"
         .bRecurse = True
      End With
      
     'get the files
      Call SearchForFilesArray(FP)
      TreeView1.Nodes("R").Expanded = True
    Else
         
       MsgBox " Could not locate favorites folder! " & _
           "This program requires Microsoft's Internet " & _
           "Explorer to be installed. Program will shutdown now!", _
           vbCritical + vbOKOnly, "FavMenu Error"
       End

    End If
    
End Sub

Private Sub tmrProgress_Timer()

    statusPic
    
End Sub

Sub statusPic()
    
    iRotate = iRotate + 1
    If iRotate = 21 Then
        iRotate = 0
    End If
    pbProgress.Picture = pcProgress.GraphicCell(iRotate)
    
End Sub


