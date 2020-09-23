VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14925
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   14925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameView 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12120
      TabIndex        =   18
      Top             =   5280
      Width           =   1695
      Begin VB.CheckBox GridLineschk 
         Caption         =   "GridLines"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox FullRowSelectchk 
         Caption         =   "Full Row Select"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.CheckBox LogCache 
      Caption         =   "Save host names Cache"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   5640
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Frame FrameFileInfo 
      Caption         =   "File Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   10
      Top             =   5280
      Width           =   5895
      Begin VB.PictureBox PicFolder 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         Picture         =   "FrmMain.frx":1FF2
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Pic32 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   11
         Top             =   200
         Width           =   480
      End
      Begin VB.Label lblOpenFolder 
         Caption         =   "Open Parent Folder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   1650
      End
      Begin VB.Label LblSize 
         Caption         =   "Size :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   14
         Top             =   480
         Width           =   1650
      End
      Begin VB.Label LblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CheckBox ResolveHostchk 
      Caption         =   "Do not resolve host names"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   5280
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2640
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2334
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2650
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":296C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2C88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   4080
      Top             =   3360
   End
   Begin VB.Frame FrameTraffic 
      Caption         =   "Traffic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   5280
      Width           =   3375
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "FrmMain.frx":2FA4
         Top             =   200
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bytes received:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   8
         Top             =   195
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bytes sent:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1260
         TabIndex        =   7
         Top             =   435
         Width           =   825
      End
      Begin VB.Label lblRecv 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label lblSent 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   435
         Width           =   1095
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   4560
      Top             =   3360
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   6000
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3600
      Top             =   3360
   End
   Begin VB.PictureBox PicQuestion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      Picture         =   "FrmMain.frx":32AE
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Pic16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9240
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   240
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   3480
      TabIndex        =   0
      Top             =   60
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "Iml32"
      SmallIcons      =   "Iml16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Remote Host"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList TreeViewImgList 
      Left            =   960
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":35F0
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":39EA
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3D3C
            Key             =   "Ping"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":405E
            Key             =   "FileCross"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":43F0
            Key             =   "CancelCon"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":487A
            Key             =   "Cross"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4C0C
            Key             =   "Excal"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4FBE
            Key             =   "FileNet"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":5484
            Key             =   "QuestionComp"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":587E
            Key             =   "NotConnected"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":5BD0
            Key             =   "Connected"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":5F22
            Key             =   "HelpFile"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Iml16 
      Left            =   1440
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu PopMenu 
      Caption         =   "PopMenu"
      Visible         =   0   'False
      Begin VB.Menu CloseConnection 
         Caption         =   "Close Connection"
      End
      Begin VB.Menu CloseProgram 
         Caption         =   "Close Program"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 27
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const MAXLEN_IFDESCR = 256
Private Const MAXLEN_PHYSADDR = 8
Private Const MAX_INTERFACE_NAME_LEN = 256
Private nid As NOTIFYICONDATA

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long


Private m_objIpHelper As CIpHelper

Dim oOld As Long
Dim oNew As Long
Dim aOld As Long
Dim aNew As Long

Dim objInterface2 As CInterface
Dim obJHelper As CInterface
Dim tValue As Long
Dim aValue As Long

Private Unloaded As Boolean
Private Processing As Boolean
Private IsOnline As Boolean

Private TVHost As Long
Private TVPath As String
Private TVTAG As Long
Private TVPI As Long

Public iphDNS As New CDictionary

Private Sub CloseConnection_Click()
If TerminateThisConnection(TVTAG) = True Then
StatusBar.Panels(3).Text = "Connection by " & GetFileNameFromPath(TVPath) & " closed succesfully"
Else
StatusBar.Panels(3).Text = "Connection by " & GetFileNameFromPath(TVPath) & " failed to close"
End If
End Sub

Private Sub CloseProgram_Click()
If KillProcessById(TVPI) = True Then
StatusBar.Panels(3).Text = GetFileNameFromPath(TVPath) & " closed successfully"
Else
StatusBar.Panels(3).Text = GetFileNameFromPath(TVPath) & " failed to close"
End If
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Form_Initialize()
Dim X As Long
X = InitCommonControls
End Sub

Public Sub RefreshList()
  Dim i
  Dim Item As ListItem

If Unloaded = True Then Exit Sub

Processing = True

    RefreshStack
    DoEvents
    
    LoadNTProcess
    DoEvents

ListView1.ListItems.Clear
ListView1.Sorted = False

For i = 0 To GetEntryCount
     
    If Connection(i).State = "2" Then GoTo IsListening

    If Connection(i).FileName = "" Then
    Set Item = ListView1.ListItems.Add(, , "Unknown")
    Item.SubItems(5) = ""
    Else
    Set Item = ListView1.ListItems.Add(, , Right(Connection(i).FileName, Len(Connection(i).FileName) - InStrRev(Connection(i).FileName, "\")))
    Item.SubItems(5) = Connection(i).FileName
    End If
    
    Item.SubItems(1) = GetPort(Connection(i).LocalPort)
    
    If ResolveHostchk.Value = 1 Then Item.SubItems(2) = GetIPAddress(Connection(i).RemoteHost) Else Item.SubItems(2) = "Resolving..."
    
    Item.SubItems(3) = GetPort(Connection(i).RemotePort)
    Item.SubItems(4) = c_state(Connection(i).State)
    Item.Tag = i

IsListening:
Next i

ListView1.Sorted = True

GetAllIcons
DoEvents

ShowIcons
DoEvents

resolveIPs
DoEvents

Finished:
Processing = False
If Unloaded = True Then Unload Me
End Sub

Private Sub resolveIPs()
Dim Item As ListItem
    
If ResolveHostchk.Value = 1 Then Exit Sub

'On Local Error Resume Next
For Each Item In ListView1.ListItems
    If ResolveHostchk.Value = 1 Then GetIPAddress (Connection(Item.Tag).RemoteHost) Else Item.SubItems(2) = iphDNS.CheckDictionary(GetIPAddress(Connection(Item.Tag).RemoteHost))
  DoEvents
Next

End Sub

Private Function GetIcon(FileName As String, index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
On Error Resume Next
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long


If Connection(ListView1.ListItems(index).Tag).FileName = "" Then
Set imgObj = Iml16.ListImages.Add(index, , PicQuestion.Image)
Exit Function
End If


'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
'hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
'         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then

  'Large Icon
  'With Pic32
  '  Set .Picture = LoadPicture("")
  '  .AutoRedraw = True
  '  r = ImageList_Draw(hLIcon, ShInfo.iIcon, Pic32.hDC, 0, 0, ILD_TRANSPARENT)
  '  .Refresh
  'End With
  
    Else
  'Small Icon
  With Pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, Pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  
  Set imgObj = Iml16.ListImages.Add(index, , Pic16.Image)
End If

End Function

Private Function GetLargeIcon(FileName As String) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
On Error Resume Next
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long


If FileName = "" Then
'Set imgObj = Iml16.ListImages.Add(Index, , PicQuestion.Image)
Exit Function
End If


'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then

  'Large Icon
  With Pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, Pic32.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  
    Else

End If

End Function

Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the lvw
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With ListView1
  '.ListItems.Clear
  .SmallIcons = Iml16   'Small
  For Each Item In .ListItems
    Item.SmallIcon = Item.index
  Next
End With

End Sub

Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim FileName As String

    ListView1.SmallIcons = Nothing
    Iml16.ListImages.Clear
    
'On Local Error Resume Next
For Each Item In ListView1.ListItems
  FileName = Connection(Item.Tag).FileName

  GetIcon FileName, Item.index
   
Next

End Sub

Private Sub Form_Load()
Set m_objIpHelper = New CIpHelper
Dim FP As FILE_PARAMS
Dim CurFile As Long
Dim AppPath As String
Dim fso As New FileSystemObject
    
If IsNetConnectOnline() = True Then
    Timer2.Enabled = True
    
    StatusBar.Panels(4).Text = "Online"
    StatusBar.Panels(4).Picture = TreeViewImgList.ListImages("Connected").Picture
    
    IsOnline = True
    
    Else
    
    ListView1.ListItems.Clear
    
    Timer2.Enabled = False
    
    StatusBar.Panels(4).Text = "Offline"
    StatusBar.Panels(4).Picture = TreeViewImgList.ListImages("NotConnected").Picture
    
    IsOnline = False
End If

With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Connection Monitor (Online)"
    End With
    
   lblOpenFolder.Enabled = False
    
If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" & App.EXEName & ".exe" Else AppPath = App.Path & App.EXEName & ".exe"

TVPath = AppPath

GetLargeIcon AppPath

   With FP
      .sFileNameExt = AppPath
   End With
   
CurFile = GetFileInformation(FP)

FullRowSelectchk_Click
GridLineschk_Click

Shell_NotifyIcon NIM_ADD, nid
'Animation.Open App.Path & "\xpsearchinternet.avi"
'Animation.AutoPlay = True
End Sub

Private Sub Form_Resize()
On Error Resume Next

FrameTraffic.Left = 60
FrameTraffic.Top = Me.Height - 1850
FrameTraffic.Width = 3275

FrameFileInfo.Left = FrameTraffic.Width + (ResolveHostchk.Width + 100)
FrameFileInfo.Top = Me.Height - 1850
FrameFileInfo.Height = FrameTraffic.Height

FrameView.Top = Me.Height - 1850
FrameView.Height = FrameTraffic.Height
FrameView.Left = FrameFileInfo.Left + (FrameFileInfo.Width + 100)

ListView1.Width = Me.Width - 220
ListView1.Left = 60
ListView1.Height = FrameTraffic.Top - 60
ListView1.Top = 60

ResolveHostchk.Left = FrameTraffic.Width + 120
ResolveHostchk.Top = ListView1.Top + (ListView1.Height + 120)

LogCache.Left = FrameTraffic.Width + 120
LogCache.Top = ResolveHostchk.Top + (ResolveHostchk.Height + 50)

ListView1.ColumnHeaders(1).Width = 1300
ListView1.ColumnHeaders(2).Width = 1100
If ResolveHostchk.Value = 0 Then ListView1.ColumnHeaders(3).Width = 3500 Else ListView1.ColumnHeaders(3).Width = 1100
ListView1.ColumnHeaders(4).Width = 1100
ListView1.ColumnHeaders(5).Width = 1100
ListView1.ColumnHeaders(6).Width = ListView1.Width \ 2 + 1000

StatusBar.Panels(1).Width = 2000
StatusBar.Panels(2).Width = 2000
StatusBar.Panels(3).Width = Me.Width - 5500
StatusBar.Panels(4).Width = 1500


End Sub

Private Sub Form_Unload(Cancel As Integer)
StatusBar.Panels(3).Text = "Closing..."
    
If Processing = True Then
Unloaded = True
Cancel = -1
Exit Sub
End If

If LogCache.Value = 1 Then iphDNS.WriteCache
DoEvents

Shell_NotifyIcon NIM_DELETE, nid

End
End Sub



Private Sub FrameFileInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOpenFolder.FontBold = False
End Sub

Private Sub FullRowSelectchk_Click()
   Dim rStyle As Long
   Dim r As Long
   
  'get the current ListView style
   rStyle = SendMessageLong(ListView1.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)

   If FullRowSelectchk.Value = 0 Then
     'remove the extended style bit
      rStyle = rStyle Xor LVS_EX_FULLROWSELECT
    
   ElseIf FullRowSelectchk.Value = 1 Then
     'set the extended style bit
      rStyle = rStyle Or LVS_EX_FULLROWSELECT
    
   End If
   
  'set the new ListView style
   r = SendMessageLong(ListView1.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)

End Sub

Private Sub GridLineschk_Click()

   Dim rStyle As Long
   Dim r As Long

  'get the current ListView style
   rStyle = SendMessageLong(ListView1.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)

   If GridLineschk.Value = 0 Then
     'remove the extended bit
      rStyle = rStyle Xor LVS_EX_GRIDLINES

   ElseIf GridLineschk.Value = 1 Then
     'set the extended bit
      rStyle = rStyle Or LVS_EX_GRIDLINES

   End If

  'set the new ListView style
   r = SendMessageLong(ListView1.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)

End Sub

Private Sub Lblinfo_Change()
PicFolder.Left = Lblinfo.Left + (Lblinfo.Width + 300)
lblOpenFolder.Left = PicFolder.Left + (PicFolder.Width + 50)
FrameFileInfo.Width = lblOpenFolder.Left + (lblOpenFolder.Width + 300)

FrameView.Top = Me.Height - 1850
FrameView.Height = FrameTraffic.Height
FrameView.Left = FrameFileInfo.Left + (FrameFileInfo.Width + 100)

End Sub

Private Sub lblOpenFolder_Click()
StartNewBrowser (GetFilePath(TVPath, True))
End Sub

Private Sub lblOpenFolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOpenFolder.FontBold = True
End Sub

Private Sub LblVersion_Change()
LblSize.Left = LblVersion.Left + (LblVersion.Width + 300)
End Sub


Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim FP As FILE_PARAMS
Dim CurFile As Long

TVHost = Connection(ListView1.ListItems(Item.index).Tag).RemoteHost
TVPath = Connection(ListView1.ListItems(Item.index).Tag).FileName
TVTAG = ListView1.ListItems(Item.index).Tag
TVPI = Connection(ListView1.ListItems(Item.index).Tag).ProcessID

GetLargeIcon (TVPath)

   With FP
      .sFileNameExt = TVPath
   End With
   
CurFile = GetFileInformation(FP)

DoEvents
'If ResolveHostchk.Value = 0 Then lblHost.Caption = "Remote Host : " & GetHostNameFromIP(GetIPAddress(TVHost)) Else lblHost.Caption = "Remote Host : " & GetIPAddress(TVHost)

'PopulateTreeview (Item.Index)
'item click
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If ListView1.SelectedItem.Selected = False Then Exit Sub


CloseProgram.Caption = "Close Program : " & GetFileNameFromPath(TVPath)
CloseConnection.Caption = "Close Connection by : " & GetFileNameFromPath(TVPath)

If TVPath = "" Then CloseProgram.Enabled = False Else CloseProgram.Enabled = True

PopupMenu PopMenu

End If

End Sub

Private Sub ResolveHostchk_Click()
If Processing = False Then RefreshList
If ResolveHostchk.Value = 0 Then ListView1.ColumnHeaders(3).Width = 3500 Else ListView1.ColumnHeaders(3).Width = 1100
End Sub

Private Sub Timer1_Timer()
NotOnline (IsNetConnectOnline())
End Sub

Public Sub NotOnline(Online As Boolean)

If Online = False Then
    If IsOnline = False Then Exit Sub

    ListView1.ListItems.Clear
    ListView1.Enabled = False
    
    Timer2.Enabled = False
    
    StatusBar.Panels(4).Text = "Offline"
    
    IsOnline = False
    
    nid.szTip = "Connection Monitor(OffLine)"
    Shell_NotifyIcon NIM_MODIFY, nid

    Exit Sub
End If

If Online = True Then
    If IsOnline = True Then GoTo CallRefresh
    
    ListView1.Enabled = True
    Timer2.Enabled = True
    
    StatusBar.Panels(4).Text = "Online"
    
    IsOnline = True
    
    nid.szTip = "Connection Monitor (Online)"
    Shell_NotifyIcon NIM_MODIFY, nid
    
End If


CallRefresh:

If GetRefresh = True Then RefreshList
End Sub
Private Sub UpdateInterfaceInfo()
Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Long
Static lngBytesSent     As Long
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean
If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
Set objInterface = m_objIpHelper.Interfaces(1)
Select Case objInterface.InterfaceType
'Case MIB_IF_TYPE_ETHERNET: lblType.Caption = "Ethernet"
'Case MIB_IF_TYPE_FDDI: lblType.Caption = "FDDI"
'Case MIB_IF_TYPE_LOOPBACK: lblType.Caption = "Loopback"
'Case MIB_IF_TYPE_OTHER: lblType.Caption = "Other"
'Case MIB_IF_TYPE_PPP: lblType.Caption = "PPP"
'Case MIB_IF_TYPE_SLIP: lblType.Caption = "SLIP"
'Case MIB_IF_TYPE_TOKENRING: lblType.Caption = "TokenRing"
End Select

If ShowTrafficInBytes = False Then
    lblRecv.Caption = GiveByteValues(Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")))
    lblSent.Caption = GiveByteValues(Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###")))
Else
    lblRecv.Caption = Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###"))
    lblSent.Caption = Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###"))
End If
  '
    blnIsRecv = (m_objIpHelper.BytesReceived > lngBytesRecv)
    blnIsSent = (m_objIpHelper.BytesSent > lngBytesSent)
    '
    If blnIsRecv And blnIsSent Then
        Set Image1.Picture = ImageList2.ListImages(4).Picture
    ElseIf (Not blnIsRecv) And blnIsSent Then
        Set Image1.Picture = ImageList2.ListImages(3).Picture
    ElseIf blnIsRecv And (Not blnIsSent) Then
        Set Image1.Picture = ImageList2.ListImages(2).Picture
    ElseIf Not (blnIsRecv And blnIsSent) Then
        Set Image1.Picture = ImageList2.ListImages(1).Picture
    End If
    '
    lngBytesRecv = m_objIpHelper.BytesReceived
    lngBytesSent = m_objIpHelper.BytesSent
    '

    Set st_objInterface = objInterface

End Sub

Private Sub Timer2_Timer()
Call UpdateInterfaceInfo
End Sub

Private Sub Timer3_Timer()
'##############################################
Set objInterface2 = New CInterface
Set obJHelper = m_objIpHelper.Interfaces(1)

oNew = Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")) '//give bandwidth
aNew = Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###")) '//give bandwidth

tValue = oNew - oOld
aValue = aNew - aOld

StatusBar.Panels(1).Text = "Incomming: " & Trim(Format((tValue / 1000), "####0.00")) & " kbs"
StatusBar.Panels(2).Text = "Outgoing: " & Trim(Format((aValue / 1000), "####0.00")) & " kbs"

oOld = oNew
aOld = aNew
End Sub

Private Function GetFileNameFromPath(ByVal sFullPath As String) As String

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
         
         If sFullPath = "" Then
         GetFileNameFromPath = "Unknown"
         Exit Function
         End If
         
   hFile = FindFirstFile(sFullPath, WFD)
   
   If hFile <> INVALID_HANDLE_VALUE Then
   
     'the filename portion is in cFileName
      GetFileNameFromPath = TrimNull(WFD.cFileName)
      Call FindClose(hFile)
      
   End If
   
End Function


Private Function TrimNull(startstr As String) As String

   Dim pos As Integer
   
   pos = InStr(startstr, Chr$(0))
   
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
  
   TrimNull = startstr
  
End Function

Public Function PingIP(IP As String)
Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long
   Dim sIPAddress As String
   
   If SocketsInitialize() Then
   
     'convert the host name into an IP address
      sIPAddress = IP
      
     'ping the ip passing the address, text
     'to use, and the ECHO structure
      success = Ping(sIPAddress, "Echo This", ECHO)
      
     'display the results
      If GetStatusCode(success) = "ip success" Then PingIP = "Success - Round Time : " & ECHO.RoundTripTime & " ms" Else PingIP = GetStatusCode(success)
      'ECHO.Address
      'ECHO.RoundTripTime & " ms"
      'ECHO.DataSize & " bytes"
      
      If Left$(ECHO.Data, 1) <> Chr$(0) Then
         pos = InStr(ECHO.Data, Chr$(0))
         'Left$(ECHO.Data, pos - 1)
      End If
   
      'ECHO.DataPointer
      
      SocketsCleanup
      
   Else
   
        MsgBox "Windows Sockets for 32 bit Windows " & _
               "is not successfully responding.", vbInformation, "Error"
   
   End If

End Function

Private Function GetFileInformation(FP As FILE_PARAMS) As Long

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim nSize As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   Dim itmx As ListItem
   Dim lv As Control
       

  'FP.sFileRoot (assigned to sRoot) contains
  'the path to search.
  '
  'FP.sFileNameExt (assigned to sPath) contains
  'the full path and filespec.
   sPath = FP.sFileNameExt
   
   FrameFileInfo.Caption = "File Information (" & GetFileNameFromPath(sPath) & ")"

  'obtain handle to the first filespec match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then

      
        'remove trailing nulls
         sTmp = TrimNull(WFD.cFileName)
         
        'Even though this routine uses filespecs,
        '*.* is still valid and will cause the search
        'to return folders as well as files, so a
        'check against folders is still required.
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) _
            = FILE_ATTRIBUTE_DIRECTORY Then
      
            
           'retrieve the size and assign to nSize to
           'be returned at the end of this function call
            nSize = nSize + (WFD.nFileSizeHigh * (MAXDWORD + 1)) + WFD.nFileSizeLow
            
           'add to the list if the flag indicates
                             
              'got the data, so add it to the listview
               'Set itmx = lv.ListItems.Add(, , LCase$(sTmp))
               
               'itmx.SubItems(1) = GetFileVersion(sRoot & sTmp)
               'itmx.SubItems(3) = GetFileSizeStr(WFD.nFileSizeHigh + WFD.nFileSizeLow)
               'itmx.SubItems(2) = GetFileDescription(sRoot & sTmp)
               'itmx.SubItems(4) = LCase$(sRoot)
               
               lblOpenFolder.Enabled = True
               
                If GetFileDescription(sPath) = "" Then Lblinfo.Caption = "Description : (No Description) " Else Lblinfo.Caption = "Description : " & GetFileDescription(sPath)
                If GetFileVersion(sPath) = "" Then LblVersion.Caption = "Version : (No Version) " Else LblVersion.Caption = "Version : " & GetFileVersion(sPath)
                LblSize.Caption = "Size : " & GetFileSizeStr(WFD.nFileSizeHigh + WFD.nFileSizeLow)

         Else
         
         lblOpenFolder.Enabled = False
         Lblinfo.Caption = "Description : (Unknown)"
         LblVersion.Caption = "Version : (Unknown)"
         LblSize.Caption = "Size : (Unknown)"
         End If
         
      
     'close the handle
      hFile = FindClose(hFile)
   
            Else
         
         lblOpenFolder.Enabled = False
         Lblinfo.Caption = "Description : (Unknown)"
         LblVersion.Caption = "Version : (Unknown)"
         LblSize.Caption = "Size : (Unknown)"
         
   End If
   
   GetFileInformation = nSize
   
End Function


Private Function GetFileSizeStr(fsize As Long) As String

    GetFileSizeStr = GiveByteValues(Format$((fsize), "###,###,###"))  '& " kb"
  
End Function

Private Function StartNewBrowser(sURL As String) As Boolean
       
  'start a new instance of the user's browser
  'at the page passed as sURL
   Dim success As Long
   Dim hProcess As Long
   Dim sBrowser As String
   Dim start As STARTUPINFO
   Dim proc As PROCESS_INFORMATION
   
   sBrowser = GetBrowserName(success)
   
  'did sBrowser get correctly filled?
   If success >= ERROR_FILE_SUCCESS Then
      
     'prepare STARTUPINFO members
      With start
         .cb = Len(start)
         .dwFlags = STARTF_USESHOWWINDOW
         .wShowWindow = SW_SHOWNORMAL
      End With
      
     'start a new instance of the default
     'browser at the specified URL. The
     'lpCommandLine member (second parameter)
     'requires a leading space or the call
     'will fail to open the specified page.
      success = CreateProcess(sBrowser, _
                              " " & sURL, _
                              0&, 0&, 0&, _
                              NORMAL_PRIORITY_CLASS, _
                              0&, 0&, start, proc)
                                  
     'if the process handle is valid, return success
      StartNewBrowser = proc.hProcess <> 0
     
     'don't need the process
     'handle anymore, so close it
      Call CloseHandle(proc.hProcess)

     'and close the handle to the thread created
      Call CloseHandle(proc.hThread)

   End If

End Function


Private Function GetBrowserName(dwFlagReturned As Long) As String

  'find the full path and name of the user's
  'associated browser
   Dim hFile As Long
   Dim sResult As String
   Dim sTempFolder As String
        
  'get the user's temp folder
   sTempFolder = GetTempDir()
   
  'create a dummy html file in the temp dir
   hFile = FreeFile
      Open sTempFolder & "dummy.html" For Output As #hFile
   Close #hFile

  'get the file path & name associated with the file
   sResult = Space$(MAX_PATH)
   dwFlagReturned = FindExecutable("dummy.html", sTempFolder, sResult)
  
  'clean up
   Kill sTempFolder & "dummy.html"
   
  'return result
   GetBrowserName = TrimNull(sResult)
   
End Function


Public Function GetTempDir() As String

  'retrieve the user's system temp folder
   Dim tmp As String
   
   tmp = Space$(256)
   Call GetTempPath(Len(tmp), tmp)
   
   GetTempDir = TrimNull(tmp)
    
End Function
'--end block--'

Public Function BasePath(ByVal fname As String, Optional delim As String = "\", Optional keeplast As Boolean = True) As String
    Dim outstr As String
    Dim llen As Long
    llen = InStrRev(fname, delim)


    If (Not keeplast) Then
        llen = llen - 1
    End If


    If (llen > 0) Then
        BasePath = Mid(fname, 1, llen)
    Else
        BasePath = fname
    End If
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

