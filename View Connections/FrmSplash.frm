VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSplash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   1635
   ClientLeft      =   5865
   ClientTop       =   4635
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4695
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3720
      Top             =   120
   End
   Begin ComctlLib.ProgressBar Progbar 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   10
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   393216
      FullWidth       =   41
      FullHeight      =   41
   End
   Begin VB.Label LblLoad 
      Caption         =   "Loading..."
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CreateManifest()
Dim f, fso As New FileSystemObject
    
fso.CreateTextFile GetAppPath & App.EXEName & ".exe" & ".manifest", False
Set f = fso.OpenTextFile(GetAppPath & App.EXEName & ".exe" & ".manifest", ForAppending, TristateFalse)
    f.Write LoadResString("101")
    f.Close
DoEvents
MsgBox "No manifest file found please restart the application", vbInformation + vbOKOnly, "Load ManifestFile"
End
End Sub

Private Sub Form_Load()
Dim TempData
Dim fso As New FileSystemObject

Progbar.Value = 0

If fso.FileExists(GetAppPath & "102.avi") = False Then

    TempData = BuildFileFromResource(GetAppPath & "102.avi", 102, "AVI")

        If TempData <> "" Then

            If fso.FileExists(TempData) = True Then
            Animation1.Open TempData
            Animation1.Play
            End If

        End If

Else
Animation1.Open GetAppPath & "102.avi"
Animation1.Play

End If

DoEvents
End Sub

Private Sub Timer1_Timer()
LoadUP
End Sub

Private Sub LoadUP()
Dim fso As New FileSystemObject

Progbar.Value = 1

LblLoad.Caption = "Loading... Manifest File"

If fso.FileExists(GetAppPath & App.EXEName & ".exe" & ".manifest") = False Then
LblLoad.Caption = "Loading... Creating Manifest File"
CreateManifest ' fso.CreateTextFile GetAppPath & App.EXEName & ".exe" & ".manifest", False
DoEvents
End If

Progbar.Value = 3

LblLoad.Caption = "Loading... GUI"

Load FrmMain
DoEvents

Progbar.Value = 7

LblLoad.Caption = "Loading... Checking Connections"

LblLoad.Caption = IIf(IsNetConnectOnline, "Online", "OffLine")

Progbar.Value = 9

FrmMain.Show
DoEvents

Unload Me
End Sub
