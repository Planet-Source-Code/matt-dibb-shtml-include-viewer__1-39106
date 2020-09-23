VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "mattSHTML"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "shtml_viewer.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4800
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox t 
      Height          =   1695
      Left            =   4560
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2990
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"shtml_viewer.frx":08DA
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      ExtentX         =   6588
      ExtentY         =   4895
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuLineexit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDunctions 
      Caption         =   "&Fucntions"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private LoadComplete As Boolean
Private FilePath As String
Private CurrentPage As String
Private Sub Form_Resize()
On Error GoTo e
web.Width = ScaleWidth
web.Height = ScaleHeight
e:
End Sub


Private Sub mnuAbout_Click()
MsgBox "mattSHTML" & Chr(10) & Chr(10) & "Previews complete local SHTML pages without a server" & Chr(10) & Chr(10) & "Matt Dibb 2002", vbInformation, "About"
End Sub

Private Sub mnuExit_Click()
End
End Sub


Private Sub mnuOpen_Click()
Dim OriginalText As String
Dim IncludeFile As String
Dim InsertText As String

cd.ShowOpen

LoadComplete = False
If cd.FileName = "" Then Exit Sub

t.LoadFile cd.FileName
CurrentPage = cd.FileName
e = InStrRev(cd.FileName, "\")
FilePath = Left(cd.FileName, e)

OriginalText = t.Text
'<!-- #include file="include.html" -->

For i = 1 To Len(OriginalText)
    a1 = InStr(i, OriginalText, "<!-- #include")
    If a1 = "0" Then GoTo n
    a2 = InStr(a1 + 15, OriginalText, """")
    a3 = InStr(a2 + 1, OriginalText, """")
    IncludeFile = Mid(OriginalText, a2 + 1, (a3 - a2) - 1)
    i = a3
    t.LoadFile IncludeFile
    InsertText = t.Text
    t.Text = OriginalText
    t.SelStart = a1 - 1
    t.SelLength = (a3 - a1) + 5
    t.SelText = InsertText
    OriginalText = t.Text
Next

n:
t.Text = OriginalText
t.SaveFile FilePath & "shtml_temp.html", rtfText
web.Navigate FilePath & "shtml_temp.html"
LoadComplete = True
End Sub


Private Sub mnuRefresh_Click()
web.Navigate CurrentPage
End Sub

Private Sub web_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error GoTo e
If LoadComplete = False Then Exit Sub
LoadComplete = False
MousePointer = 11
'e = InStrRev(URL, "\")
'newurl = Right(URL, Len(URL) - e) ' Left(URL, e)
't.LoadFile FilePath & newurl
t.LoadFile URL, rtfText

K = InStr(1, URL, "shtml_temp.html")
If K <> "0" Then GoTo k2
CurrentPage = URL

k2:
Dim OriginalText As String
Dim IncludeFile As String
Dim InsertText As String


OriginalText = t.Text
'<!--#include file="top.html" -->

For i = 1 To Len(OriginalText)
a1 = InStr(i, OriginalText, "<!--#include")
If a1 = "0" Then GoTo n
a2 = InStr(a1 + 15, OriginalText, """")
a3 = InStr(a2 + 1, OriginalText, """")
IncludeFile = Mid(OriginalText, a2 + 1, (a3 - a2) - 1)

i = a3
t.LoadFile IncludeFile
InsertText = t.Text
t.Text = OriginalText
t.SelStart = a1 - 1
t.SelLength = (a3 - a1) + 5
t.SelText = InsertText
OriginalText = t.Text
Next

n:
t.Text = OriginalText
t.SaveFile FilePath & "shtml_temp.html", rtfText
web.Stop
web.Navigate FilePath & "shtml_temp.html"
e:
LoadComplete = True
MousePointer = 0
End Sub

