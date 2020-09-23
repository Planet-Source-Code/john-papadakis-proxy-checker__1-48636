VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proxy Checker By ||Ni0||"
   ClientHeight    =   9660
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   11835
   FillStyle       =   6  'Cross
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Tools 
      BackColor       =   &H80000007&
      Caption         =   "Tools"
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   11535
      Begin MSComDlg.CommonDialog dialog 
         Left            =   4080
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "txt"
         Orientation     =   2
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   9960
         TabIndex        =   3
         Text            =   "4"
         Top             =   480
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3480
         Top             =   360
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   5520
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label command2 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "Pause"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Command1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "Check"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "Connection Time Out Sec:"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   8040
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11535
      Begin VB.Frame Frame3 
         BackColor       =   &H80000007&
         Caption         =   "List"
         ForeColor       =   &H0000FF00&
         Height          =   7935
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3855
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            ForeColor       =   &H0000FF00&
            Height          =   7635
            Left            =   2760
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            ForeColor       =   &H0000FF00&
            Height          =   7635
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000006&
         Caption         =   "Tested"
         ForeColor       =   &H0000FF00&
         Height          =   7935
         Left            =   3720
         TabIndex        =   4
         Top             =   0
         Width           =   7815
         Begin MSComctlLib.ListView ListView1 
            Height          =   7695
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   13573
            View            =   3
            Arrange         =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   65535
            BackColor       =   4210752
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   10200
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Listen 
      Left            =   10080
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar bar 
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   9240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Menu lis 
      Caption         =   "List"
      Begin VB.Menu insert 
         Caption         =   "Insert"
      End
      Begin VB.Menu clear_un 
         Caption         =   "Clear Unchecked"
      End
      Begin VB.Menu clear_ch 
         Caption         =   "Clear Checked"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Copy 
         Caption         =   "Copy"
         Begin VB.Menu IPport 
            Caption         =   "IP:PORT"
         End
         Begin VB.Menu IPp 
            Caption         =   "IP"
         End
      End
      Begin VB.Menu check 
         Caption         =   "Check"
      End
      Begin VB.Menu uncheck 
         Caption         =   "UnCheck"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IP, I, p, File_Name
Dim proxok As Boolean
Dim httP As Boolean
Public arR, striN, outT
Dim chk As Boolean
Dim lItem, opp As ListItem
Dim nett As Boolean
Private Sub check_Click()
For c = 1 To ListView1.ListItems.Count
 If ListView1.ListItems.Item(c).Selected = True Then
 ListView1.ListItems.Item(c).Checked = True
 End If
 Next
End Sub
Private Sub Command1_Click()
On Error Resume Next
If p <> 0 Then
ListView1.ListItems.Remove (p + 1)
End If
Text1.Enabled = False
Command1.Enabled = False
command2.Enabled = True
If p = 0 Then
List1.Selected(0) = True
List2.Selected(0) = True
End If
For p = List1.ListIndex To List1.ListCount
If List1.ListIndex = List1.ListCount - 1 Then
command2.Enabled = False
p = 0
I = 0
Command1.Enabled = True
Exit Sub
End If
List1.Selected(p) = True
List2.Selected(p) = True
proxok = False
httP = False
chk = False
outT = 0
Winsock1.Close
IP = List1.Text
If CheckConnection = False Then
If InStr(1, IP, "127.0.0") = 0 Then
Call MsgBox("Not Connected to Any Net!!!", vbCritical, "Error"): Command1.Enabled = True: command2.Enabled = False: Exit Sub
End If
End If
Winsock1.Connect IP, List2.Text
Set lItem = ListView1.ListItems.Add(, , IP)
lItem.ListSubItems.Add , , List2.Text
Timer1.Enabled = True
Do While Winsock1.State <> 7
DoEvents
If outT = Text1.Text Then outT = 0:  lItem.ListSubItems.Add , , "Conn Failed":  Timer1.Enabled = False: GoTo conti
Loop
Timer1.Enabled = False
Winsock1.SendData "CONNECT " & Winsock1.LocalIP & ":" & "5555 " & "HTTP/1.0" & vbCrLf & vbCrLf
Timer1.Enabled = True
Do While chk = False
DoEvents
If outT = Text1 Then outT = 0: Timer1.Enabled = False:  Exit Do
Loop
If httP = True Then lItem.ListSubItems.Add , , "HTTP Annon": lItem.ListSubItems.Add , , ResolveHostName(IP): GoTo conti
If proxok = True Then lItem.ListSubItems.Add , , "HTTP,elite SSL":  lItem.ListSubItems.Add , , ResolveHostName(IP): GoTo conti
If httP = False Then
If proxok = False Then
lItem.ListSubItems.Add , , "Invalid!"
 lItem.ListSubItems.Add , , ResolveHostName(IP)
End If
End If
conti:
Timer1.Enabled = True
bar.Value = bar.Value + 1
Next
End Sub
Private Sub Command2_Click()
Command1.Enabled = True
command2.Enabled = False
Text1.Enabled = True
Do Until Command1.Enabled = False
DoEvents
Loop
End Sub
Private Sub Copy_Click()
Clipboard.SetText List1.Text
End Sub

Private Sub Form_Load()
On Error Resume Next
ListView1.Checkboxes = True
Command1.Enabled = False
I = 0
p = 0
command2.Enabled = False
ListView1.ColumnHeaders.Add , , "IP", ListView1.Width / 4 - 20
ListView1.ColumnHeaders.Add , , "Port", ListView1.Width / 4 - 800
ListView1.ColumnHeaders.Add , , "Status", ListView1.Width / 4 - 200
ListView1.ColumnHeaders.Add , , "Host", ListView1.Width / 4 + 963
outT = 0
Listen.Close
Listen.LocalPort = 5555
Listen.Listen
End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Width = Form1.Width - 400
Frame2.Width = Frame1.Width - Frame3.Width + 95
ListView1.Width = Frame2.Width
 ListView1.ColumnHeaders.Item(1).Width = ListView1.Width / 4 - 20
 ListView1.ColumnHeaders.Item(2).Width = ListView1.Width / 4 - 800
 ListView1.ColumnHeaders.Item(3).Width = ListView1.Width / 4 - 200
 ListView1.ColumnHeaders.Item(4).Width = ListView1.Width / 4 + 960
 bar.Width = Form1.Width - 900
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub List3_Click()
List4.Selected(List3.ListIndex) = True
List5.Selected(List3.ListIndex) = True
End Sub

Private Sub insert_Click()
On Error Resume Next
List1.Clear
List2.Clear
ListView1.ListItems.Clear
dialog.ShowOpen
File_Name = dialog.FileName
Open File_Name For Binary Access Read As #1
Dim lista As String
si = FileLen(File_Name)
lista = Space$(si)
Get #1, , lista
arR = Split(lista, vbCrLf)
For I = 0 To UBound(arR)
arR = Split(lista, vbCrLf)
striN = Split(arR(I), ":")
Port = striN(1)
 List2.AddItem striN(1)
 List1.AddItem striN(0)
 Next
 bar.Max = UBound(arR)
 Close #1
 Command1.Enabled = True
 Label2 = UBound(arR) & "  Proxys Added"
End Sub

Private Sub IPp_Click()
Dim cop As String
For c = 1 To ListView1.ListItems.Count
 If ListView1.ListItems.Item(c).Checked = True Then
 cop = cop & ListView1.ListItems.Item(c).Text & vbCrLf
 End If
Next
 Clipboard.Clear
  Clipboard.SetText cop
End Sub

Private Sub IPport_Click()
Dim sum, cop, ii As String
u = InputBox("The Char between IP and port,Usually : or Space")
For c = 1 To ListView1.ListItems.Count
 If ListView1.ListItems.Item(c).Checked = True Then
sum = ListView1.ListItems.Item(c).Text
List1.Selected(c - 1) = True
List2.Selected(c - 1) = True
ii = List2.Text
cop = cop & sum & u & ii & vbCrLf
End If
Next
Clipboard.Clear
 Clipboard.SetText cop
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
List2.Selected(List1.ListIndex) = True
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 1 Then List2.Selected(List1.ListIndex) = True
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
List1.Selected(List2.ListIndex) = True
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 1 Then List1.Selected(List2.ListIndex) = True
End Sub

Private Sub Listen_ConnectionRequest(ByVal requestID As Long)
Listen.Close
Listen.Accept requestID
End Sub

Private Sub net_Timer()
If CheckConnection = False Then Call MsgBox("Disconnected From Net!!!", vbCritical, "Error")
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu Edit
End If
End Sub

Private Sub Timer1_Timer()
outT = outT + 1
End Sub

Private Sub uncheck_Click()
For c = 1 To ListView1.ListItems.Count
 
 If ListView1.ListItems.Item(c).Selected = True Then
 ListView1.ListItems.Item(c).Checked = False
 End If
 Next
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim a As String
Winsock1.GetData a
chk = True
If InStr(1, a, "Connection established") <> 0 Then proxok = True: httP = False
If InStr(1, a, "HTTP/1.") <> 0 Then
If InStr(1, a, "40") <> 0 Then httP = True: proxok = False
End If
End Sub

