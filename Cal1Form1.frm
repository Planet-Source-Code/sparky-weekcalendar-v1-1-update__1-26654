VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00876034&
   BorderStyle     =   0  'None
   Caption         =   "WeekCalendar"
   ClientHeight    =   7980
   ClientLeft      =   5175
   ClientTop       =   225
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   6810
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      Caption         =   "Ä"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   360
      Width           =   210
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   225
      Left            =   60
      Picture         =   "Cal1Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   75
      Width           =   210
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   7830
      Left            =   3510
      ScaleHeight     =   7800
      ScaleWidth      =   3180
      TabIndex        =   2
      Top             =   75
      Width           =   3210
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   780
         Index           =   6
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   6780
         Width           =   3000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   780
         Index           =   5
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   5700
         Width           =   3000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   780
         Index           =   4
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   4620
         Width           =   3000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   780
         Index           =   3
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   3540
         Width           =   3000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   780
         Index           =   2
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2460
         Width           =   3000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   780
         Index           =   1
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1380
         Width           =   3000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   780
         Index           =   0
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   300
         Width           =   3000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0090D0D8&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   90
         TabIndex        =   16
         Top             =   6570
         Width           =   3000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0090D0D8&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   75
         TabIndex        =   15
         Top             =   5475
         Width           =   3000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0090D0D8&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   14
         Top             =   4395
         Width           =   3000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0090D0D8&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   13
         Top             =   3315
         Width           =   3000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0090D0D8&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   12
         Top             =   2235
         Width           =   3000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0090D0D8&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   1155
         Width           =   3000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0090D0D8&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   75
         Width           =   3000
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7935
      Index           =   0
      Left            =   0
      Picture         =   "Cal1Form1.frx":014A
      ScaleHeight     =   7935
      ScaleWidth      =   360
      TabIndex        =   1
      Top             =   -15
      Width           =   360
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   7440
      Left            =   420
      TabIndex        =   0
      Top             =   255
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   13123
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthRows       =   3
      MultiSelect     =   -1  'True
      ShowToday       =   0   'False
      StartOfWeek     =   24444929
      TitleBackColor  =   8421376
      TitleForeColor  =   14219772
      CurrentDate     =   37125
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7935
      Index           =   1
      Left            =   0
      Picture         =   "Cal1Form1.frx":2503
      ScaleHeight     =   7935
      ScaleWidth      =   360
      TabIndex        =   17
      Top             =   15
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sDay As String
Dim eDay As String
Dim SelDay As Integer
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim SvDate As String
Dim isDirty As Boolean

Private Sub Command1_Click()
SaveWeeksData

Unload Me
End Sub

Private Sub Command2_Click()
Text1(0).SetFocus
Me.WindowState = 1


End Sub

Private Sub Form_Load()

Me.Left = 8300 '5190 '8300
Me.Top = 280
Me.Width = 3705
Picture2.Left = 390
Me.Show
DoEvents
isDirty = False
OpenDataBase

SelectWeek CStr(Now)
ClearText
GetWeeksData
Text1(0).SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
If isDirty = True Then SaveWeeksData
cnn.Close
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
' update database from text1()
SaveWeeksData
isDirty = False
'select the week
SelectWeek CStr(DateClicked)
' load data from database
ClearText
GetWeeksData

End Sub

Private Sub SelectWeek(sDate As String)
Dim x As Integer

MonthView1.Value = sDate
SelDay = MonthView1.DayOfWeek

If SelDay = 1 Then
  eDay = DateAdd("y", 6, MonthView1.Value)
  MonthView1.SelStart = MonthView1.Value
  MonthView1.SelEnd = eDay
  For x = 0 To 6
    Label1(x).Caption = Format(MonthView1.Value + x, "long date")
  Next x
Else
  sDay = MonthView1.Value - (MonthView1.DayOfWeek - 1)
  MonthView1.Value = sDay
  eDay = DateAdd("y", 6, MonthView1.Value)
  MonthView1.SelStart = MonthView1.Value
  MonthView1.SelEnd = eDay
  For x = 0 To 6
    Label1(x).Caption = Format(MonthView1.Value + x, "long date")
  Next x
End If

' were switching to a new date so save the previous date
SvDate = MonthView1.Value

End Sub

Private Sub OpenDataBase()
   
   Set cnn = New ADODB.Connection
   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\cal1proj1.mdb;"

End Sub

Sub GetWeeksData()
Dim gWeek As String
Dim wDataMetrics As String
gWeek = MonthView1.SelStart

 Set rst = New ADODB.Recordset
 rst.Open "WeekData", cnn, adOpenKeyset, adLockOptimistic
 rst.Find "SWeek='" & gWeek & "'", 0, adSearchForward
 
 If rst.EOF = True Or rst.BOF = True Then Exit Sub
  
 wDataMetrics = rst.Fields("WeekData").Value
  
 rst.Close
 
 ParseData wDataMetrics
 
End Sub

Private Sub ParseData(wDataMetrics As String)
Dim wData As Variant
Dim x As Integer

  wData = Split(wDataMetrics, "*")
 For x = 0 To UBound(wData) - 1
   If InStr(wData(x), "~") Then
      Text1(x) = ""
   Else
      Text1(x) = wData(x)
   End If
 Next x
  
End Sub

Private Sub ClearText()
Dim x As Integer
  For x = 0 To 6
   Text1(x) = ""
  Next x
End Sub

Private Sub SaveWeeksData()
Dim x As Integer
Dim sWeek As String
Dim sData As String
  
  For x = 0 To 6
   If Len(Text1(x)) > 0 Then
        sData = sData & Text1(x) & "*"
   Else
        sData = sData & "~*" ' empty box
   End If
   
  Next x

If Len(sData) = 0 Then Exit Sub

sWeek = Format(SvDate, "m/d/yy")

Set rst = New ADODB.Recordset
' Open the recordset
rst.Open "SELECT * FROM WeekData WHERE SWeek = '" & sWeek & "'", _
      cnn, adOpenKeyset, adLockOptimistic
' Add a new record
If rst.EOF = True Or rst.BOF = True Then  'new record
  rst.AddNew
     rst!sWeek = sWeek
     rst!WeekData = sData
     rst.Update
Else ' update record
    rst.Fields("sWeek").Value = sWeek
    rst.Fields("WeekData").Value = sData
    rst.Update
End If

rst.Close
isDirty = False
End Sub


Private Sub MonthView1_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
Dim M As String
Dim D As String
Dim Y As String

M = MonthView1.Month
D = MonthView1.Day
Y = MonthView1.Year
SaveWeeksData
SelectWeek CStr(StartDate)
ClearText
GetWeeksData
Text1(0).SetFocus

End Sub

Private Sub Picture1_Click(Index As Integer)
Dim fLeft As Long
Dim pLeft As Long
Dim fWidth As Long

Select Case Index
   Case 0
        Me.Move 5190, 280, 6810, 7980
        Picture2.Move 3510, 75, 3210, 7830
        Picture1(0).ZOrder 1
   Case 1
        Me.Move 8300, 280, 3705, 7980
        Picture2.Move 390, 75, 3210, 7830
        Picture1(1).ZOrder 1
End Select
Text1(0).SetFocus
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'126 42
isDirty = True
If KeyAscii = 42 Or KeyAscii = 126 Then
  KeyAscii = 0
End If

End Sub
