VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Türkçe Sözlük"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7860
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   5355
      TabIndex        =   36
      Top             =   1755
      Width           =   2400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Arama"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   135
      TabIndex        =   26
      Top             =   135
      Width           =   5130
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   180
         TabIndex        =   32
         Top             =   315
         Width           =   1050
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bul"
         Height          =   285
         Left            =   3825
         TabIndex        =   31
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   180
         TabIndex        =   30
         Top             =   675
         Width           =   1050
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   810
         TabIndex        =   29
         Top             =   1035
         Width           =   1050
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Bul"
         Height          =   285
         Left            =   3825
         TabIndex        =   28
         Top             =   675
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Bul"
         Height          =   285
         Left            =   3840
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ile baþlayan sözcükleri bul."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1350
         TabIndex        =   35
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "ile biten sözcükleri bul."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1305
         TabIndex        =   34
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Ýçinde                     geçen sözcükleri bul. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   33
         Top             =   1080
         Width           =   4575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Düzenle"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   135
      TabIndex        =   22
      Top             =   1710
      Width           =   5175
      Begin VB.CommandButton Command4 
         Caption         =   "Düzenle"
         Height          =   285
         Left            =   3690
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1710
         TabIndex        =   23
         Top             =   225
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Karýþýk kelime"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   270
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulmaca"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   135
      TabIndex        =   1
      Top             =   2565
      Width           =   5130
      Begin VB.CommandButton Command5 
         Caption         =   "Oluþtur"
         Height          =   285
         Left            =   2070
         TabIndex        =   20
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1215
         TabIndex        =   19
         Top             =   270
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   1
         TabIndex        =   18
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   1
         Left            =   360
         MaxLength       =   1
         TabIndex        =   17
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   2
         Left            =   600
         MaxLength       =   1
         TabIndex        =   16
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   3
         Left            =   840
         MaxLength       =   1
         TabIndex        =   15
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   4
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   14
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   5
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   13
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   6
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   12
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   7
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   11
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   8
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   10
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   9
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   9
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   10
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   8
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   11
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   7
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   12
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   6
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   13
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   5
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   14
         Left            =   3480
         MaxLength       =   1
         TabIndex        =   4
         Top             =   360
         Width           =   210
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ara"
         Height          =   285
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Yeni sözcük"
         Height          =   255
         Left            =   3825
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Kaç harf"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   270
         Width           =   1935
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3930
      Left            =   90
      TabIndex        =   0
      Top             =   3690
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   6932
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6435
      Top             =   45
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":0442
      OLEDBString     =   $"Form1.frx":04E1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim kelimeler() As String
Dim anlamlar() As String
Dim crossl As Integer

Private Sub Command1_Click()
Adodc1.RecordSource = "select * from sozcukler where sozcuk like '" & Text1.Text & "%'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Columns(0).Width = 2400
DataGrid1.Columns(1).Width = 4500

End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select * from sozcukler where sozcuk like '%" & Text2.Text & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Columns(0).Width = 2400
DataGrid1.Columns(1).Width = 4500

End Sub


Private Sub Command3_Click()
Adodc1.RecordSource = "select * from sozcukler where sozcuk like '%" & Text3.Text & "%'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Columns(0).Width = 2400
DataGrid1.Columns(1).Width = 4500


End Sub

'Function Copyright stephen whittle
Private Sub Command4_Click()
Dim l As Integer, x As Integer, z As Integer, ins As Integer
Dim str1 As String
Dim found As Boolean

Set DataGrid1.DataSource = Nothing
List1.Clear
l = Len(Text4.Text)

Adodc1.RecordSource = "SELECT sozcuk From Sozcukler WHERE (Len(sozcuk)='" & l & "')"
Adodc1.Refresh

Adodc1.Recordset.MoveFirst
While Not Adodc1.Recordset.EOF

     str1 = Text4.Text
     found = True
   
       For x = 1 To l
         ins = InStr(1, str1, Mid$(Adodc1.Recordset.Fields(0).Value, x, 1))
        If ins = 0 Then
         found = False
         Exit For
        Else
         str1 = Left$(str1, ins - 1) & Mid$(str1, ins + 1)
        End If
       Next x
    
    If found = True Then
      List1.AddItem Adodc1.Recordset.Fields(0).Value
    End If
    
 Adodc1.Recordset.MoveNext
Wend

If List1.ListCount = 0 Then
 List1.AddItem "<< Sonuç Yok >>"
End If

End Sub

'Function Copyright stephen whittle
Private Sub Command5_Click()
Dim i As Integer, x As Integer


If IsNumeric(Text5.Text) Then
  x = Val(Text5.Text)
 If x < 16 Then
  Command5.Visible = False
  Text5.Visible = False
  Label5.Visible = False
  Command6.Visible = True
  Command7.Visible = True
  crossl = x
For i = 0 To x - 1
 Text6(i).Visible = True
Next i
 Else
  MsgBox "En fazla 15 harfli sözcük giriniz.", vbInformation, "Hata"
 End If
Else
 MsgBox "Sayý girmelisiniz.", vbInformation, "Kaç harf?"
End If

End Sub

'Function Copyright stephen whittle
Private Sub Command6_Click()
Dim i As Long, x As Integer, z As Integer, found As Boolean

Set DataGrid1.DataSource = Nothing
List1.Clear

Adodc1.RecordSource = "SELECT sozcuk From Sozcukler WHERE (Len(sozcuk)='" & crossl & "')"
Adodc1.Refresh

Adodc1.Recordset.MoveFirst
While Not Adodc1.Recordset.EOF
    
    found = True
     
     For x = 0 To crossl - 1
      If Text6(x).Text <> "" Then
        If Mid$(Adodc1.Recordset.Fields(0).Value, x + 1, 1) = Text6(x).Text Then
          found = True
        Else
          found = False
          Exit For
        End If
      End If
     Next x
  If found = True Then
    List1.AddItem Adodc1.Recordset.Fields(0).Value
  End If
 
 Adodc1.Recordset.MoveNext
Wend

If List1.ListCount = 1 Then
 List1.Clear
 List1.AddItem "<< Sonuç Yok >>"
End If

End Sub


Private Sub Command7_Click()
Dim i As Integer

Command6.Visible = False
Command5.Visible = True
Text5.Visible = True
Text5.Text = ""
Label5.Visible = True

For i = Text6.LBound To Text6.UBound
 Text6(i).Text = ""
 Text6(i).Visible = False
Next i

List1.Clear

Command7.Visible = False
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = Text6.LBound To Text6.UBound
 Text6(i).Visible = False
Next i

End Sub

