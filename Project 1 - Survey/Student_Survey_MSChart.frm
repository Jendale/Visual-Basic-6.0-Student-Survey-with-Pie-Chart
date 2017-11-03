VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Student_Survey_MSChart 
   Caption         =   "Student Survey Chart"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Legend"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   7200
      TabIndex        =   5
      Top             =   120
      Width           =   2055
      Begin VB.Label Label8 
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton btnCivilStatus 
      Caption         =   "Civil Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      ToolTipText     =   "View by civil status"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   840
      ScaleHeight     =   5715
      ScaleWidth      =   6195
      TabIndex        =   3
      Top             =   240
      Width           =   6255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   120
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
      Connect         =   ""
      OLEDBString     =   ""
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
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton btnCourse 
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "View by course"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton btnBloodtype 
      Caption         =   "Bloodtype"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      ToolTipText     =   "View by bloodtype"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton btnGender 
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   0
      ToolTipText     =   "View by gender"
      Top             =   6120
      Width           =   1335
   End
End
Attribute VB_Name = "Student_Survey_MSChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************** Global Variable ***********************************'
'***** Gender *****'
Public dMale As Double, dFemale As Double
'***** Bloodtype *****'
Public bA As Double, bB As Double, bO As Double, bRh As Double, bAB As Double, total As Double
'***** Course *****'
Public cBSHRM As Double, cBEED As Double, cBSIT As Double, cBSCS As Double
'***** Civil Status *****'
Public dSingle As Double, dMarried As Double

'***** Gender *****'
Public male As Integer, female As Integer
'***** Bloodtype *****'
Public A As Integer, B As Integer, AB As Integer, O As Integer
'***** Course *****'
Public BSHRM As Integer, BEED As Integer, BSIT As Integer, BSCS As Integer
'***** Civil Status *****'
Public iSingle As Integer, iMarried As Integer
'**********************************************************************************'
Option Explicit
Private Sub btnBloodtype_Click()
'***** Set Color and Pie Slice *****'
    Call DrawPiePiece(QBColor(1), 0.001, bA)
    Call DrawPiePiece(QBColor(2), bA, bA + bB)
    Call DrawPiePiece(QBColor(3), bA + bB, bA + bB + bO)
    Call DrawPiePiece(QBColor(4), bA + bB + bO, 99.999) 'bA + bB + bO + bAB
    'Call DrawPiePiece(QBColor(5), bA + bB + bO + bAB, 99.999)
    
'***** Set label Visibility *****'
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    
'***** Set label forecolor *****'
    Label1.ForeColor = QBColor(1)
    Label2.ForeColor = QBColor(2)
    Label3.ForeColor = QBColor(3)
    Label4.ForeColor = QBColor(4)
    
'***** Set label caption *****'
    Label1.Caption = "A"
    Label2.Caption = "B"
    Label3.Caption = "O"
    Label4.Caption = "AB"
    
'***** Set label result *****'
    Label5.Caption = A
    Label6.Caption = B
    Label7.Caption = O
    Label8.Caption = AB
    
    
End Sub
Private Sub btnCivilStatus_Click()
'***** Set Color and Pie Slice *****'
    Call DrawPiePiece(QBColor(1), 0.001, dSingle)
    Call DrawPiePiece(QBColor(4), dSingle, 99.999)

'***** Set label Visibility *****'
    Label1.Visible = True
    Label2.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    
'***** Set label forecolor *****'
    Label1.ForeColor = QBColor(1)
    Label2.ForeColor = QBColor(4)
       
'***** Set label caption *****'
    Label1.Caption = "Single"
    Label2.Caption = "Married"
    
'***** Set label result *****'
    Label5.Caption = iSingle
    Label6.Caption = iMarried
    
'***** Set label Visibility *****'
    Label3.Visible = False
    Label4.Visible = False
    Label7.Visible = False
    Label8.Visible = False
End Sub
Private Sub btnCourse_Click()
'***** Set Color and Pie Slice *****'
    Call DrawPiePiece(QBColor(1), 0.001, cBSHRM)
    Call DrawPiePiece(QBColor(2), cBSHRM, cBSHRM + cBEED)
    Call DrawPiePiece(QBColor(3), cBSHRM + cBEED, cBSHRM + cBEED + cBSIT)
    Call DrawPiePiece(QBColor(4), cBSHRM + cBEED + cBSIT, 99.999)

'***** Set label Visibility *****'
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    
'***** Set label forecolor *****'
    Label1.ForeColor = QBColor(1)
    Label2.ForeColor = QBColor(2)
    Label3.ForeColor = QBColor(3)
    Label4.ForeColor = QBColor(4)
    
'***** Set label caption *****'
    Label1.Caption = "BSHRM"
    Label2.Caption = "BEED"
    Label3.Caption = "BSIT"
    Label4.Caption = "BSCS"
    
'***** Set label result *****'
    Label5.Caption = BSHRM
    Label6.Caption = BEED
    Label7.Caption = BSIT
    Label8.Caption = BSCS
End Sub
Private Sub btnGender_Click()
'***** Set Color and Pie Slice *****'
    Call DrawPiePiece(QBColor(1), 0.001, dMale)
    Call DrawPiePiece(QBColor(4), dMale, 99.999)
    
'***** Set label Visibility *****'
    Label1.Visible = True
    Label2.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    
'***** Set label forecolor *****'
    Label1.ForeColor = QBColor(1)
    Label2.ForeColor = QBColor(4)
    
'***** Set label caption *****'
    Label1.Caption = "Male"
    Label2.Caption = "Female"
    
'***** Set label result *****'
    Label5.Caption = male
    Label6.Caption = female
    
'***** Set label Visibility *****'
    Label3.Visible = False
    Label4.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    
End Sub
Private Sub DrawPiePiece(lColor As Long, fStart As Double, fEnd As Double)
'***** Set Color and Pie Slice *****'
        Dim PI As Double
        Dim CircleEnd As Double
        PI = 3.14159265359
        CircleEnd = -2 * PI
        Dim dStart As Double
        Dim dEnd As Double
        Picture1.FillColor = lColor
        Picture1.FillStyle = 0
        dStart = fStart * (CircleEnd / 100)
        dEnd = fEnd * (CircleEnd / 100)
        Picture1.Circle (200, 190), 150, , dStart, dEnd
 End Sub
Private Sub Form_Load()
    Set con = New ADODB.Connection
        con.Open "Provider=Microsoft.Jet.Oledb.4.0; data source =" & App.Path & "\StudentSurveyDB.mdb; Persist Security Info=False;Jet OLEDB:"

'***** Set label Visibility *****'
    Label1.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    
'***** Load Database *****'
    loadGender
    loadBloodtype
    loadCourse
    loadCivilStatus
End Sub
Public Sub loadGender()
'***** Load Gender *****'
    Picture1.ScaleMode = vbPixels
    Set rx = New ADODB.Recordset
    rx.Open "select * from Student_Survey", con, adOpenStatic, adLockOptimistic
    rx.MoveFirst
    With rx
        While .EOF = False
            If rx!Student_Gender = "Male" Then
                male = male + 1
            ElseIf rx!Student_Gender = "Female" Then
                female = female + 1
            End If
        .MoveNext
        Wend
    dMale = male * 100 / (male + female)
    dFemale = female * 100 / (male + female)
    End With
    Set rx = Nothing
End Sub
Public Sub loadBloodtype()
'***** Load Bloodtype *****'
    Picture1.ScaleMode = vbPixels
    Set rx = New ADODB.Recordset
    rx.Open "select * from Student_Survey", con, adOpenStatic, adLockOptimistic
    rx.MoveFirst
    With rx
        While .EOF = False
            If rx!Student_Bloodtype = "A" Then
                A = A + 1
            ElseIf rx!Student_Bloodtype = "B" Then
                B = B + 1
            ElseIf rx!Student_Bloodtype = "O" Then
                O = O + 1
            'ElseIf rx!Student_Bloodtype = "Rh" Then
                'Rh = Rh + 1
            ElseIf rx!Student_Bloodtype = "AB" Then
                AB = AB + 1
            End If
        .MoveNext
        Wend
    bA = A * 100 / (A + B + O + AB)
    bB = B * 100 / (A + B + O + AB)
    bO = O * 100 / (A + B + O + AB)
    bAB = AB * 100 / (A + B + O + AB)
    'bRh = Rh * 100 / (A + B + O + AB + Rh)
    End With
    Set rx = Nothing
End Sub
Public Sub loadCourse()
'***** Load Course *****'
    Picture1.ScaleMode = vbPixels
    Set rx = New ADODB.Recordset
    rx.Open "select * from Student_Survey", con, adOpenStatic, adLockOptimistic
    rx.MoveFirst
    With rx
        While .EOF = False
            If rx!Student_Course = "BSHRM" Then
                BSHRM = BSHRM + 1
            ElseIf rx!Student_Course = "BEED" Then
                BEED = BEED + 1
            ElseIf rx!Student_Course = "BSIT" Then
                BSIT = BSIT + 1
            ElseIf rx!Student_Course = "BSCS" Then
                BSCS = BSCS + 1
            End If
        .MoveNext
        Wend
    cBSHRM = BSHRM * 100 / (BSHRM + BEED + BSIT + BSCS)
    cBEED = BEED * 100 / (BSHRM + BEED + BSIT + BSCS)
    cBSIT = BSIT * 100 / (BSHRM + BEED + BSIT + BSCS)
    cBSCS = BSCS * 100 / (BSHRM + BEED + BSIT + BSCS)
    End With
    Set rx = Nothing
End Sub
Public Sub loadCivilStatus()
'***** Load Civil Status *****'
    Picture1.ScaleMode = vbPixels
    Set rx = New ADODB.Recordset
    rx.Open "select * from Student_Survey", con, adOpenStatic, adLockOptimistic
    rx.MoveFirst
    With rx
        While .EOF = False
            If rx!Student_CivilStatus = "Single" Then
                iSingle = iSingle + 1
            ElseIf rx!Student_CivilStatus = "Married" Then
                iMarried = iMarried + 1
            End If
        .MoveNext
        Wend
    dSingle = iSingle * 100 / (iSingle + iMarried)
    dMarried = iMarried * 100 / (iSingle + iMarried)
    End With
    Set rx = Nothing
End Sub

