VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Student_Survey_DB 
   Caption         =   "Student Survey Information"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Search Student"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   720
      TabIndex        =   23
      Top             =   6120
      Width           =   5895
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   26
         Top             =   480
         Width           =   3615
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Student Information Action"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   4920
      Width           =   7095
      Begin VB.CommandButton btnExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Student Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   7320
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      Begin MSComctlLib.ListView ListView1 
         Height          =   7935
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   13996
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Age"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date of Birth"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Gender"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Civil Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Bloodtype"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Course"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Information"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1200
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   18
         Top             =   3240
         Width           =   5655
      End
      Begin VB.ComboBox cmbCourse 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Student_Survey_DB.frx":0000
         Left            =   1200
         List            =   "Student_Survey_DB.frx":0010
         TabIndex        =   16
         Top             =   4200
         Width           =   5655
      End
      Begin VB.ComboBox cmbBloodtype 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Student_Survey_DB.frx":0096
         Left            =   1200
         List            =   "Student_Survey_DB.frx":00A6
         TabIndex        =   15
         Top             =   3720
         Width           =   1575
      End
      Begin VB.ComboBox cmbCivilStatus 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Student_Survey_DB.frx":00B7
         Left            =   1200
         List            =   "Student_Survey_DB.frx":00C1
         TabIndex        =   14
         Top             =   2760
         Width           =   2055
      End
      Begin VB.ComboBox cmbGender 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Student_Survey_DB.frx":00D6
         Left            =   1200
         List            =   "Student_Survey_DB.frx":00E0
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpickBirthday 
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   95092737
         CurrentDate     =   42967
      End
      Begin VB.TextBox txtAge 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1200
         TabIndex        =   11
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label Label11 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Bloodtype"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Civil Status"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Birthday"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   615
      End
   End
End
Attribute VB_Name = "Student_Survey_DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AddStudent As Integer
Dim EditStudent As Integer
Private Sub btnDelete_Click()
'****** Deletes the data from the database *********'
    Set rx = New ADODB.Recordset
        sql = "delete * from Student_Survey where Student_Name='" & txtName & "'"
        rx.Open sql, con, adOpenDynamic, adLockPessimistic
        LoadDatabase
End Sub
Private Sub btnEdit_Click()
'***** Set Buttons *****'
    btnAdd.Enabled = False
    btnSave.Enabled = True
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    btnExit.Enabled = False
    btnSearch.Enabled = False
    
'***** Set Table *****'
    ListView1.Enabled = False
    
'***** Enable Textfields *****'
    Textfields_Enable
    txtName.Enabled = False
    
'***** Set AddStudent *****'
    EditStudent = 1
End Sub
Private Sub btnSave_Click()
'***** Check if the button is click *****'
If AddStudent = 1 Then

'***** Check if fields are empty *****'
    Dim isMissing As String
    
    If isEmptyTxt(txtID) = True Then
        isMissing = "ID."
    End If
    If isEmptyTxt(txtName) = True Then
        isMissing = isMissing & "Name."
    End If
    If isEmptyTxt(txtAge) = True Then
        isMissing = isMissing & "Age."
    End If
    If isEmptyCmb(cmbGender) = True Then
        isMissing = isMissing & "Gender."
    End If
    If isEmptyCmb(cmbCivilStatus) = True Then
        isMissing = isMissing & "Religion."
    End If
    If isEmptyTxt(txtAddress) = True Then
        isMissing = isMissing & "Address."
    End If
    If isEmptyCmb(cmbBloodtype) = True Then
        isMissing = isMissing & "Bloodtype."
    End If
    If isEmptyCmb(cmbCourse) = True Then
        isMissing = isMissing & "Course."
    End If
    If isMissing <> "" Then
        MsgBox "Operation cannot be completed, please make sure you fill the necessary data: " & isMissing, vbCritical, "Error"
        isSuccess
        Textfields_Disable
        Textfields_Clear
        Exit Sub
    End If
    
'***** Save the data to Database *****'
    Set rx = New ADODB.Recordset
    rx.Open "select * from Student_Survey where Student_Name='" & txtName.Text & "'", con, adOpenKeyset, adLockOptimistic
        If rx.RecordCount > 0 Then
            MsgBox "This name already exists!", vbCritical + vbOKOnly, "Error"
            Textfields_Disable
            Textfields_Clear
            Set rx = Nothing
            Exit Sub
        End If
            With rx
                .AddNew
                .Fields("ID") = txtID.Text
                .Fields("Student_Name") = txtName.Text
                .Fields("Student_Age") = txtAge.Text
                .Fields("Student_DateofBirth") = dtpickBirthday.Value
                .Fields("Student_Gender") = cmbGender.Text
                .Fields("Student_CivilStatus") = cmbCivilStatus.Text
                .Fields("Student_Address") = txtAddress.Text
                .Fields("Student_Bloodtype") = cmbBloodtype.Text
                .Fields("Student_Course") = cmbCourse.Text
                .Update
            End With
    MsgBox "Successfully save!", vbInformation + vbOKOnly
    Set rx = Nothing
    Call LoadDatabase
        
'***** Call isSuccess *****'
        isSuccess
        
'***** Set AddStudent *****'
        AddStudent = 0

ElseIf EditStudent = 1 Then

'***** Save the data to Database *****'
    Set rx = New ADODB.Recordset
    rx.Open "select * from Student_Survey where Student_Name='" & txtName.Text & "'", con, adOpenKeyset, adLockOptimistic
        With rx
            .Fields("ID") = txtID.Text
            .Fields("Student_Name") = txtName.Text
            .Fields("Student_Age") = txtAge.Text
            .Fields("Student_DateofBirth") = dtpickBirthday.Value
            .Fields("Student_Gender") = cmbGender.Text
            .Fields("Student_CivilStatus") = cmbCivilStatus.Text
            .Fields("Student_Address") = txtAddress.Text
            .Fields("Student_Bloodtype") = cmbBloodtype.Text
            .Fields("Student_Course") = cmbCourse.Text
            .Update
        End With
    MsgBox "Successfully save!", vbInformation + vbOKOnly
    Set rx = Nothing
    Call LoadDatabase

'***** Call isSuccess *****'
        isSuccess
    
'***** Set AddStudent *****'
        EditStudent = 0
End If
End Sub
Private Sub btnSearch_Click()
'***** Call SearchStudent *****'
    SearchStudent
End Sub
Private Sub Form_Load()
'************** Add the components *****************'
'*  1. Microsoft ADO Data Control (OLEDB)          *'
'*  2. Microsoft Chart Control 6.0                 *'
'*  3. Microsoft Common Dialog Control 6.0         *'
'*  4. Microsoft Windows Common Controls 6.0 (SP6) *'
'*  5. Microsoft Windows Common Controls-2 6.0     *'
'***************************************************'

'***** Initialize *****'
    Set Connect = New Student_Survey_Class1
    Call LoadDatabase
    Textfields_Disable
End Sub
Private Sub btnAdd_Click()
'***** Call Textfields_Clear *****'
    Textfields_Clear
    
'***** Set Buttons *****'
    btnAdd.Enabled = False
    btnSave.Enabled = True
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    btnExit.Enabled = False
    btnSearch.Enabled = False
    
'***** Set Table *****'
    ListView1.Enabled = False
    
'***** Call Textfields_Enable *****'
    Textfields_Enable
    
'***** Set AddStudent *****'
    AddStudent = 1
End Sub
Public Sub LoadDatabase()
'*********** Load data from database to the list ***********'
    ListView1.ListItems.Clear
    Set rx = New ADODB.Recordset
        With rx
            sql = "SELECT * FROM Student_Survey order by Student_Name"
                .Open sql, con, adOpenKeyset, adLockOptimistic
            Do While Not .EOF
                ListView1.ListItems.Add , , !ID
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !Student_Name
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !Student_Age
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !Student_DateofBirth
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !Student_Gender
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Student_CivilStatus
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !Student_Address
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & !Student_Bloodtype
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !Student_Course
                .MoveNext
            Loop
                .Close
        End With
        Set rx = Nothing
End Sub
Private Sub ListView1_Click()
'***** Selects the data on the list and outputs it to the Textfields *****'
    On Error Resume Next
        txtID.Text = ListView1.SelectedItem
        txtName.Text = ListView1.SelectedItem.SubItems(1)
        txtAge.Text = ListView1.SelectedItem.SubItems(2)
        dtpickBirthday.Value = ListView1.SelectedItem.SubItems(3)
        cmbGender.Text = ListView1.SelectedItem.SubItems(4)
        cmbCivilStatus.Text = ListView1.SelectedItem.SubItems(5)
        txtAddress.Text = ListView1.SelectedItem.SubItems(6)
        cmbBloodtype.Text = ListView1.SelectedItem.SubItems(7)
        cmbCourse.Text = ListView1.SelectedItem.SubItems(8)
End Sub
Public Sub Textfields_Disable()
'***** Set Textfields to Disable *****'
    txtID.Enabled = False
    txtName.Enabled = False
    txtAge.Enabled = False
    dtpickBirthday.Enabled = False
    cmbGender.Enabled = False
    cmbCivilStatus.Enabled = False
    txtAddress.Enabled = False
    cmbBloodtype.Enabled = False
    cmbCourse.Enabled = False
End Sub
Public Sub Textfields_Enable()
'***** Set Textfields to Enable *****'
    txtID.Enabled = True
    txtName.Enabled = True
    txtAge.Enabled = True
    dtpickBirthday.Enabled = True
    cmbGender.Enabled = True
    cmbCivilStatus.Enabled = True
    txtAddress.Enabled = True
    cmbBloodtype.Enabled = True
    cmbCourse.Enabled = True
End Sub
Public Sub Textfields_Clear()
'***** Clear Textfields *****'
    txtID.Text = ""
    txtName.Text = ""
    txtAge.Text = ""
    dtpickBirthday.Value = Date
    cmbGender.Text = ""
    cmbCivilStatus.Text = ""
    txtAddress.Text = ""
    cmbBloodtype.Text = ""
    cmbCourse.Text = ""
End Sub
Public Function isEmptyTxt(Txt As TextBox) As Boolean
'***** Fucntion Textbox Clear *****'
    isEmptyTxt = False
    If Len(Txt.Text) = 0 Then isEmptyTxt = True
    If Txt.Text = Null Then isEmptyTxt = True
End Function
Public Function isEmptyCmb(Txt As ComboBox) As Boolean
'***** Fucntion Combobox Clear *****'
    isEmptyCmb = False
    If Len(Txt.Text) = 0 Then isEmptyCmb = True
    If Txt.Text = Null Then isEmptyCmb = True
End Function
Public Sub isSuccess()
'***** Set Buttons *****'
        btnAdd.Enabled = True
        btnSave.Enabled = False
        btnEdit.Enabled = True
        btnDelete.Enabled = True
        btnExit.Enabled = True
        btnSearch.Enabled = True

'***** Set Table *****'
        ListView1.Enabled = True
    
'***** Disable Textfields *****'
        Textfields_Disable
End Sub
Public Function Textboxes()
'***** Set the textfield data from adodc *****'
        txtID.Text = rx!S
        txtName.Text = Adodc1.Recordset.Fields("Student_Name")
        txtAge.Text = Adodc1.Recordset.Fields("Student_Age")
        dtpickBirthday.Value = Adodc1.Recordset.Fields("Student_DateofBirth")
        cmbGender.Text = Adodc1.Recordset.Fields("Student_Gender")
        cmbCivilStatus.Text = Adodc1.Recordset.Fields("Student_CivilStatus")
        txtAddress.Text = Adodc1.Recordset.Fields("Student_Address")
        cmbBloodtype.Text = Adodc1.Recordset.Fields("Student_Bloodtype")
        cmbCourse.Text = Adodc1.Recordset.Fields("Student_Course")
End Function
Sub SearchStudent()
'***** Search Students on the database *****'
    ListView1.ListItems.Clear
    Set rx = New ADODB.Recordset
        With rx
            sql = "select * from Student_Survey  WHERE " _
                & " ID like '%" & txtSearch.Text & "%' OR  " _
                & " Student_Name like '%" & txtSearch.Text & "%' OR " _
                & " Student_Address like '%" & txtSearch.Text & "%' OR " _
                & " Student_Age like '%" & txtSearch.Text & "%' OR " _
                & " Student_Gender like '%" & txtSearch.Text & "%' OR " _
                & " Student_DateofBirth like '%" & txtSearch.Text & "%' OR " _
                & " Student_CivilStatus like '%" & txtSearch.Text & "%' OR " _
                & " Student_Course like '%" & txtSearch.Text & "%' OR " _
                & " Student_Bloodtype like '%" & txtSearch.Text & "%'"
                .Open sql, con, adOpenKeyset, adLockOptimistic
            Do While Not .EOF
                    ListView1.ListItems.Add , , !ID
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "" & !Student_Name
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "" & !Student_Age
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = "" & !Student_DateofBirth
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = "" & !Student_Gender
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = "" & !Student_CivilStatus
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = "" & !Student_Address
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = "" & !Student_Bloodtype
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = "" & !Student_Course
                .MoveNext
            Loop
                .Close
        End With
        Set rx = Nothing
End Sub
Private Sub txtSearch_Change()
    If txtSearch.Text = "" Then
        LoadDatabase
    End If
End Sub
