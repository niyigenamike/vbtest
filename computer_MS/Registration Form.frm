VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RegisterForm 
   BackColor       =   &H8000000D&
   Caption         =   "REGISTRATION FORM"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9825
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H80000011&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1440
      Top             =   7080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"Registration Form.frx":0000
      OLEDBString     =   $"Registration Form.frx":00A3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "user"
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
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   315
      Left            =   2400
      TabIndex        =   15
      Text            =   "Sex/Gender"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox passwd 
      BackColor       =   &H80000007&
      DataField       =   "password"
      DataSource      =   "RegistAdo"
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   4800
      Width           =   2895
   End
   Begin VB.TextBox uname 
      BackColor       =   &H80000007&
      DataField       =   "username"
      DataSource      =   "RegistAdo"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox phone 
      BackColor       =   &H80000007&
      DataField       =   "telephone"
      DataSource      =   "RegistAdo"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txtemail 
      BackColor       =   &H80000007&
      DataField       =   "email"
      DataSource      =   "RegistAdo"
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   2400
      TabIndex        =   11
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox lname 
      BackColor       =   &H80000007&
      DataField       =   "last_name"
      DataSource      =   "RegistAdo"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox fname 
      BackColor       =   &H80000007&
      DataField       =   "first_name"
      DataSource      =   "RegistAdo"
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cancelbtn 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton registbtn 
      Caption         =   "Register"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000008&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000008&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000008&
      Caption         =   "Telephone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "RegisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub registbtn_Click()
    ' Declare variables for the input fields
    Dim first_name As String, last_name As String, email As String, gender As String
    Dim telephone As String, username As String, password As String

    ' Retrieve input values from textboxes
    first_name = fname.Text
    last_name = lname.Text
    email = txtemail.Text
    telephone = phone.Text
    username = uname.Text
    password = passwd.Text
    gender = Combo1.Text ' Retrieve the gender selection

    ' Validate required fields
    If fname.Text = "" Or lname.Text = "" Or txtemail.Text = "" Or Combo1.Text = "" Or phone.Text = "" Or uname.Text = "" Or passwd.Text = "" Then
        MsgBox "Please fill all fields before registering.", vbExclamation, "Missing Information"
        Exit Sub
    End If
  ' Validate email format (simple check)
    If InStr(email, "@") = 0 Or InStr(email, ".") = 0 Then
        MsgBox "Please enter a valid email address.", vbExclamation, "Invalid Email"
        Exit Sub
    End If
    ' Database connection string
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MIKE\Downloads\computer_MS\computer_MS\database\computer_ms_db_production.mdb;Persist Security Info=False;"
    conn.Open

    ' Create the SQL command
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "INSERT INTO [user] (first_name, last_name, email, gender, telephone, username, [password]) " & _
                       "VALUES (?, ?, ?, ?, ?, ?, ?)"
        ' Append parameters
        .Parameters.Append .CreateParameter("firstname", 200, 1, 50, first_name) ' 200 = adVarChar, 1 = adParamInput
        .Parameters.Append .CreateParameter("lastname", 200, 1, 50, last_name)
        .Parameters.Append .CreateParameter("email", 200, 1, 50, email)
        .Parameters.Append .CreateParameter("gender", 200, 1, 10, gender)
        .Parameters.Append .CreateParameter("telephone", 200, 1, 15, telephone)
        .Parameters.Append .CreateParameter("username", 200, 1, 20, username)
        .Parameters.Append .CreateParameter("password", 200, 1, 20, password)
        ' Execute the command
        .Execute
    End With

    MsgBox "User registered successfully!", vbInformation, "Registration Complete"

    ' Navigate to the login form
    LoginForm.Show ' Assuming the login form is named LoginForm

    ' Close the registration form if desired
    Me.Hide

    ' Clean up
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during registration: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set cmd = Nothing
    Set conn = Nothing
End Sub

Private Sub Form_Load()
    ' Populate gender combo box
    Combo1.Clear ' Clear any existing items to avoid duplicates
    Combo1.AddItem "Female"
    Combo1.AddItem "Male"
End Sub

Private Sub CancelBtn_Click()
    ' Clear all input fields
    ClearInputFields
 
    ' Optionally, display a message confirming the cancellation
    MsgBox "Input fields have been cleared.", vbInformation, "Canceled"
End Sub

' Subroutine to clear all input fields
Private Sub ClearInputFields()
    fname.Text = ""
    lname.Text = ""
    txtemail.Text = ""
    phone.Text = ""
    uname.Text = ""
    passwd.Text = ""
    Combo1.Text = ""
End Sub


