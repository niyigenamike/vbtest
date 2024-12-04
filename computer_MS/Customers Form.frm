VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form CustomerForm 
   BackColor       =   &H8000000D&
   Caption         =   "CUSTOMERS FORM"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   2160
      Top             =   7680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1296
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
      Connect         =   $"Customers Form.frx":0000
      OLEDBString     =   $"Customers Form.frx":00A3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblCustomers"
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
   Begin VB.CommandButton CommandBack 
      Caption         =   "Back to Home"
      Height          =   495
      Left            =   10200
      TabIndex        =   24
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   14520
      TabIndex        =   23
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton SearchCommand 
      Caption         =   "Search"
      Height          =   495
      Left            =   17160
      TabIndex        =   22
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   315
      Left            =   2520
      TabIndex        =   21
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton ExitBtn 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4800
      TabIndex        =   20
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton CancelBtn 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   19
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton CommandDelete 
      Caption         =   "Delect customer"
      Height          =   495
      Left            =   600
      TabIndex        =   18
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton CommandUpdate 
      Caption         =   "Update customer"
      Height          =   615
      Left            =   4560
      TabIndex        =   17
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton CommandRead 
      Caption         =   "View all customers"
      Height          =   615
      Left            =   2640
      TabIndex        =   16
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton CommandInsert 
      BackColor       =   &H8000000D&
      Caption         =   "Add A new customer"
      Height          =   615
      Left            =   600
      TabIndex        =   15
      Top             =   5760
      Width           =   1455
   End
   Begin VB.PictureBox PictureBox1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   3975
      Left            =   6480
      ScaleHeight     =   3915
      ScaleWidth      =   10395
      TabIndex        =   14
      Top             =   1680
      Width           =   10455
   End
   Begin VB.TextBox txtShippingAddress 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   4920
      Width           =   3255
   End
   Begin VB.TextBox txtemail 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   2520
      TabIndex        =   11
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox txtContactNumber 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   2520
      TabIndex        =   10
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox txtLastName 
      BackColor       =   &H80000008&
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox txtFirstName 
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   405
      Left            =   2520
      TabIndex        =   8
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtCustomerID 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "View All customers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   8040
      TabIndex        =   13
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000008&
      Caption         =   "Shiping Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      Caption         =   "Customer Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000008&
      Caption         =   "Customer Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000008&
      Caption         =   "Customer Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Customer Last NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Customer First NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Customer Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "CustomerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandInsert_Click()
    ' Declare variables for the input fields
    Dim customer_id As String, first_name As String, last_name As String
    Dim contact_number As String, email As String, shipping_address As String, gender As String

    ' Retrieve input values from textboxes and combo box
    customer_id = txtCustomerID.Text
    first_name = txtFirstName.Text
    last_name = txtLastName.Text
    contact_number = txtContactNumber.Text
    email = txtemail.Text
    shipping_address = txtShippingAddress.Text
    gender = Combo1.Text ' Retrieve the gender selection

    ' Validate required fields
    If customer_id = "" Or first_name = "" Or last_name = "" Or contact_number = "" Or email = "" Or shipping_address = "" Or gender = "" Then
        MsgBox "Please fill all fields before adding the customer.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Validate email format (simple check)
    If InStr(email, "@") = 0 Or InStr(email, ".") = 0 Then
        MsgBox "Please enter a valid email address.", vbExclamation, "Invalid Email"
        Exit Sub
    End If

    ' Database connection and command objects
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler ' Add error handling

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MIKE\Downloads\computer_MS\computer_MS\database\computer_ms_db_production.mdb;Persist Security Info=False;"
    conn.Open

    ' Prepare the SQL command for inserting data
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "INSERT INTO tblCustomers (customer_id, first_name, last_name, contact_number, gender, email, shipping_address) " & _
                       "VALUES (?, ?, ?, ?, ?, ?, ?)"
        ' Append parameters
        .Parameters.Append .CreateParameter("customer_id", 200, 1, Len(customer_id), customer_id) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("first_name", 200, 1, Len(first_name), first_name)
        .Parameters.Append .CreateParameter("last_name", 200, 1, Len(last_name), last_name)
        .Parameters.Append .CreateParameter("contact_number", 200, 1, Len(contact_number), contact_number)
        .Parameters.Append .CreateParameter("gender", 200, 1, Len(gender), gender)
        .Parameters.Append .CreateParameter("email", 200, 1, Len(email), email)
        .Parameters.Append .CreateParameter("shipping_address", 200, 1, Len(shipping_address), shipping_address)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Customer added successfully!", vbInformation, "Registration Complete"

    ' Clear input fields
    ClearInputFields

    ' Clean up
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    ' Handle errors
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
End Sub

Private Sub CancelBtn_Click()
    ' Clear all input fields
    ClearInputFields

    ' Notify user of cancellation
    MsgBox "Input fields have been cleared.", vbInformation, "Canceled"
End Sub

Private Sub Form_Load()
    ' Populate gender combo box
    Combo1.AddItem "Female"
    Combo1.AddItem "Male"
End Sub

' Subroutine to clear all input fields
Private Sub ClearInputFields()
    txtCustomerID.Text = ""
    txtFirstName.Text = ""
    txtLastName.Text = ""
    txtContactNumber.Text = ""
    txtemail.Text = ""
    txtShippingAddress.Text = ""
    Combo1.Text = ""
End Sub

Private Sub ExitBtn_Click()
    ' Close the form properly
    Unload Me
End Sub
Private Sub CommandRead_Click()
    
    Dim conn As Object
    Dim rs As Object
    On Error GoTo ErrorHandler

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MIKE\Downloads\computer_MS\computer_MS\database\computer_ms_db_production.mdb;Persist Security Info=False;"
    conn.Open

    ' Query to retrieve data
    Dim sql As String
    sql = "SELECT customer_id, first_name, last_name, contact_number, gender, email, shipping_address FROM tblCustomers"

    ' Execute the query
    Set rs = conn.Execute(sql)

    ' Clear the PictureBox
    PictureBox1.Cls
    PictureBox1.Font.Size = 10
    PictureBox1.Font.Name = "Arial"

    ' Display column headers
    Dim header As String
    header = "Customer ID" & vbTab & "First Name" & vbTab & "Last Name" & vbTab & "Contact No." & vbTab & "Gender" & vbTab & "Email" & vbTab & "Shipping Address"
    PictureBox1.Print header
    PictureBox1.Line (0, 30)-(PictureBox1.ScaleWidth, 30), vbBlack

    ' Display data rows
    Dim y As Single
    y = 45
    Do While Not rs.EOF
        Dim line As String
        line = rs("customer_id") & vbTab & rs("first_name") & vbTab & rs("last_name") & vbTab & rs("contact_number") & vbTab & rs("gender") & vbTab & rs("email") & vbTab & rs("shipping_address")
        PictureBox1.Print line
        y = y + 15
        rs.MoveNext
    Loop

    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub


Private Sub CommandUpdate_Click()
    ' Database connection and command objects
    Dim conn As Object
    Dim cmd As Object
    Dim customer_id As String

    ' Retrieve customer_id for identifying the record
    customer_id = txtCustomerID.Text

    ' Validate that customer_id is entered
    If customer_id = "" Then
        MsgBox "Please enter the Customer ID of the record to update.", vbExclamation, "Missing ID"
        Exit Sub
    End If

    ' Retrieve updated data from input fields
    Dim first_name As String, last_name As String, contact_number As String
    Dim email As String, shipping_address As String, gender As String
    first_name = txtFirstName.Text
    last_name = txtLastName.Text
    contact_number = txtContactNumber.Text
    email = txtemail.Text
    shipping_address = txtShippingAddress.Text
    gender = Combo1.Text

    ' Validate required fields
    If first_name = "" Or last_name = "" Or contact_number = "" Or email = "" Or shipping_address = "" Or gender = "" Then
        MsgBox "Please fill all fields before updating the record.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MIKE\Downloads\computer_MS\computer_MS\database\computer_ms_db_production.mdb;Persist Security Info=False;"
    conn.Open

    ' Prepare the SQL command for updating data
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "UPDATE tblCustomers SET first_name = ?, last_name = ?, contact_number = ?, gender = ?, email = ?, shipping_address = ? WHERE customer_id = ?"
        ' Append parameters
        .Parameters.Append .CreateParameter("first_name", 200, 1, Len(first_name), first_name)
        .Parameters.Append .CreateParameter("last_name", 200, 1, Len(last_name), last_name)
        .Parameters.Append .CreateParameter("contact_number", 200, 1, Len(contact_number), contact_number)
        .Parameters.Append .CreateParameter("gender", 200, 1, Len(gender), gender)
        .Parameters.Append .CreateParameter("email", 200, 1, Len(email), email)
        .Parameters.Append .CreateParameter("shipping_address", 200, 1, Len(shipping_address), shipping_address)
        .Parameters.Append .CreateParameter("customer_id", 200, 1, Len(customer_id), customer_id)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Customer record updated successfully!", vbInformation, "Update Complete"

    ' Clear input fields
    ClearInputFields

    ' Clean up
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
End Sub
Private Sub CommandDelete_Click()
    ' Database connection and command objects
    Dim conn As Object
    Dim cmd As Object
    Dim customer_id As String

    ' Retrieve customer_id for identifying the record
    customer_id = txtCustomerID.Text

    ' Validate that customer_id is entered
    If customer_id = "" Then
        MsgBox "Please enter the Customer ID of the record to delete.", vbExclamation, "Missing ID"
        Exit Sub
    End If

    ' Confirm deletion
    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
        Exit Sub
    End If

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MIKE\Downloads\computer_MS\computer_MS\database\computer_ms_db_production.mdb;Persist Security Info=False;"
    conn.Open

    ' Prepare the SQL command for deleting data
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "DELETE FROM tblCustomers WHERE customer_id = ?"
        .Parameters.Append .CreateParameter("customer_id", 200, 1, Len(customer_id), customer_id)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Customer record deleted successfully!", vbInformation, "Delete Complete"

    ' Clear input fields
    ClearInputFields

    ' Clean up
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
End Sub


Private Sub CommandBack_Click()
DashboardForm.Show

End Sub
