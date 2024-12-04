VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form SupplierForm 
   BackColor       =   &H8000000D&
   Caption         =   "SUPPLIER FORM"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   960
      Top             =   6720
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
      Connect         =   $"Supplier Form.frx":0000
      OLEDBString     =   $"Supplier Form.frx":00A3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblsupplier"
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
      Height          =   375
      Left            =   8520
      TabIndex        =   18
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3960
      TabIndex        =   17
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton CommandCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2280
      TabIndex        =   16
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton CommandDelete 
      Caption         =   "Delete "
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton CommandUpdate 
      Caption         =   "Update New"
      Height          =   615
      Left            =   3960
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton CommandRead 
      Caption         =   "View Supplier"
      Height          =   615
      Left            =   2280
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton CommandInsert 
      Caption         =   "Add Supplier"
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox PictureBox1 
      BackColor       =   &H80000008&
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
      Height          =   2895
      Left            =   5640
      ScaleHeight     =   2835
      ScaleWidth      =   8235
      TabIndex        =   11
      Top             =   1080
      Width           =   8295
   End
   Begin VB.TextBox txtProductID 
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txtemail 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   2160
      TabIndex        =   8
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtContactNumber 
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   405
      Left            =   2160
      TabIndex        =   7
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtSupplierName 
      BackColor       =   &H80000008&
      ForeColor       =   &H80000005&
      Height          =   405
      Left            =   2160
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtSupplierID 
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   405
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "VIEW ALL SUPPLIERS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   7680
      TabIndex        =   10
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000008&
      Caption         =   "Product Id"
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
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000008&
      Caption         =   "Email"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Contact Number"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Supplier Name"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Supplier Id"
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
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "SupplierForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandInsert_Click()
    ' Declare variables for the input fields
    Dim supplier_id As String, supplier_name As String, contact_number As String
    Dim email As String, product_id As String

    ' Retrieve input values from textboxes
    supplier_id = txtSupplierID.Text
    supplier_name = txtSupplierName.Text
    contact_number = txtContactNumber.Text
    email = txtemail.Text
    product_id = txtProductID.Text

    ' Validate required fields
    If supplier_id = "" Or supplier_name = "" Or contact_number = "" Or email = "" Or product_id = "" Then
        MsgBox "Please fill all fields before adding the supplier.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Validate email format
    If InStr(1, email, "@") = 0 Or InStr(1, email, ".") = 0 Then
        MsgBox "Please enter a valid email address.", vbExclamation, "Invalid Email"
        Exit Sub
    End If

    ' Validate contact number (ensure it's numeric)
    If Not IsNumeric(contact_number) Then
        MsgBox "Contact number must contain only numbers.", vbExclamation, "Invalid Contact Number"
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
        .CommandText = "INSERT INTO tblsupplier (supplier_id, supplier_name, contact_number, email, product_id) " & _
                       "VALUES (?, ?, ?, ?, ?)"
        ' Append parameters
        .Parameters.Append .CreateParameter("supplier_id", 200, 1, Len(supplier_id), supplier_id) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("supplier_name", 200, 1, Len(supplier_name), supplier_name)
        .Parameters.Append .CreateParameter("contact_number", 200, 1, Len(contact_number), contact_number)
        .Parameters.Append .CreateParameter("email", 200, 1, Len(email), email)
        .Parameters.Append .CreateParameter("product_id", 200, 1, Len(product_id), product_id)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Supplier added successfully!", vbInformation, "Registration Complete"

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

Private Sub CommandCancel_Click()
    ' Clear all input fields
    ClearInputFields

    ' Notify user of cancellation
    MsgBox "Input fields have been cleared.", vbInformation, "Canceled"
End Sub

' Subroutine to clear all input fields
Private Sub ClearInputFields()
    txtSupplierID.Text = ""
    txtSupplierName.Text = ""
    txtContactNumber.Text = ""
    txtemail.Text = ""
    txtProductID.Text = ""
End Sub

Private Sub CommandExit_Click()
    ' Close the form properly
   End
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
    sql = "SELECT supplier_id, supplier_name, contact_number, email, product_id FROM tblsupplier"

    ' Execute the query
    Set rs = conn.Execute(sql)

    ' Clear the PictureBox
    PictureBox1.Cls
    PictureBox1.Font.Size = 10
    PictureBox1.Font.Name = "Arial"

    ' Display column headers
    Dim header As String
    header = "Supplier ID" & vbTab & "Supplier Name" & vbTab & "Contact No." & vbTab & "Email" & vbTab & "Product ID"
    PictureBox1.Print header
    PictureBox1.Line (0, 30)-(PictureBox1.ScaleWidth, 30), vbBlack

    ' Display data rows
    Dim y As Single
    y = 45
    Do While Not rs.EOF
        Dim line As String
        line = rs("supplier_id") & vbTab & rs("supplier_name") & vbTab & rs("contact_number") & vbTab & rs("email") & vbTab & rs("product_id")
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
    ' Declare variables for the input fields
    Dim supplier_id As String, supplier_name As String, contact_number As String
    Dim email As String, product_id As String

    ' Retrieve input values from textboxes
    supplier_id = txtSupplierID.Text
    supplier_name = txtSupplierName.Text
    contact_number = txtContactNumber.Text
    email = txtemail.Text
    product_id = txtProductID.Text

    ' Validate required fields
    If supplier_id = "" Or supplier_name = "" Or contact_number = "" Or email = "" Or product_id = "" Then
        MsgBox "Please fill all fields before updating the supplier.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Validate email format
    If InStr(1, email, "@") = 0 Or InStr(1, email, ".") = 0 Then
        MsgBox "Please enter a valid email address.", vbExclamation, "Invalid Email"
        Exit Sub
    End If

    ' Validate contact number (ensure it's numeric)
    If Not IsNumeric(contact_number) Then
        MsgBox "Contact number must contain only numbers.", vbExclamation, "Invalid Contact Number"
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

    ' Prepare the SQL command for updating data
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "UPDATE tblsupplier SET supplier_name = ?, contact_number = ?, email = ?, product_id = ? WHERE supplier_id = ?"
        ' Append parameters
        .Parameters.Append .CreateParameter("supplier_name", 200, 1, Len(supplier_name), supplier_name)
        .Parameters.Append .CreateParameter("contact_number", 200, 1, Len(contact_number), contact_number)
        .Parameters.Append .CreateParameter("email", 200, 1, Len(email), email)
        .Parameters.Append .CreateParameter("product_id", 200, 1, Len(product_id), product_id)
        .Parameters.Append .CreateParameter("supplier_id", 200, 1, Len(supplier_id), supplier_id)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Supplier updated successfully!", vbInformation, "Update Complete"

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
Private Sub CommandDelete_Click()
    ' Declare variable for supplier_id
    Dim supplier_id As String

    ' Retrieve input value from textbox
    supplier_id = txtSupplierID.Text

    ' Validate required field
    If supplier_id = "" Then
        MsgBox "Please enter the Supplier ID to delete.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Confirm deletion
    If MsgBox("Are you sure you want to delete this supplier?", vbYesNo + vbQuestion, "Confirm Deletion") = vbNo Then
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

    ' Prepare the SQL command for deleting data
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "DELETE FROM tblsupplier WHERE supplier_id = ?"
        ' Append parameter
        .Parameters.Append .CreateParameter("supplier_id", 200, 1, Len(supplier_id), supplier_id)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Supplier deleted successfully!", vbInformation, "Deletion Complete"

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

Private Sub CommandBack_Click()
DashboardForm.Show

End Sub

