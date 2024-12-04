VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form OrderForm 
   BackColor       =   &H8000000D&
   Caption         =   "ORDERS FORM"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1440
      Top             =   6360
      Width           =   2535
      _ExtentX        =   4471
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
      Connect         =   $"Orders Form.frx":0000
      OLEDBString     =   $"Orders Form.frx":00A3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblOrders"
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
      Width           =   2175
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3720
      TabIndex        =   17
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton CommandCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2160
      TabIndex        =   16
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton CommandDelete 
      Caption         =   "Delete"
      Height          =   615
      Left            =   480
      TabIndex        =   15
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton CommandUpdate 
      Caption         =   "Update Order"
      Height          =   615
      Left            =   3720
      TabIndex        =   14
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton CommandRead 
      Caption         =   "View Order"
      Height          =   555
      Left            =   2160
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton CommandInsert 
      Caption         =   "Insert Order"
      Height          =   555
      Left            =   480
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.PictureBox PictureBox1 
      BackColor       =   &H80000007&
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
      Height          =   2775
      Left            =   5280
      ScaleHeight     =   2715
      ScaleWidth      =   8115
      TabIndex        =   10
      Top             =   1320
      Width           =   8175
   End
   Begin VB.TextBox txtTotalAmount 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   2280
      TabIndex        =   9
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtOrderDate 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   2280
      TabIndex        =   8
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtCustomerID 
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   405
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtOrderID 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   2280
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.ComboBox ComboOrderStatus 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "ALL CUSTOMERS"
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
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000008&
      Caption         =   "Order Status"
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
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000008&
      Caption         =   "Total Amount"
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
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      Caption         =   "Order Date"
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
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
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
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Order Id"
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
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "OrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandInsert_Click()
    ' Declare variables for the input fields
    Dim order_id As String, customer_id As String, order_date As String
    Dim order_status As String, total_amount As Double

    ' Retrieve input values from textboxes and combo box
    order_id = txtOrderID.Text
    customer_id = txtCustomerID.Text
    order_date = txtOrderDate.Text
    order_status = ComboOrderStatus.Text ' Retrieve the order status selection
    total_amount = Val(txtTotalAmount.Text) ' Convert to double

    ' Validate required fields
    If order_id = "" Or customer_id = "" Or order_date = "" Or order_status = "" Or txtTotalAmount.Text = "" Then
        MsgBox "Please fill all fields before adding the order.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Validate numeric fields
    If total_amount <= 0 Then
        MsgBox "Total amount must be greater than 0.", vbExclamation, "Invalid Amount"
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
        .CommandText = "INSERT INTO tblOrders (order_id, customer_id, order_date, order_status, total_amount) " & _
                       "VALUES (?, ?, ?, ?, ?)"
        ' Append parameters
        .Parameters.Append .CreateParameter("order_id", 200, 1, Len(order_id), order_id) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("customer_id", 200, 1, Len(customer_id), customer_id)
        .Parameters.Append .CreateParameter("order_date", 200, 1, Len(order_date), order_date)
        .Parameters.Append .CreateParameter("order_status", 200, 1, Len(order_status), order_status)
        .Parameters.Append .CreateParameter("total_amount", 5, 1, , total_amount) ' 5 = adDouble
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Order added successfully!", vbInformation, "Registration Complete"

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

Private Sub Form_Load()
    ' Populate order status combo box
    ComboOrderStatus.AddItem "Pending"
    ComboOrderStatus.AddItem "Processing"
    ComboOrderStatus.AddItem "Shipped"
    ComboOrderStatus.AddItem "Delivered"
    ComboOrderStatus.AddItem "Cancelled"
End Sub

' Subroutine to clear all input fields
Private Sub ClearInputFields()
    txtOrderID.Text = ""
    txtCustomerID.Text = ""
    txtOrderDate.Text = ""
    txtTotalAmount.Text = ""
    ComboOrderStatus.Text = ""
End Sub

Private Sub CommandExit_Click()
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
    sql = "SELECT order_id, customer_id, order_date, order_status, total_amount FROM tblOrders"

    ' Execute the query
    Set rs = conn.Execute(sql)

    ' Clear the PictureBox
    PictureBox1.Cls
    PictureBox1.Font.Size = 10
    PictureBox1.Font.Name = "Arial"

    ' Display column headers
    Dim header As String
    header = "Order ID" & vbTab & "Customer ID" & vbTab & "Order Date" & vbTab & "Order Status" & vbTab & "Total Amount"
    PictureBox1.Print header
    PictureBox1.Line (0, 30)-(PictureBox1.ScaleWidth, 30), vbBlack

    ' Display data rows
    Dim y As Single
    y = 45
    Do While Not rs.EOF
        Dim line As String
        line = rs("order_id") & vbTab & rs("customer_id") & vbTab & rs("order_date") & vbTab & rs("order_status") & vbTab & FormatCurrency(rs("total_amount"))
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
    Dim order_id As String, customer_id As String, order_date As String
    Dim order_status As String, total_amount As Double

    ' Retrieve input values from textboxes and combo box
    order_id = txtOrderID.Text
    customer_id = txtCustomerID.Text
    order_date = txtOrderDate.Text
    order_status = ComboOrderStatus.Text ' Retrieve the order status selection
    total_amount = Val(txtTotalAmount.Text) ' Convert to double

    ' Validate required fields
    If order_id = "" Or customer_id = "" Or order_date = "" Or order_status = "" Or txtTotalAmount.Text = "" Then
        MsgBox "Please fill all fields before updating the order.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Validate numeric fields
    If total_amount <= 0 Then
        MsgBox "Total amount must be greater than 0.", vbExclamation, "Invalid Amount"
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
        .CommandText = "UPDATE tblOrders SET customer_id = ?, order_date = ?, order_status = ?, total_amount = ? WHERE order_id = ?"
        ' Append parameters
        .Parameters.Append .CreateParameter("customer_id", 200, 1, Len(customer_id), customer_id) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("order_date", 200, 1, Len(order_date), order_date)
        .Parameters.Append .CreateParameter("order_status", 200, 1, Len(order_status), order_status)
        .Parameters.Append .CreateParameter("total_amount", 5, 1, , total_amount) ' 5 = adDouble
        .Parameters.Append .CreateParameter("order_id", 200, 1, Len(order_id), order_id)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Order updated successfully!", vbInformation, "Update Complete"

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
    ' Declare a variable for the order ID
    Dim order_id As String

    ' Retrieve the order ID from the textbox
    order_id = txtOrderID.Text

    ' Validate that the order ID field is not empty
    If order_id = "" Then
        MsgBox "Please enter the Order ID to delete the order.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Confirm deletion with the user
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to delete this order?", vbQuestion + vbYesNo, "Confirm Deletion")
    If response = vbNo Then Exit Sub

    ' Database connection and command objects
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler ' Add error handling

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MIKE\Downloads\computer_MS\computer_MS\database\computer_ms_db_production.mdb;Persist Security Info=False;"
    conn.Open

    ' Prepare the SQL command for deleting the order
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "DELETE FROM tblOrders WHERE order_id = ?"
        ' Append parameter
        .Parameters.Append .CreateParameter("order_id", 200, 1, Len(order_id), order_id) ' 200 = adVarChar
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Order deleted successfully!", vbInformation, "Deletion Complete"

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
