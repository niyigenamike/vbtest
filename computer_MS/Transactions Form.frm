VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form TransactionForm 
   Caption         =   "TRANSACTION FORM"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11235
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1440
      Top             =   8760
      Width           =   3135
      _ExtentX        =   5530
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
      Connect         =   $"Transactions Form.frx":0000
      OLEDBString     =   $"Transactions Form.frx":00A3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblTransactions"
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   24
      Top             =   7320
      Width           =   2055
   End
   Begin VB.ComboBox ComboStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   23
      Top             =   6000
      Width           =   2415
   End
   Begin VB.TextBox txtPaymentMethod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtTransactionAmount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      TabIndex        =   21
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox txtTransactionDate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txtQuantity 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      TabIndex        =   19
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txtProductID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      TabIndex        =   18
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtOrderID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      TabIndex        =   17
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtTransactionID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      TabIndex        =   16
      Top             =   720
      Width           =   2415
   End
   Begin VB.PictureBox PictureBox1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   6000
      ScaleHeight     =   4995
      ScaleWidth      =   7875
      TabIndex        =   14
      Top             =   1440
      Width           =   7935
   End
   Begin VB.CommandButton ExitCommand 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4320
      TabIndex        =   13
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton UpdateCommand 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton CancelCommand 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton CommandRead 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton DeleteCommand 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton InsertCommand 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "All Transactions with Inventory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   15
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Transaction Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Payment Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Transaction Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Transaction Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Product Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Order Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Transaction Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "TransactionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub InsertCommand_Click()
    ' Declare variables for the input fields
    Dim transaction_id As String, order_id As String, product_id As String
    Dim Quantity As Integer, transaction_date As String, transaction_amount As Double
    Dim payment_method As String, status As String

    ' Retrieve input values from textboxes and combo box
    transaction_id = txtTransactionID.Text
    order_id = txtOrderID.Text
    product_id = txtProductID.Text
    Quantity = Val(txtQuantity.Text) ' Convert to integer
    transaction_date = txtTransactionDate.Text ' Convert to date
    transaction_amount = Val(txtTransactionAmount.Text) ' Convert to double
    payment_method = txtPaymentMethod.Text
    status = ComboStatus.Text ' Retrieve the status selection

    ' Validate required fields
    If transaction_id = "" Or order_id = "" Or product_id = "" Or txtQuantity.Text = "" Or transaction_date = "" Or _
       txtTransactionAmount.Text = "" Or payment_method = "" Or status = "" Then
        MsgBox "Please fill all fields before adding the transaction.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Validate numeric fields
    If Quantity <= 0 Then
        MsgBox "Quantity must be greater than 0.", vbExclamation, "Invalid Quantity"
        Exit Sub
    End If

    If transaction_amount <= 0 Then
        MsgBox "Transaction amount must be greater than 0.", vbExclamation, "Invalid Amount"
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
        .CommandText = "INSERT INTO tblTransactions (transaction_id, order_id, product_id, quantity, transaction_date, transaction_amount, payment_method, status) " & _
                       "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
        ' Append parameters
        .Parameters.Append .CreateParameter("transaction_id", 200, 1, Len(transaction_id), transaction_id) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("order_id", 200, 1, Len(order_id), order_id)
        .Parameters.Append .CreateParameter("product_id", 200, 1, Len(product_id), product_id)
        .Parameters.Append .CreateParameter("quantity", 3, 1, , Quantity) ' 3 = adInteger
        .Parameters.Append .CreateParameter("transaction_date", 7, 1, , transaction_date) ' 7 = adDate
        .Parameters.Append .CreateParameter("transaction_amount", 5, 1, , transaction_amount) ' 5 = adDouble
        .Parameters.Append .CreateParameter("payment_method", 200, 1, Len(payment_method), payment_method)
        .Parameters.Append .CreateParameter("status", 200, 1, Len(status), status)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Transaction added successfully!", vbInformation, "Registration Complete"

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

Private Sub CancelCommand_Click()
    ' Clear all input fields
    ClearInputFields

    ' Notify user of cancellation
    MsgBox "Input fields have been cleared.", vbInformation, "Canceled"
End Sub

Private Sub Form_Load()
    ' Populate status combo box
    ComboStatus.AddItem "Pending"
    ComboStatus.AddItem "Completed"
    ComboStatus.AddItem "Cancelled"
    ComboStatus.AddItem "Shipped"
End Sub

' Subroutine to clear all input fields
Private Sub ClearInputFields()
    txtTransactionID.Text = ""
    txtOrderID.Text = ""
    txtProductID.Text = ""
    txtQuantity.Text = ""
    txtTransactionDate.Text = ""
    txtTransactionAmount.Text = ""
    txtPaymentMethod.Text = ""
    ComboStatus.Text = ""
End Sub

Private Sub ExitCommand_Click()
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
    sql = "SELECT transaction_id, order_id, product_id, quantity, transaction_date, transaction_amount, payment_method, status FROM tblTransactions"

    ' Execute the query
    Set rs = conn.Execute(sql)

    ' Clear the PictureBox
    PictureBox1.Cls
    PictureBox1.Font.Size = 10
    PictureBox1.Font.Name = "Arial"

    ' Display column headers
    Dim header As String
    header = "Transaction ID" & vbTab & "Order ID" & vbTab & "Product ID" & vbTab & "Quantity" & vbTab & "Transaction Date" & vbTab & "Amount" & vbTab & "Payment Method" & vbTab & "Status"
    PictureBox1.Print header
    PictureBox1.Line (0, 30)-(PictureBox1.ScaleWidth, 30), vbBlack

    ' Display data rows
    Dim y As Single
    y = 45
    Do While Not rs.EOF
        Dim line As String
        line = rs("transaction_id") & vbTab & rs("order_id") & vbTab & rs("product_id") & vbTab & rs("quantity") & vbTab & rs("transaction_date") & vbTab & rs("transaction_amount") & vbTab & rs("payment_method") & vbTab & rs("status")
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

Private Sub UpdateCommand_Click()
    ' Declare variables for the input fields
    Dim transaction_id As String, order_id As String, product_id As String
    Dim Quantity As String, transaction_date As String
    Dim transaction_amount As Double, payment_method As String, status As String

    ' Retrieve input values from textboxes and combo box
    transaction_id = txtTransactionID.Text
    order_id = txtOrderID.Text
    product_id = txtProductID.Text
    Quantity = txtQuantity.Text ' Convert to integer
    transaction_date = txtTransactionDate.Text
    transaction_amount = Val(txtTransactionAmount.Text) ' Convert to double
    payment_method = txtPaymentMethod.Text
    status = ComboStatus.Text ' Retrieve the status selection

    ' Validate required fields
    If txtTransactionID.Text = "" Or txtOrderID.Text = "" Or txtProductID.Text = "" Or txtQuantity.Text = "" Or _
       txtTransactionDate.Text = "" Or txtTransactionAmount.Text = "" Or txtPaymentMethod.Text = "" Or ComboStatus.Text = "" Then
        MsgBox "Please fill all fields before updating the transaction.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Validate numeric fields
    If Val(Quantity) <= 0 Then
        MsgBox "Quantity must be greater than 0.", vbExclamation, "Invalid Quantity"
        Exit Sub
    End If

    If transaction_amount <= 0 Then
        MsgBox "Transaction amount must be greater than 0.", vbExclamation, "Invalid Transaction Amount"
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
        .CommandText = "UPDATE tblTransactions SET order_id = ?, product_id = ?, quantity = ?, transaction_date = ?, " & _
                       "transaction_amount = ?, payment_method = ?, status = ? WHERE transaction_id = ?"
        ' Append parameters
        .Parameters.Append .CreateParameter("order_id", 200, 1, Len(order_id), order_id)
        .Parameters.Append .CreateParameter("product_id", 200, 1, Len(product_id), product_id)
        .Parameters.Append .CreateParameter("quantity", 3, 1, , Val(Quantity)) ' 3 = adInteger
        .Parameters.Append .CreateParameter("transaction_date", 200, 1, Len(transaction_date), transaction_date)
        .Parameters.Append .CreateParameter("transaction_amount", 5, 1, , transaction_amount) ' 5 = adDouble
        .Parameters.Append .CreateParameter("payment_method", 200, 1, Len(payment_method), payment_method)
        .Parameters.Append .CreateParameter("status", 200, 1, Len(status), status)
        .Parameters.Append .CreateParameter("transaction_id", 200, 1, Len(transaction_id), transaction_id)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Transaction updated successfully!", vbInformation, "Update Complete"

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
Private Sub DeleteCommand_Click()
    ' Declare a variable for the transaction ID
    Dim transaction_id As String

    ' Retrieve the transaction ID from the textbox
    transaction_id = txtTransactionID.Text

    ' Validate that the transaction ID field is not empty
    If txtTransactionID.Text = "" Then
        MsgBox "Please enter the Transaction ID to delete the transaction.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Confirm deletion with the user
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo, "Confirm Deletion")
    If response = vbNo Then Exit Sub

    ' Database connection and command objects
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler ' Add error handling

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MIKE\Downloads\computer_MS\computer_MS\database\computer_ms_db_production.mdb;Persist Security Info=False;"
    conn.Open

    ' Prepare the SQL command for deleting the transaction
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "DELETE FROM tblTransactions WHERE transaction_id = ?"
        ' Append parameter
        .Parameters.Append .CreateParameter("transaction_id", 200, 1, Len(transaction_id), transaction_id) ' 200 = adVarChar
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Transaction deleted successfully!", vbInformation, "Deletion Complete"

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
