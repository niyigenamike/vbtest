VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ProductForm 
   BackColor       =   &H8000000D&
   Caption         =   "PRODUCT FORM"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   14445
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1560
      Top             =   6840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
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
      Connect         =   $"Products Form.frx":0000
      OLEDBString     =   $"Products Form.frx":00A3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblProducts"
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
      Height          =   435
      Left            =   8880
      TabIndex        =   18
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox Quantity 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   2760
      TabIndex        =   17
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox txtProductNames 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   2760
      TabIndex        =   16
      Top             =   1200
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   315
      Left            =   2760
      TabIndex        =   15
      Top             =   2040
      Width           =   3015
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
      Height          =   3015
      Left            =   6600
      ScaleHeight     =   2955
      ScaleWidth      =   7035
      TabIndex        =   14
      Top             =   1680
      Width           =   7095
   End
   Begin VB.CommandButton ExitBtn 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4560
      TabIndex        =   12
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton CancelBtn 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2520
      TabIndex        =   11
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton DeleteBtn 
      Caption         =   "Delete"
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton UpdateBtn 
      Caption         =   "Update New"
      Height          =   615
      Left            =   4560
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton ViewBtn 
      Caption         =   "View Product"
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton InsertBtn 
      Caption         =   "Add Product"
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox txtPrice 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox txtProductID 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000008&
      Caption         =   "ALL COMPUTERS IN STOCK"
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
      Height          =   975
      Left            =   8400
      TabIndex        =   13
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Price"
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
      Left            =   600
      TabIndex        =   4
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "Quantity In Stock"
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
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      Caption         =   "Product Category"
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
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Caption         =   "Product Name"
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
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
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
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "ProductForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub InsertBtn_Click()
    ' Declare variables for the input fields
    Dim product_id As String, product_name As String, prod_category As String
    Dim quantity_in_stock As String, price As Double

    ' Retrieve input values from textboxes and combo box
    product_id = txtProductID.Text
    product_name = txtProductNames.Text
    prod_category = Combo1.Text ' Retrieve the category selection
    quantity_in_stock = Quantity.Text ' Convert to integer
    price = Val(txtPrice.Text) ' Convert to double

    ' Validate required fields
    If txtProductID.Text = "" Or txtProductNames.Text = "" Or Combo1.Text = "" Or Quantity.Text = "" Or txtPrice.Text = "" Then
        MsgBox "Please fill all fields before adding the product.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Validate numeric fields
    If quantity_in_stock <= 0 Then
        MsgBox "Quantity in stock must be greater than 0.", vbExclamation, "Invalid Quantity"
        Exit Sub
    End If

    If price <= 0 Then
        MsgBox "Price must be greater than 0.", vbExclamation, "Invalid Price"
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
        .CommandText = "INSERT INTO tblProducts (product_id, product_name, prod_category, quantity_in_stock, price) " & _
                       "VALUES (?, ?, ?, ?, ?)"
        ' Append parameters
        .Parameters.Append .CreateParameter("product_id", 200, 1, Len(product_id), product_id) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("product_name", 200, 1, Len(product_name), product_name)
        .Parameters.Append .CreateParameter("prod_category", 200, 1, Len(prod_category), prod_category)
        .Parameters.Append .CreateParameter("quantity_in_stock", 3, 1, , quantity_in_stock) ' 3 = adInteger
        .Parameters.Append .CreateParameter("price", 5, 1, , price) ' 5 = adDouble
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Product added successfully!", vbInformation, "Registration Complete"

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
    ' Populate product category combo box
    Combo1.AddItem "Reinforcement"
    Combo1.AddItem "Wood"
    Combo1.AddItem "Building Materials"
    Combo1.AddItem "Accessories"
    Combo1.AddItem "Finishing Materials"
End Sub

' Subroutine to clear all input fields
Private Sub ClearInputFields()
    txtProductID.Text = ""
    txtProductNames.Text = ""
    Quantity.Text = ""
    txtPrice.Text = ""
    Combo1.Text = ""
End Sub

Private Sub ExitBtn_Click()
    ' Close the form properly
    Unload Me
End Sub




Private Sub ViewBtn_Click()
    Dim conn As Object
    Dim rs As Object
    On Error GoTo ErrorHandler

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MIKE\Downloads\computer_MS\computer_MS\database\computer_ms_db_production.mdb;Persist Security Info=False;"
    conn.Open

    ' Query to retrieve data
    Dim sql As String
    sql = "SELECT product_id, product_name, prod_category, quantity_in_stock, price FROM tblProducts"

    ' Execute the query
    Set rs = conn.Execute(sql)

    ' Clear the PictureBox
    PictureBox1.Cls
    PictureBox1.Font.Size = 10
    PictureBox1.Font.Name = "Arial"

    ' Display column headers
    Dim header As String
    header = "Product ID" & vbTab & "Product Name" & vbTab & "Category" & vbTab & "Quantity" & vbTab & "Price"
    PictureBox1.Print header
    PictureBox1.Line (0, 30)-(PictureBox1.ScaleWidth, 30), vbBlack

    ' Display data rows
    Dim y As Single
    y = 45
    Do While Not rs.EOF
        Dim line As String
        line = rs("product_id") & vbTab & rs("product_name") & vbTab & rs("prod_category") & vbTab & rs("quantity_in_stock") & vbTab & rs("price")
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



Private Sub UpdateBtn_Click()

    ' Declare variables for the input fields
    Dim product_id As String, product_name As String, prod_category As String
    Dim quantity_in_stock As String, price As Double

    ' Retrieve input values from textboxes and combo box
    product_id = txtProductID.Text
    product_name = txtProductNames.Text
    prod_category = Combo1.Text ' Retrieve the category selection
    quantity_in_stock = Quantity.Text ' Convert to integer
    price = Val(txtPrice.Text) ' Convert to double

    ' Validate required fields
    If txtProductID.Text = "" Or txtProductNames.Text = "" Or Combo1.Text = "" Or Quantity.Text = "" Or txtPrice.Text = "" Then
        MsgBox "Please fill all fields before updating the product.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Validate numeric fields
    If quantity_in_stock <= 0 Then
        MsgBox "Quantity in stock must be greater than 0.", vbExclamation, "Invalid Quantity"
        Exit Sub
    End If

    If price <= 0 Then
        MsgBox "Price must be greater than 0.", vbExclamation, "Invalid Price"
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
        .CommandText = "UPDATE tblProducts SET product_name = ?, prod_category = ?, quantity_in_stock = ?, price = ? WHERE product_id = ?"
        ' Append parameters
        .Parameters.Append .CreateParameter("product_name", 200, 1, Len(product_name), product_name) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("prod_category", 200, 1, Len(prod_category), prod_category)
        .Parameters.Append .CreateParameter("quantity_in_stock", 3, 1, , quantity_in_stock) ' 3 = adInteger
        .Parameters.Append .CreateParameter("price", 5, 1, , price) ' 5 = adDouble
        .Parameters.Append .CreateParameter("product_id", 200, 1, Len(product_id), product_id)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Product updated successfully!", vbInformation, "Update Complete"

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


Private Sub DeleteBtn_Click()

    ' Declare a variable for the product ID
    Dim product_id As String

    ' Retrieve the product ID from the textbox
    product_id = txtProductID.Text

    ' Validate that the product ID field is not empty
    If txtProductID.Text = "" Then
        MsgBox "Please enter the Product ID to delete the product.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Confirm deletion with the user
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to delete this product?", vbQuestion + vbYesNo, "Confirm Deletion")
    If response = vbNo Then Exit Sub

    ' Database connection and command objects
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler ' Add error handling

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MIKE\Downloads\computer_MS\computer_MS\database\computer_ms_db_production.mdb;Persist Security Info=False;"
    conn.Open

    ' Prepare the SQL command for deleting the product
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "DELETE FROM tblProducts WHERE product_id = ?"
        ' Append parameter
        .Parameters.Append .CreateParameter("product_id", 200, 1, Len(product_id), product_id) ' 200 = adVarChar
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Product deleted successfully!", vbInformation, "Deletion Complete"

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

Private Sub SearchBtn_Click()

    ' Declare variables for the search input and record fields
    Dim product_id As String, product_name As String, prod_category As String
    Dim quantity_in_stock As Integer, price As Double
    Dim conn As Object
    Dim rs As Object

    ' Retrieve product ID to search from the text box
    product_id = txtSearch.Text

    ' Validate that a product ID is entered
    If product_id = "" Then
        MsgBox "Please enter a Product ID to search.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Database connection and recordset objects
    On Error GoTo ErrorHandler ' Add error handling

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MIKE\Downloads\computer_MS\computer_MS\database\computer_ms_db_production.mdb;Persist Security Info=False;"
    conn.Open

    ' Initialize recordset object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' SQL query to find the product by its ID
    rs.Open "SELECT * FROM tblProducts WHERE product_id = '" & product_id & "'", conn

    ' Check if any record is found
    If Not rs.EOF Then
        ' Retrieve and display the product details
        product_name = rs.Fields("product_name").Value
        prod_category = rs.Fields("prod_category").Value
        quantity_in_stock = rs.Fields("quantity_in_stock").Value
        price = rs.Fields("price").Value
        
        ' Populate the textboxes with the retrieved values
        txtProductNames.Text = product_name
        Combo1.Text = prod_category
        Quantity.Text = quantity_in_stock
        txtPrice.Text = price
    Else
        ' If no matching record is found, notify the user
        MsgBox "Product not found.", vbInformation, "Search Result"
    End If

    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    ' Handle any errors that occur
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub


Private Sub CommandBack_Click()
DashboardForm.Show

End Sub
