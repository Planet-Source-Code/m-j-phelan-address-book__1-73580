VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Address Book"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTxt 
      Caption         =   "Text Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   32
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   31
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   20
      Top             =   5160
      Width           =   7335
      Begin VB.OptionButton optZip 
         Caption         =   "Zip"
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optLName 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optCity 
         Caption         =   "City"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optState 
         Caption         =   "State"
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optFName 
         Caption         =   "First Name"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
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
         Left            =   6600
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   5280
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Check which field to search and enter the search criteria in the text box."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   6255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "SEARCH FOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   22
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "Reload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   19
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   7560
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   6360
      TabIndex        =   8
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   5160
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtLname 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtFname 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddEdit 
      Caption         =   "Add/Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwAddress 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3413
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Control ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Double click list item to Edit record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Zip"
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
      Left            =   7680
      TabIndex        =   16
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "State"
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
      Left            =   6480
      TabIndex        =   14
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "City"
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
      Left            =   5280
      TabIndex        =   13
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Address"
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
      Left            =   3840
      TabIndex        =   12
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblControlID 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
'SET UP

'Create a database with the following table
'table name  tblAdresses
'fields in the table
'ControlID  type - AutoNumber - make this field the Primary Key
'FirstName  type - Text
'LastName   type - Text
'Address    type - Text
'City       type - Text
'State      type - Text
'Zip        type - Text


'To use ListView in the menu "Project/Components" select
'"Microsoft Windows Common Controls 6.0 (SP6)" or whatever common controls you
'have available, i used the one listed, then add the list view to your form
'rename the listview to "lvwAddress"


'****************************************************************************
'****************************************************************************

'PROGRAM
Option Explicit

'Create a variable that will be global to be used as a connection
'Keep it global so you can connect once and remain connected until
'you terminate the program
'In the menu "Project/References", select "Microsoft ActiveX Data Objects 2.?"
'I used Microsoft ActiveX Data Objects 2.8 because that was the highest one
'I had available
Public cn As Connection

'Set the string name for the database path
'Where ever you create your program create a folder called "dbFiles"
'That is where you need to put the database called AddressBook.mdb
'The database may have a different extension under Access 2007 (may not be .mdb)

Public dbFile As String

'create a boolean variable (True or False) to be used with
'column sorting
Public blnColumn As Boolean

Private Sub cmdAddEdit_Click()
    'create a variable to hold the recordset
    Dim rs As Recordset
    'create a variable to hold the SQL statement
    Dim sql As String
    
    'create a variable to hold the control id from lblControlID
    
    Dim lngControlID As Long 'use long if the database has a huge amount of records; other wise use integer
    
    If lblControlID.Caption <> "" Then
        lngControlID = lblControlID.Caption
    End If
    
    'check to see that each text box has info in it, if not
    'end the sub
    If txtFname.Text = "" Then
        MsgBox "Fill in the First Name", vbOKOnly, "Text Validation"
        txtFname.SetFocus
        Exit Sub
    End If
    If txtLname.Text = "" Then
        MsgBox "Fill in the Last Name", vbOKOnly, "Text Validation"
        txtLname.SetFocus
        Exit Sub
    End If
    If txtAddress.Text = "" Then
        MsgBox "Fill in the Address", vbOKOnly, "Text Validation"
        txtAddress.SetFocus
        Exit Sub
    End If
    If txtCity.Text = "" Then
        MsgBox "Fill in the City", vbOKOnly, "Text Validation"
        txtCity.SetFocus
        Exit Sub
    End If
    If txtState.Text = "" Then
        MsgBox "Fill in the State", vbOKOnly, "Text Validation"
        txtState.SetFocus
        Exit Sub
    End If
    If txtZip.Text = "" Then
        MsgBox "Fill in the Zip", vbOKOnly, "Text Validation"
        txtZip.SetFocus
        Exit Sub
    End If
          
    'create the recordset
    Set rs = New Recordset
    
    'Google LockType and CursorType for more info on these
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    'this sql statement will tell which records to get from which table
    'in this case we select all fields and one record from the database
    'the recordset will hold only the record that has the unique ControlID
    'if the recordset is empty, we will add a new record
    sql = "Select * from tblAddresses where ControlID = " & lngControlID
    
    'this code with open the recordset and fill it with records
    'on a large database, the code may take a few seconds or more to fill
    'the recordset
    rs.Open sql, cn, , adLockOptimistic
    
    'if there are no records in the recordset end of file or EOF will be true
    'that is how we know to add a new record
    If rs.EOF Then
        rs.AddNew
        rs!firstname = txtFname.Text
        rs!lastname = txtLname.Text
        rs!Address = txtAddress.Text
        rs!city = txtCity.Text
        rs!State = txtState.Text
        rs!Zip = txtZip.Text
        rs.Update
    'else if not end of file or EOF
    'then we'll edit the existing record
    Else
        rs!firstname = txtFname.Text
        rs!lastname = txtLname.Text
        rs!Address = txtAddress.Text
        rs!city = txtCity.Text
        rs!State = txtState.Text
        rs!Zip = txtZip.Text
        rs.Update
    End If
    'close the recordset
    rs.Close
    
    'set the recordset to nothing to save memory
    Set rs = Nothing
    
    'reload the list view to add the updated or new record
    FillListView
    
    'clear the text boxes and label
    ClearLabelAndTexts
End Sub

Private Sub cmdExcel_Click()
    'create a variable to hold the recordset
    Dim rs As Recordset
    'create a variable to hold the SQL statement
    Dim sql As String
    
    'before this button will work you will need Microsoft Excel and you will
    'need to reference Excel under Project>References>Microsoft Excel #.# Object Library
    'in my case I referenced Microsoft Excel 12.0 Object Library
    'if you don't have that Library, you can reference whatever library you do have
    
    'this button will get ALL records from the database and write them
    'to an excel sheet, I created a blank sheet first and named it ExcelReport.xls
    
    'declare the object of Excel
    Dim xlApp As New Excel.Application
    
    'declare an object of Excel Workbook
    Dim xlwk As New Excel.Workbook
    
    'declare a variable act as a counter for a loop
    'could use integer unless there are a huge amount of records
    Dim i As Long
    
    'turn the mousepointer into an hourglass
    Screen.MousePointer = vbHourglass
    
     'create the recordset
    Set rs = New Recordset
    
    'Google LockType and CursorType for more info on these
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    'this is the sequal command to get records from the database
    sql = "Select * from tblAddresses"
    
    'this code with open the recordset and fill it with records
    'on a large database, the code may take a few seconds or more to fill
    'the recordset
    rs.Open sql, cn, , adLockOptimistic
    
    
    'set the excel application to interactive
    'which allows vb to write to it
    xlApp.Interactive = True
    
    'set the workbook to equal the open excel application and file
    Set xlwk = xlApp.Workbooks.Open(App.Path & "\DBFiles\ExcelReport.xls")
    
    'write your column titles before you write the records in excel
    xlApp.Range("A" & Trim(Str(1))).Value = "First Name"
    xlApp.Range("B" & Trim(Str(1))).Value = "Last Name"
    xlApp.Range("C" & Trim(Str(1))).Value = "Address"
    xlApp.Range("D" & Trim(Str(1))).Value = "City Name"
    xlApp.Range("E" & Trim(Str(1))).Value = "State"
    xlApp.Range("F" & Trim(Str(1))).Value = "Zip"
    
    'set i to 2 - which will be the first row in Excel
    'that will hold your records
    i = 2
    
    'loop through the records and write each one
    'to excel
    Do Until rs.EOF
        'note i starts at 2.  that is the row the first
        'record will start on
        'in the Excel app, ColumnA, row(i) the first name will be written
        xlApp.Range("A" & Trim(Str(i))).Value = rs!firstname
        'in the Excel app, ColumnB, row(i) the last name will be written, etc, etc
        xlApp.Range("B" & Trim(Str(i))).Value = rs!lastname
        xlApp.Range("C" & Trim(Str(i))).Value = rs!Address
        xlApp.Range("D" & Trim(Str(i))).Value = rs!city
        xlApp.Range("E" & Trim(Str(i))).Value = rs!State
        xlApp.Range("f" & Trim(Str(i))).Value = rs!Zip
        'increment i by 1 each time the code reaches here
        'the i will designate which row we are on in Excel
        i = i + 1
        'move to the next record in the recordset
        rs.MoveNext
    Loop
    
    'make the report visible
    xlApp.Visible = True
    
    'turn the mouse pointer back to its normal state
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdExit_Click()
    'unload the form and end program
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'create a variable to hold the recordset
    Dim rs As Recordset
    'create a variable to hold the SQL statement
    Dim sql As String
    
    'use long if the database has a huge amount of records; other wise use integer
    Dim i As Long
    
    'create a variable to hold the search text
    Dim strSearch As String

    'create a variable to hold the ControlID selected
    'from the list view
    Dim lngControlID As Long
    
    'if there is no record in the text boxes, no record has been selected
    'to be deleted, exit sub
    If txtSearch.Text = "" Then
        Exit Sub
    End If
    

    'create the recordset
    Set rs = New Recordset
    
    'Google LockType and CursorType for more info on these
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    'the sql statement will tell which record to get from which table
    'in this case we select all fields and one record from the database
    'depending on which check box is selected by the user
    
    'if there is no search text, give a message to enter it
    'and exit sub
    If txtSearch.Text = "" Then
        MsgBox "Enter search criteria.", vbOKOnly, "Search Criteria Validation"
        txtSearch.SetFocus
        Exit Sub
    End If
    
    'set the search variable to the value in the search text box
    strSearch = txtSearch.Text
    
    'The sql statement will be different depending on which option is chosen
    'for the search criteria
    If optFName.Value = True Then
        sql = "Select * from tblAddresses where FirstName = '" & strSearch & "'"
    End If
    
    If optLName.Value = True Then
        sql = "Select * from tblAddresses where LastName = '" & strSearch & "'"
    End If
    
    If optAddress.Value = True Then
        sql = "Select * from tblAddresses where Address = '" & strSearch & "'"
    End If
    
    If optCity.Value = True Then
        sql = "Select * from tblAddresses where City = '" & strSearch & "'"
    End If
    
    If optState.Value = True Then
        sql = "Select * from tblAddresses where State = '" & strSearch & "'"
    End If
    
    If optZip.Value = True Then
        sql = "Select * from tblAddresses where Zip = '" & strSearch & "'"
    End If
    
    'this code with open the recordset and fill it with records
    'on a large database, the code may take a few seconds or more to fill
    'the recordset
    rs.Open sql, cn, , adLockOptimistic
    
    'clear the listview
    lvwAddress.ListItems.Clear
    
    'populate the list view with the record(s) found from the search
    With lvwAddress
        'now loop through all of the records until EOF or "end of file"
        Do Until rs.EOF
            
            '"i" will be used to point to the row number for the items being added
            'to the listview
            'increment the i each time it loops past it
            i = i + 1
            'now write the records from the recordset to the listview fields
            .ListItems.Add , , rs!ControlID 'control id
            
            'you can short cut this, but it is important not to do so
            'check for blank fields, otherwise you'll get errors when there is a field
            'that is blank (for example, if there is no address)
            If Not rs!firstname = "" Then .ListItems(i).ListSubItems.Add , , rs!firstname Else .ListItems(i).ListSubItems.Add , , ""
            
            If Not rs!lastname = "" Then .ListItems(i).ListSubItems.Add , , rs!lastname Else .ListItems(i).ListSubItems.Add , , ""
            
            If Not rs!Address = "" Then .ListItems(i).ListSubItems.Add , , rs!Address Else .ListItems(i).ListSubItems.Add , , ""
            If Not rs!city = "" Then .ListItems(i).ListSubItems.Add , , rs!city Else .ListItems(i).ListSubItems.Add , , ""
            If Not rs!State = "" Then .ListItems(i).ListSubItems.Add , , rs!State Else .ListItems(i).ListSubItems.Add , , ""
            If Not rs!Zip = "" Then .ListItems(i).ListSubItems.Add , , rs!Zip Else .ListItems(i).ListSubItems.Add , , ""
            
            rs.MoveNext
        Loop
    End With
    'close the recordset
    rs.Close
    
    'kill the recordset to save memory
    Set rs = Nothing
    
    
End Sub

Private Sub cmdDelete_Click()
    'create a variable to hold the recordset
    Dim rs As Recordset
    'create a variable to hold the SQL statement
    Dim sql As String

    'create a variable to hold the ControlID selected
    'from the list view
    Dim lngControlID As Long
    
    'if there is no record in the text boxes, no record has been selected
    'to be deleted, exit sub
    If lblControlID.Caption = "" Then
        Exit Sub
    End If
    
    'assign the selected item's to the variable
    lngControlID = lblControlID.Caption

    'create the recordset
    Set rs = New Recordset
    
    'Google LockType and CursorType for more info on these
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    'this sql statement will tell which records to get from which table
    'in this case we select all fields and one record from the database
    'the recordset will hold only the record that has the unique ControlID
    sql = "Select * from tblAddresses where ControlID = " & lngControlID
    
    'this code with open the recordset and fill it with records
    'on a large database, the code may take a few seconds or more to fill
    'the recordset
    rs.Open sql, cn, , adLockOptimistic
    
    'if the file is not eof of file or EOF then the record was found
    'delete the record
    If Not rs.EOF Then
        rs.Delete
    End If
    'close the recordset
    rs.Close
    
    'set the recordset to nothing to save memory
    Set rs = Nothing
    
    'reload the list view to add the updated or new record
    FillListView
    
    'clear the text boxes and label
    ClearLabelAndTexts
End Sub

Private Sub cmdTxt_Click()
    'create a variable to hold the recordset
    Dim rs As Recordset
    'create a variable to hold the SQL statement
    Dim sql As String
    'create a variable to hold the freefile number
    Dim FN As Long
    'create a variable to hold the path of the file
    Dim strPath As String
    'create a variable to hold the path and application name of Notepad.exe
    Dim nPath As String
    'create a variable to hold the path and file name of the new report
    Dim dTaskID As String
    'create a variable to hold the file name
    Dim strFile As String
    
    'create a variable to use as a counter
    Dim i As Integer
    
    'create a variable to use with the for loop
    Dim ii As Integer
    
    'create a variable to use as the body of the email
    'stringing all the strings together into this one
    Dim strBody As String
    
    ''create a variables to hold the items from the database
    Dim strFName As String
    Dim strLName As String
    Dim strAddress As String
    Dim strCity As String
    Dim strState As String
    Dim strZip As String
    
    'create a variable to hold the name, address, city, etc
    Dim strRow As String
    
    'set the strPath to equal the application path
    strPath = App.Path
    'add a "\" to the end of the strPath
    If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    'add the report name to strPath
    strPath = strPath & "TextReport"
    'add the report extension ".txt" to strPath
    strFile = strPath & ".txt"
    
    
    'create the recordset
    Set rs = New Recordset
    
    'Google LockType and CursorType for more info on these
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    'this sql statement will tell which records to get from which table
    'in this case we select all fields and one record from the database
    'the recordset will hold only the record that has the unique ControlID
    sql = "Select * from tblAddresses"
    
    'this code with open the recordset and fill it with records
    'on a large database, the code may take a few seconds or more to fill
    'the recordset
    rs.Open sql, cn, , adLockOptimistic
    
    'assing a free file number to FN
    FN = FreeFile
    'open the strFile as an output file with the free file number
    Open strFile For Output As FN
    
    'print the first line to the report
    Print #FN, "REPORT DATE/TIME " & Format(Now(), "MM/DD/YYYY HH:MM AMPM")
    
    'write a blank line
    Print #FN, ""
    
    'create the report column names
    Print #FN, "First Name   " & " " & "Last Name     " & " " & "    Address             " & " " & "City" & ",       " & " " & "State      " & " " & "Zip"
    
    'loop through the records from the recordset and print each one
    Do Until rs.EOF
        'assign the information to variables, then make each variable the
        'same length, so you can have even looking columns in the printed file
        strFName = rs!firstname
        strLName = rs!lastname
        strAddress = rs!Address
        strCity = rs!city
        strState = rs!State
        strZip = rs!Zip
        
        'get the lenght of strFname
        i = Len(strFName)
        'if the length is not equal to 13 (i picked 13 at random, you can adjust this
        'length by increasing or decreasing the number to fit your report better)
        'then make it the length to 13 by adding blank spaces
        If i < 13 Then
            ii = 13 - i
            For i = 1 To ii
                'this will put a blank space on the end of strFname
                'until it is 13 characters long
                strFName = strFName & " "
            Next i
        End If
        
        'repeat the process for each variable from the database
        'get the lenght of strLname
        i = Len(strLName)
        If i < 15 Then
            ii = 15 - i
            For i = 1 To ii
                'this will put a blank space on the end of strFname
                'until it is 15 characters long
                strLName = strLName & " "
            Next i
        End If
        
        i = Len(strAddress)
        If i < 25 Then
            ii = 25 - i
            For i = 1 To ii
                'this will put a blank space on the end of strFname
                'until it is 20 characters long
                strAddress = strAddress & " "
            Next i
        End If
        
        i = Len(strCity)
        If i < 12 Then
            ii = 12 - i
            For i = 1 To ii
                'this will put a blank space on the end of strFname
                'until it is 12 characters long
                strCity = strCity & " "
            Next i
        End If
        
        i = Len(strState)
        If i < 12 Then
            ii = 12 - i
            For i = 1 To ii
                'this will put a blank space on the end of strFname
                'until it is 12 characters long
                strState = strState & " "
            Next i
        End If
        
        'since Zip is the last column, there is no need to make it any longer
        
        'make the body by putting all of the variables together
        strBody = strFName & strLName & strAddress & strCity & strState & strZip
        
        
        'print the boyd
        Print #FN, strBody
        rs.MoveNext
    Loop
    'close the recordset
    rs.Close
    
    'set the recordset to nothing to save memory
    Set rs = Nothing
    
    'close the file so you can open it for viewing
    Close FN
    
    
    'open notepad
    nPath = "C:\WINDOWS\notepad.exe"
    'show the new report
    dTaskID = Shell(nPath + " " + strFile, vbNormalFocus)
    
End Sub

Private Sub lvwAddress_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    'this sub will allow the columns in the listview to
    'be sorted ascending or descending if that column is clicked
    'one quirk is that it will sort alpabetically, you'll notice
    'numbers will sort alphabetically instead of in numeric sequence

    lvwAddress.SortKey = ColumnHeader.Index - 1
    lvwAddress.Sorted = True
    
    
    If blnColumn Then
        lvwAddress.SortOrder = lvwAscending
        blnColumn = False
    Else
        lvwAddress.SortOrder = lvwDescending
        blnColumn = True
    End If
    lvwAddress.Sorted = False

End Sub
Private Sub cmdReload_Click()
    'call FillListView to reload the list
    FillListView
    
    'clear the label and text boxes
    ClearLabelAndTexts
    
End Sub
Private Sub ClearLabelAndTexts()
    'clear the label
    lblControlID.Caption = ""
    
    'clear all of the text boxes
    txtFname.Text = ""
    txtLname.Text = ""
    txtAddress.Text = ""
    txtCity.Text = ""
    txtState.Text = ""
    txtZip.Text = ""
    
End Sub
Private Sub Form_Load()
    
    'get the data base path
    'this line says use the same path as the application app.path
    dbFile = App.Path
    
    'now add the folder we created called "dbFiles" to the path
    'so we can point to the database
    If Right$(dbFile, 1) <> "\" Then
        dbFile = dbFile & "\"
    End If
    
    'now add the database name to the end of the path
    dbFile = dbFile & "dbFiles\AddressBook.mdb"
    Debug.Print dbFile

    'Open the database connection
    Set cn = New Connection
    
    'Set the database type
    'use this string for any Access database version lower than
    'Access 2007
    'cn.Provider = "Microsoft.Jet.OLEDB.4.0" 'for Access below version 2007
    
    'use this string for Access database Access 2007
    cn.Provider = "Microsoft.ACE.OLEDB.12.0" 'for Access 2007
    
    'set the connection for read/write access
    cn.Mode = adModeReadWrite
    
    'if you database is password protected, add this code
    'cnAdmin.Properties("Jet OLEDB:Database Password") = "yourPasswordHere"  'for Access below version 2007
    
    'I have not tested this since I don't have Access 2007 but assume
    'the password code for the ACE.OLEDB should be
    'cnAdmin.Properties("ACE OLEDB:Database Password") = "yourPasswordHere"
    
    'Open the database
    cn.Open (dbFile)
    
    'call the sub to create the colum headers
    SetListViewColumns
     
    'call the sub to fill the list view with data from your database
    FillListView
    
End Sub
Private Sub SetListViewColumns()
    'loads column headers and sets the labels that capture the employee
    'and exception info to false
    lvwAddress.ColumnHeaders.Clear
    
    lvwAddress.Font.Size = 10

    'there are different types of Views, I've only used "lvwReport" view
    lvwAddress.View = lvwReport
    
    'adds column headers to the list view in the exception log
    With lvwAddress.ColumnHeaders
        .Add , , "ID", 1 'this colum will be very narrow and not visible, but is needed for the lookup
        .Add , , "F Name", 1200 'you can change the numbers here to make the colum wider or narrower
        .Add , , "L Name", 1200
        .Add , , "Address", 2500
        .Add , , "City", 1400
        .Add , , "State", 1500
        .Add , , "Zip", 800
    End With
    'the 1 here means "True" - this will make gridlines appear, set it to 0 if you don't want them
    lvwAddress.Gridlines = 1
    
    '1 and 0 here are True and False, this will allow you to select he whole row
    lvwAddress.FullRowSelect = 1
End Sub
Private Sub FillListView()
    'create a variable to hold the recordset
    Dim rs As Recordset
    'create a variable to hold the SQL statement
    Dim sql As String
    
    Dim i As Long 'use long if the database has a huge amount of records; other wise use integer
    
    
    'clear the listview
    lvwAddress.ListItems.Clear
    
    'create the recordset
    Set rs = New Recordset
    
    'Google LockType and CursorType for more info on these
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    'this sql statement will tell which records to get from which table
    'in this case we select all fields and all records from the database
    'the recordset will hold all of these records
    sql = "Select * from tblAddresses"
    
    'this code with open the recordset and fill it with records
    'on a large database, the code may take a few seconds or more to fill
    'the recordset
    rs.Open sql, cn, , adLockOptimistic
    
    i = 0
    'using the with statement allows me not to have to retype lvwAddress
    'in front of all of the .ListItems.Add or any functions the listview provides
    'simply add a "." after the "With lvwAddress" and you'll see many options in a drop down list
    'when I'm done using it, I have to "End With" as seen below
    With lvwAddress
        'now loop through all of the records until EOF or "end of file"
        Do Until rs.EOF
            
            '"i" will be used to point to the row number for the items being added
            'to the listview
            'increment the i each time it loops past it
            i = i + 1
            'now write the records from the recordset to the listview fields
            .ListItems.Add , , rs!ControlID 'control id
            
            'you can short cut this, but it is important not to do so
            'check for blank fields, otherwise you'll get errors when there is a field
            'that is blank (for example, if there is no address)
            If Not rs!firstname = "" Then .ListItems(i).ListSubItems.Add , , rs!firstname Else .ListItems(i).ListSubItems.Add , , ""
            
            If Not rs!lastname = "" Then .ListItems(i).ListSubItems.Add , , rs!lastname Else .ListItems(i).ListSubItems.Add , , ""
            
            If Not rs!Address = "" Then .ListItems(i).ListSubItems.Add , , rs!Address Else .ListItems(i).ListSubItems.Add , , ""
            If Not rs!city = "" Then .ListItems(i).ListSubItems.Add , , rs!city Else .ListItems(i).ListSubItems.Add , , ""
            If Not rs!State = "" Then .ListItems(i).ListSubItems.Add , , rs!State Else .ListItems(i).ListSubItems.Add , , ""
            If Not rs!Zip = "" Then .ListItems(i).ListSubItems.Add , , rs!Zip Else .ListItems(i).ListSubItems.Add , , ""
            
            rs.MoveNext
        Loop
    End With
    'close the recordset
    rs.Close
    
    'kill the recordset to save memory
    Set rs = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'close the connection
    cn.Close
    
    'kill the connection to save memory
    Set cn = Nothing
End Sub

Private Sub lvwAddress_DblClick()
    'this sub will take the double clicked record and add them to the
    'text boxes and remove it from the listview
    'once the record is in the text boxes, you can edit or delete it
    
    'create a variable to hold the record list item index number
    Dim intIndex As Integer
    
    'assign the selected item's index number to the variable
    intIndex = lvwAddress.SelectedItem.Index
    
    With lvwAddress
        'the first column on the listview will be the SelectedItem
        'remember, this column is hidden since we created it as
        '.Add , , "ID", 1  in "Private Sub SetListViewColumns"
        lblControlID.Caption = .SelectedItem
        'any subsequent items after .SelectedItem will be .SelectedItem.SubItems(#)
        txtFname.Text = .SelectedItem.SubItems(1)
        txtLname.Text = .SelectedItem.SubItems(2)
        txtAddress.Text = .SelectedItem.SubItems(3)
        txtCity.Text = .SelectedItem.SubItems(4)
        txtState.Text = .SelectedItem.SubItems(5)
        txtZip.Text = .SelectedItem.SubItems(6)
    End With
    
    'remove the selected item from the listview
    'we'll put it back if updated - see "Private Sub cmdAddEdit_Click"
    lvwAddress.ListItems.Remove intIndex
    
End Sub
