VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Convert Database Files to CSV or Tab Delimited text files"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   2400
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3495
      Left            =   0
      TabIndex        =   12
      Top             =   2640
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6165
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2400
         TabIndex        =   11
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   285
         Left            =   7080
         TabIndex        =   9
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   5295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   285
         Left            =   7080
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   5295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save As Tab Delimited"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save As Comma Delimited"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Select Table To Convert:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Save As:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Open Database:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sDatabase As String
Private sSaveas As String
Dim pr As Property ' setup the table properties object
Dim fl As Field ' setup the fields property

Private Sub Command1_Click()
' save as CSV
Dim sTemp As String ' temporary string to hold values
Dim fD As TableDef ' database
Dim sTempFields() As String ' temp array to hold the datbase fields
Dim x As Integer ' x as integer

On Error GoTo errHandler ' just incase we have an error happen

If Text1.Text = "" Then ' make sure text1 isn't empty
    MsgBox "You must select a database file to convert"
    Text1.SetFocus
    Exit Sub
End If
If Text2.Text = "" Then ' make sure text2 isn't empty
    MsgBox "You must enter a filename to save the file as!"
    Text2.SetFocus
    Exit Sub
End If
If Combo1.Text = "" Then ' make sure combo1 isn't empty
    MsgBox "You must select a table to convert the records from!"
    Text2.SetFocus
    Exit Sub
End If

Data1.RecordSource = "SELECT * FROM " & Combo1.Text ' setup the database record source
Data1.Refresh
Data1.UpdateControls

' check and make sure there are records in the database
If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
    MsgBox "There are no records in this table, Please select another table or add records first!"
    Exit Sub
End If

Data1.Recordset.MoveLast
ProgressBar1.Max = Data1.Recordset.RecordCount
Data1.Recordset.MoveFirst ' move to the first record in the datbase

x = 0

Set fD = Data1.Database.TableDefs(Combo1.Text)  ' setup the table object for the selected table

  For Each fl In fD.Fields ' now run through the list of fields and add them x
    x = x + 1 ' just setting up the value of x so we can setup the new array value
  Next
  
ReDim sTempFields(x) ' setup the array that will hold all the fields

x = 0
  For Each fl In fD.Fields ' now run through the list of fields and add them to the array
    sTempFields(x) = fl.Name
    x = x + 1
  Next
  
Dim sTF As String

Do While Data1.Recordset.EOF = False
    For x = 0 To UBound(sTempFields) - 1 ' loop through all the fields and add them to a temporary string
    
        If x = UBound(sTempFields) - 1 Then ' check to see if we are on the last field in the table, if so dont add the comma at the end
            ' add the record to the temp string with "around" and a comma at the end
            sTF = Chr(34) & Data1.Recordset.Fields(sTempFields(x)) & Chr(34) & Chr(13)
            Else
            ' not adding the comma at the end of the field value
            sTF = Chr(34) & Data1.Recordset.Fields(sTempFields(x)) & Chr(34) & Chr(44)
        End If
        
        sTemp = sTemp + sTF ' add the value of sTF to the value of sTemp
    Next x
    RichTextBox1.Text = sTemp ' insert the value of sTemp to the richtextbox
    ProgressBar1.Value = ProgressBar1.Value + 1
    Data1.Recordset.MoveNext ' move to next record in recordset
Loop ' continue loop

RichTextBox1.SaveFile Text2.Text, 1 ' save the text file after it is displayed to the user in the richtextbox

Exit Sub
errHandler:
If Err.Number = 94 Then ' found a null value so lets just add a empty space to the temp string
sTF = " "
Resume Next
End If

MsgBox Err.Number & vbCrLf & Err.Description ' found another error, notify user and exit sub
End Sub

Private Sub Command2_Click()
' Save as Tab Delimited
' all the code is the same in this sub as the comma delimited sub
' except we add a tab to the end of the fields value
Dim sTemp As String
Dim fD As TableDef
Dim sTempFields() As String
Dim x As Integer

On Error GoTo errHandler
If Text1.Text = "" Then
MsgBox "You must select a database file to convert"
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "You must enter a filename to save the file as!"
Text2.SetFocus
Exit Sub
End If
If Combo1.Text = "" Then
MsgBox "You must select a table to convert the records from!"
Text2.SetFocus
Exit Sub
End If
sComma = Chr(34) & Chr(44) & Chr(34)
Data1.RecordSource = "SELECT * FROM " & Combo1.Text
Data1.Refresh
Data1.UpdateControls

If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
MsgBox "There are no records in this table, Please select another table or add records first!"
Exit Sub
End If

Data1.Recordset.MoveLast
ProgressBar1.Max = Data1.Recordset.RecordCount
Data1.Recordset.MoveFirst

x = 0

Set fD = Data1.Database.TableDefs(Combo1.Text)  ' setup the table object for the selected table

  For Each fl In fD.Fields ' now run through the list of fields
    x = x + 1
  Next
  
ReDim sTempFields(x) ' setup the array that will hold all the fields

x = 0
  For Each fl In fD.Fields ' now run through the list of fields and add them to the array
    sTempFields(x) = fl.Name
    x = x + 1
  Next
Dim sTF As String

Do While Data1.Recordset.EOF = False
    For x = 0 To UBound(sTempFields) - 1
        If x = UBound(sTempFields) - 1 Then
            sTF = Chr(34) & Data1.Recordset.Fields(sTempFields(x)) & Chr(34) & Chr(13)
        Else
            ' add the tab value to the end of the string chr(9)
            sTF = Chr(34) & Data1.Recordset.Fields(sTempFields(x)) & Chr(34) & Chr(9)
        End If

        sTemp = sTemp + sTF
    Next x
    RichTextBox1.Text = sTemp
    ProgressBar1.Value = ProgressBar1.Value + 1
    Data1.Recordset.MoveNext
Loop
RichTextBox1.SaveFile Text2.Text, 1

Exit Sub
errHandler:
If Err.Number = 94 Then
sTF = " "
Resume Next
End If

MsgBox Err.Number & vbCrLf & Err.Description

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Dim td As TableDef
Screen.MousePointer = vbHourglass
CommonDialog1.CancelError = True
  On Error GoTo errHandler
  ' Set filters
  CommonDialog1.Filter = "Access Files (*.mdb)|*.mdb"
  ' Specify default filter
  CommonDialog1.FilterIndex = 2
  ' Display the Open dialog box
  CommonDialog1.ShowOpen
  ' Set database to name of selected file
  Data1.DatabaseName = CommonDialog1.FileName
  Text1.Text = CommonDialog1.FileName
  ' Get database file.
  If Data1.DatabaseName = "" Then
  Screen.MousePointer = vbNormal
  Exit Sub
  End If
  
  Data1.Refresh   ' Open the Database.
  ' Read and print the name of each table in the database to the combobox.
  For Each td In Data1.Database.TableDefs
      Combo1.AddItem td.Name
  Next

  Screen.MousePointer = vbNormal
  Exit Sub
  
errHandler:
  'User pressed the Cancel button
  Exit Sub

End Sub

Private Sub Command5_Click()
CommonDialog1.CancelError = True
  On Error GoTo errHandler
  ' Set filters
  CommonDialog1.Filter = "Comma Dilimited Text (*.csv)|*.csv|Tab Delimited Text (*.txt)|*.txt"
  ' Specify default filter
  CommonDialog1.FilterIndex = 1
  ' Display the Open dialog box
  CommonDialog1.FileName = "Default.csv"
  CommonDialog1.ShowSave
  Text2.Text = CommonDialog1.FileName
  Exit Sub
  
errHandler:
  'User pressed the Cancel button
  Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHandler
Data1.Database.Close
Exit Sub
errHandler:
If Err.Number = 91 Then Exit Sub
End Sub
