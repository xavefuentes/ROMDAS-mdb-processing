VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ROMDAS mdb Processing Software"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11835
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBatchExport 
      Caption         =   "&Batch Export"
      Height          =   375
      Left            =   9840
      TabIndex        =   12
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuickExport 
      Caption         =   "&Quick Export"
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   7320
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   100
      Left            =   120
      TabIndex        =   10
      Top             =   8400
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdExportCSV 
      Caption         =   "&Export to CSV"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelSID 
      Caption         =   "&Del SURVEY_ID"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddSID 
      Caption         =   "&Add SURVEY_ID"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Records"
      Height          =   6495
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   7695
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6135
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   10821
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9495
   End
   Begin VB.CommandButton cmdLoadmdbFile 
      Caption         =   "L&oad MDB File"
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9840
      Top             =   -360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tables"
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3855
      Begin VB.ListBox List1 
         Height          =   6105
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   10705
      Picture         =   "Main.frx":08CA
      Top             =   7880
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   8040
      Width           =   9495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'declare variables
Public GPSDB, surveyid As String
Public GPSRec, IRIRec As Integer
Public con As ADODB.Connection
Public recs As ADODB.Recordset


'send mail to xfuentes@gmail.com
Private Const IDC_HAND = 32649&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5


'open database connection
Public Sub open_mdb()
    Set con = New ADODB.Connection
    Set recs = New ADODB.Recordset
    With con
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & GPSDB & ";Persist Security Info=False"
    End With
    
    With recs
        .CursorLocation = adUseClient
        .ActiveConnection = con
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
End Sub


'close database connection
Public Sub close_mdb()
    Set con = Nothing
End Sub


'display tables
Private Sub display_tbl()
table:
    Dim x As Integer
    x = 0
    List1.Clear
    Call open_mdb
    Set recs = con.OpenSchema(adSchemaTables)
    With recs
    Do While Not .EOF
        x = x + 1
        Dim y As String
        y = recs("TABLE_NAME")
        If Not (y Like "MSys*") Then
        'filter tables with MSys and F_D names
        'If Not (y Like "MSys*" Or y Like "F_D*") Then
            List1.AddItem y
        End If
        'display all tables
        'List1.AddItem recs("TABLE_NAME")
        .MoveNext
    Loop
    End With
    Call close_mdb
End Sub


'refresh datagrid on click
Private Sub refresh_dgrid()
    Call open_mdb
    With recs
        .Open "Select * from " & "[" & List1.Text & "]" & ""
            If .RecordCount <> 0 Then
                Set DataGrid1.DataSource = recs
            Else
                Set DataGrid1.DataSource = Nothing
                'replaced with autoclose
                'MsgBox "No records to display.", vbInformation, "Message"
                Call ACmsgbox(3, "No records to display.", vbInformation, "Message")
            End If
    End With
    Call close_mdb
End Sub


'refresh datagrid and focus on gps_processed
Private Sub refresh_gps()
    Call open_mdb
    With recs
        .Open "Select * from " & "[GPS_Processed_" & surveyid & "]"
        List1.Text = "GPS_Processed_" & surveyid
        Set DataGrid1.DataSource = recs
    End With
    Call close_mdb
End Sub


'append survey_id to gps_processed and plaser_iri
Private Sub append_sid()
    cmdAddSID.Enabled = False
    cmdDelSID.Enabled = False
    cmdExportCSV.Enabled = False
    cmdLoadmdbFile.Enabled = False

    Label2.Caption = ""
    ProgressBar1.Max = GPSRec
    ProgressBar1.Value = 0
    Call open_mdb
    With recs
        .Open "Select * from " & "GPS_Processed_" & surveyid
        .MoveFirst
        While Not .EOF
            .Fields("SURVEY_ID") = surveyid
            Label2.Caption = "Adding SURVEY_ID to " & "GPS_Processed_" & surveyid & ": " & Fix((ProgressBar1.Value / ProgressBar1.Max) * 100) & "%" & " completed."
            Label2.Refresh
            ProgressBar1.Value = ProgressBar1.Value + 1
            .MoveNext
            DoEvents
        Wend
        Label2.Caption = "GPS_Processed_" & surveyid & " finished."
        List1.ListIndex = 1
        Set DataGrid1.DataSource = recs
    End With
    Call close_mdb
        
'appending survey_id to plaser_iri
    Label2.Caption = ""
    ProgressBar1.Max = IRIRec
    ProgressBar1.Value = 0
    Call open_mdb
    With recs
        .Open "Select * from " & "PLaser_IRI_" & surveyid
        .MoveFirst
        While Not .EOF
            .Fields("SURVEY_ID") = surveyid
            Label2.Caption = "Adding SURVEY_ID to " & "PLaser_IRI_" & surveyid & ": " & Fix((ProgressBar1.Value / ProgressBar1.Max) * 100) & "%" & " completed."
            Label2.Refresh
            ProgressBar1.Value = ProgressBar1.Value + 1
            .MoveNext
            DoEvents
        Wend
        Label2.Caption = "PLaser_IRI_" & surveyid & " finished."
    End With
    Call close_mdb
    ProgressBar1.Value = 0
    Label2.Caption = "SURVEY_ID fields added."
    'replaced with autoclose
    'MsgBox "SURVEY_ID fields added.", vbInformation, "Message"
    Call ACmsgbox(3, "SURVEY_ID fields added.", vbInformation, "Message")
    cmdAddSID.Enabled = True
    cmdDelSID.Enabled = True
    cmdExportCSV.Enabled = True
    cmdLoadmdbFile.Enabled = True
End Sub


'append survey_id to gps_processed and plaser_iri(2)
Private Sub append_sid2()
    Dim sql As String
    Call open_mdb
    With con
        sql = "UPDATE [GPS_Processed_" & surveyid & "]" & " SET [GPS_Processed_" & surveyid & "]" & ".SURVEY_ID=" & "'" & surveyid & "'" & ""
        'testing the update string
        'MsgBox sql, vbInformation, "Message"
        .Execute sql
        
        'working update query
        '.Execute "UPDATE GPS_Processed_01_07A_41B_CL SET GPS_Processed_01_07A_41B_CL.SURVEY_ID = '01_07A_41B_CL'"
    End With
    Call close_mdb
        
'appending survey_id to plaser_iri
    Call open_mdb
    With con
        sql = "UPDATE [PLaser_IRI_" & surveyid & "]" & " SET [PLaser_IRI_" & surveyid & "]" & ".SURVEY_ID=" & "'" & surveyid & "'" & ""
        'testing the update string
        'MsgBox sql, vbInformation, "Message"
        .Execute sql
    End With
    Call close_mdb
    
    Label2.Caption = "SURVEY_ID fields added."
    'replaced with autoclose
    'MsgBox "SURVEY_ID fields added.", vbInformation, "Message"
    Call ACmsgbox(3, "SURVEY_ID fields added.", vbInformation, "Message")
End Sub

'add survey_id field in gps_processed and plaser_iri
Private Sub cmdAddSID_Click()
    On Error GoTo Error
    'removing code or update query
    'Dim x As Integer
    'x = MsgBox("Adding the SURVEY_ID fields will take time. Please wait for the append process confirmation to appear. Do you want to continue?", vbYesNo, "Message")
    'If x = 6 Then
        Call open_mdb
        With con
            .Execute "Alter table " & "[GPS_Processed_" & surveyid & "]" & "  add SURVEY_ID Text"
            .Execute "Alter table " & "[PLaser_IRI_" & surveyid & "]" & " add SURVEY_ID Text"
        End With
        Call close_mdb
    'Else
    '    Label2.Caption = "Add SURVEY_ID fields canceled."
    '    GoTo Term
    'End If
    
    GoTo FieldAdded

'error handler
Error:
    Label2.Caption = "Add SURVEY_ID fields canceled."
    'replaced with autoclose
    'MsgBox "Unable to add SURVEY_ID. The fields already exist in the database.", vbCritical, "Message"
    Call ACmsgbox(3, "Unable to add SURVEY_ID. The fields already exist in the database.", vbCritical, "Message")
    GoTo Term

'appending survey_id to gps_processed
FieldAdded:
    Call append_sid2
    Call refresh_gps
    
Term:
End Sub


'delete survey_id fied in gps_processes and plaser_iri tables
Private Sub cmdDelSID_Click()
    On Error GoTo Error
    Dim x As Integer
    'replaced with autoclose
    'x = MsgBox("Are you sure you want to delete the SURVEY_ID fields?", vbYesNo, "Message")
    x = ACmsgbox(3, "Are you sure you want to delete the SURVEY_ID fields?", vbYesNo, "Message")
    If x = 6 Then
        Call open_mdb
        With con
            .Execute "Alter table " & "[GPS_Processed_" & surveyid & "]" & "  drop SURVEY_ID Text"
            .Execute "Alter table " & "[PLaser_IRI_" & surveyid & "]" & " drop SURVEY_ID Text"
        End With
        Call close_mdb
    Else
        Label2.Caption = "Delete SURVEY_ID fields canceled."
        GoTo Term
    End If
    GoTo FieldDeleted
    
Error:
    Label2.Caption = "Delete SURVEY_ID fields canceled."
    'replaced with autoclose
    'MsgBox "There are no SURVEY_ID fields to delete.", vbInformation, "Message"
    Call ACmsgbox(3, "There are no SURVEY_ID fields to delete.", vbInformation, "Message")
    GoTo Term
    
FieldDeleted:
    Call refresh_gps
    Label2.Caption = "SURVEY_ID fields deleted."
    'replaced with autoclose
    'MsgBox "SURVEY_ID fields deleted.", vbInformation, "Message"
    Call ACmsgbox(3, "SURVEY_ID fields deleted.", vbInformation, "Message")
    
Term:
End Sub


'select mdb file to process
Private Sub cmdLoadmdbFile_Click()
    On Error GoTo Error
    'filter to select only mdb files
    CommonDialog1.Filter = "mdb Files | *.mdb"
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen

    GPSDB = CommonDialog1.FileName
    txtFileName.Text = GPSDB
    
    Call display_tbl

    Call open_mdb
    With recs
        .Open ("Select * from Survey_Header")
            If .RecordCount <> 0 Then
                Set DataGrid1.DataSource = recs
                'table index varies, commented for future update
                'List1.ListIndex = 12
                List1.Text = "Survey_Header"
                surveyid = recs!SURVEY_ID
                Label2.Caption = "SURVEY_ID " & surveyid & " loaded."
                Label2.Refresh
            End If
    End With
    Call close_mdb
    cmdAddSID.Enabled = True
    cmdDelSID.Enabled = True
    cmdExportCSV.Enabled = True
    
    'find the records counts for gps_processed and plaser_iri
    Call open_mdb
    With recs
        .Open "Select * from " & "[GPS_Processed_" & surveyid & "]"
        GPSRec = .RecordCount
        'MsgBox GPSRec, vbInformation, "Message"
    End With
    Call close_mdb
    
    Call open_mdb
    With recs
        .Open "Select * from " & "[PLaser_IRI_" & surveyid & "]"
        IRIRec = .RecordCount
        'MsgBox IRIRec, vbInformation, "Message"
    End With
    Call close_mdb

    GoTo Term

Error:
    'replaced with autoclose
    'MsgBox "No mdb file was selected.", vbInformation, "Message"
    Call ACmsgbox(3, "No mdb file was selected.", vbInformation, "Message")
    
Term:

End Sub


'export function gps_processed and plaser_iri tables to csv
Private Function DBExport() As Long
    On Error Resume Next
    'exporting gps_processed
    Dim x, y As String
    x = "GPS_Processed_" & surveyid
    
    'delete invalid characters in arcgis
    y = x
    
    y = Replace$(y, "Processed_", "")
    y = Replace$(y, "-", "_")
    y = Replace$(y, " ", "_")
    y = Replace$(y, ".", "_")
    y = Replace$(y, ",", "_")
    y = Replace$(y, "&", "_")
    y = Replace$(y, "(", "_")
    y = Replace$(y, ")", "")
    
    'delete extra underscores
    y = Replace$(y, "_____", "_")
    y = Replace$(y, "____", "_")
    y = Replace$(y, "___", "_")
    y = Replace$(y, "__", "_")
    
    y = y & ".csv"
    Kill App.Path & "\" & y
    
    Call open_mdb
    With con
        .Execute "SELECT * INTO [Text;Database=" & App.Path & ";HDR=Yes;FMT=Delimited].[" & y & "] FROM [" & x & "]", DBExport, adCmdText Or adExecuteNoRecords
    End With
    Call close_mdb
    Kill App.Path & "\schema.ini"
    
    'exporting plaser_iri
    Dim z As Integer
    'replaced with autoclose
    'z = MsgBox("Do you want to export PLaser_IRI_" & surveyid, vbYesNo, "Message")
    z = ACmsgbox(3, "Do you want to export PLaser_IRI_" & surveyid, vbYesNo, "Message")
    If z = 6 Then
        Dim a, b As String
        a = "PLaser_IRI_" & surveyid
        
        'delete invalid characters in arcgis
        b = a
        
        b = Replace$(b, "PLaser_", "")
        b = Replace$(b, "-", "_")
        b = Replace$(b, " ", "_")
        b = Replace$(b, ".", "_")
        b = Replace$(b, ",", "_")
        b = Replace$(b, "&", "_")
        b = Replace$(b, "(", "_")
        b = Replace$(b, ")", "")
    
        'delete extra underscores
        b = Replace$(b, "_____", "_")
        b = Replace$(b, "____", "_")
        b = Replace$(b, "___", "_")
        b = Replace$(b, "__", "_")
        
        b = b & ".csv"
        Kill App.Path & "\" & b
        
        Call open_mdb
        With con
            .Execute "SELECT * INTO [Text;Database=" & App.Path & ";HDR=Yes;FMT=Delimited].[" & b & "] FROM [" & a & "]", DBExport, adCmdText Or adExecuteNoRecords
        End With
        Call close_mdb
        Kill App.Path & "\schema.ini"
    Else
        GoTo Term
    End If

Term:
End Function


'export gps_processed and plaser_iri tables to csv
Private Sub cmdExportCSV_Click()
    On Error GoTo Error
    Call open_mdb
    With recs
        .Open "Select SURVEY_ID from " & "[GPS_Processed_" & surveyid & "]"
        If .RecordCount <> 0 Then
                Label2.Caption = CStr(DBExport()) & " records exported."
        End If
    End With
    Call close_mdb
    GoTo Term

Error:
    Label2.Caption = "SURVEY_ID fields missing. Export to CSV aborted."
    'replaced with autoclose
    'MsgBox "SURVEY_ID fields missing. Exporting to CSV aborted.", vbCritical, "Message"
    Call ACmsgbox(3, "SURVEY_ID fields missing. Exporting to CSV aborted.", vbCritical, "Message")

Term:
    
End Sub


'for lazy people who wants everything in one click
Private Sub cmdQuickExport_Click()
    On Error GoTo Error
    
    CommonDialog1.Filter = "mdb Files | *.mdb"
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen

    GPSDB = CommonDialog1.FileName
    txtFileName.Text = GPSDB
    
    Call display_tbl

    Call open_mdb
    With recs
        .Open ("Select * from Survey_Header")
            If .RecordCount <> 0 Then
                Set DataGrid1.DataSource = recs
                'table index varies, commented for future update
                'List1.ListIndex = 12
                List1.Text = "Survey_Header"
                surveyid = recs!SURVEY_ID
                Label2.Caption = "SURVEY_ID " & surveyid & " loaded."
                Label2.Refresh
            End If
    End With

    Call cmdAddSID_Click
    Call cmdExportCSV_Click
    Label2.Caption = "Export to CSV completed successfully."
    'replaced with autoclose
    'MsgBox "Quick Export to CSV completed successfully.", vbInformation, "Message"
    Call ACmsgbox(3, "Quick Export to CSV completed successfully.", vbInformation, "Message")
    GoTo Term

Error:
    Call ACmsgbox(3, "No mdb file was selected.", vbInformation, "Message")
    GoTo Term

Term:

End Sub


'batch process folder
Private Sub cmdBatchExport_Click()
    On Error Resume Next
    'count the mdb files in the current directory
    Dim e As String
    Dim f As Integer
    f = 0
    e = Dir(CurDir() & "\" & "*.mdb")
    Do While e <> ""
    f = f + 1
    e = Dir()
    Loop
    'display the number of mdb files
    'Call ACmsgbox(3, f & " mdb of files in current directory.", vbInformation, "Message")
    
    'doing the progress bar
    ProgressBar1.Max = f
    ProgressBar1.Value = 0
    
    'disbale buttons until exporting is finished
    cmdAddSID.Enabled = False
    cmdDelSID.Enabled = False
    cmdExportCSV.Enabled = False
    cmdQuickExport.Enabled = False
    cmdLoadmdbFile.Enabled = False
    cmdBatchExport.Enabled = False
    
    'user input to export the plaser table to csv
    Dim plaserexp As Integer
    plaserexp = ACmsgbox(5, "Export the PLaser_IRI table to CSV?", vbYesNoCancel, "Message")

    If plaserexp = 2 Then
    GoTo Cancel
    
    Else
        'batch process mdb and export gps_processed and plaser_iri tables to csv
        Dim myfile As String
        myfile = Dir(CurDir() & "\" & "*.mdb")
        
        'check if mdb files are present in current directory
        If myfile = "" Then
            GoTo Missing
        Else
            Do While myfile <> ""
            GPSDB = myfile
            txtFileName.Text = GPSDB
                Call display_tbl
                Call open_mdb
                    With recs
                        .Open ("Select * from Survey_Header")
                        If .RecordCount <> 0 Then
                            Set DataGrid1.DataSource = recs
                            'table index varies, commented for future update
                            'List1.ListIndex = 12
                            List1.Text = "Survey_Header"
                            surveyid = recs!SURVEY_ID
                            Label2.Caption = "SURVEY_ID " & surveyid & " loaded: " & Fix((ProgressBar1.Value / ProgressBar1.Max) * 100) & "%" & " completed."
                            Label2.Refresh
                        End If
                    End With
                Call close_mdb
            
            'adding the survey_id fields
            Call open_mdb
            With con
                .Execute "Alter table " & "[GPS_Processed_" & surveyid & "]" & "  add SURVEY_ID Text"
                .Execute "Alter table " & "[PLaser_IRI_" & surveyid & "]" & " add SURVEY_ID Text"
            End With
            Call close_mdb
            
            'populating survey_id fields
            Dim sql As String
            Call open_mdb
            With con
                sql = "UPDATE [GPS_Processed_" & surveyid & "]" & " SET [GPS_Processed_" & surveyid & "]" & ".SURVEY_ID=" & "'" & surveyid & "'" & ""
                .Execute sql
        
            End With
            Call close_mdb
        
            Call open_mdb
            With con
                sql = "UPDATE [PLaser_IRI_" & surveyid & "]" & " SET [PLaser_IRI_" & surveyid & "]" & ".SURVEY_ID=" & "'" & surveyid & "'" & ""
                .Execute sql
            End With
            Call close_mdb
                        
            'exporting gps_processed and plaser_iri to csv
            Dim x, y As String
            x = "GPS_Processed_" & surveyid
    
            'delete invalid characters in arcgis
            y = x
    
            y = Replace$(y, "Processed_", "")
            y = Replace$(y, "-", "_")
            y = Replace$(y, " ", "_")
            y = Replace$(y, ".", "_")
            y = Replace$(y, ",", "_")
            y = Replace$(y, "&", "_")
            y = Replace$(y, "(", "_")
            y = Replace$(y, ")", "")
    
            'delete extra underscores
            y = Replace$(y, "_____", "_")
            y = Replace$(y, "____", "_")
            y = Replace$(y, "___", "_")
            y = Replace$(y, "__", "_")
    
            y = y & ".csv"
            Kill App.Path & "\" & y
            
            Call open_mdb
            With con
                .Execute "SELECT * INTO [Text;Database=" & App.Path & ";HDR=Yes;FMT=Delimited].[" & y & "] FROM [" & x & "]"
            End With
            Call close_mdb
            Kill App.Path & "\schema.ini"
    
    
            If plaserexp = 6 Then
                'exporting plaser_iri
                Dim a, b As String
                a = "PLaser_IRI_" & surveyid
        
                'delete invalid characters in arcgis
                b = a
        
                b = Replace$(b, "PLaser_", "")
                b = Replace$(b, "-", "_")
                b = Replace$(b, " ", "_")
                b = Replace$(b, ".", "_")
                b = Replace$(b, ",", "_")
                b = Replace$(b, "&", "_")
                b = Replace$(b, "(", "_")
                b = Replace$(b, ")", "")
    
                'delete extra underscores
                b = Replace$(b, "_____", "_")
                b = Replace$(b, "____", "_")
                b = Replace$(b, "___", "_")
                b = Replace$(b, "__", "_")
        
                b = b & ".csv"
                Kill App.Path & "\" & b
                
                Call open_mdb
                With con
                    .Execute "SELECT * INTO [Text;Database=" & App.Path & ";HDR=Yes;FMT=Delimited].[" & b & "] FROM [" & a & "]"
                End With
                Call close_mdb
                Kill App.Path & "\schema.ini"
                GoTo Continue
            Else
                GoTo Continue
            End If
            
Continue:
            ProgressBar1.Value = ProgressBar1.Value + 1
            myfile = Dir()
            DoEvents
            Loop
        End If
    End If
    
Export:
    Label2.Caption = "Batch Export to CSV completed successfully."
    'replaced with autoclose
    'MsgBox "Batch Export to CSV completed successfully.", vbInformation, "Message"
    Call ACmsgbox(3, "Batch Export to CSV completed successfully.", vbInformation, "Message")
    ProgressBar1.Value = 0
    
    cmdAddSID.Enabled = True
    cmdDelSID.Enabled = True
    cmdExportCSV.Enabled = True
    GoTo Term
    
Missing:
    'replaced with autoclose
    'MsgBox "No mdb files found in the current directory.", vbInformation, "Message"
    Call ACmsgbox(3, "No mdb files found in the current directory.", vbInformation, "Message")
    GoTo Term
    
Cancel:
    Call ACmsgbox(3, "Batch Export canceled.", vbInformation, "Message")

Term:
    cmdQuickExport.Enabled = True
    cmdLoadmdbFile.Enabled = True
    cmdBatchExport.Enabled = True

End Sub


'disabling the buttons
Private Sub Form_Load()
    cmdAddSID.Enabled = False
    cmdDelSID.Enabled = False
    cmdExportCSV.Enabled = False
End Sub


'display table records
Private Sub List1_Click()
    Call refresh_dgrid
End Sub


'mail to aex.gisco@gmail.com
Private Sub Image1_Click()
    ShellExecute hWnd, "open", "mailto:aex.gisco@gmail.com" & vbNullString & vbNullString & vbNullString & vbNullString, vbNullString, vbNullString, SW_SHOW
End Sub


'change pointer on mouse over
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub
