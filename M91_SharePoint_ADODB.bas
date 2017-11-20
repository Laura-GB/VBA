Attribute VB_Name = "M91_SharePoint_ADODB"
Option Explicit

Const zzshForm = "Form"
Const zzshData = "Data"

Const zzrngSite = "SiteURL"
Const zzrngList = "ListGUID"

Const zzConnectSharepoint = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes;"

Enum xxLookupColumns
    xxSheetName = 1
    xxURL = 2
    xxGUID = 4
End Enum

Sub GetData()

    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim Connect As String
    Dim SQL As String
    Dim sh As Worksheet
    Dim Counter As Integer

    '=== Initialise
    Set cnn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    '=== Open connection
    Connect = zzConnectSharepoint & _
                "DATABASE=" & Range(zzrngSite).Value & ";" & _
                "LIST=" & Range(zzrngList).Value & ";"
    cnn.Open Connect
    
    '=== Get Data
    SQL = "Select * from List"
    rs.Open SQL, cnn
    
    '=== Put on sheet
    Set sh = Sheets(zzshData)
    sh.Range("A2").CopyFromRecordset rs
    
    '=== Add Titles
    For Counter = 0 To rs.Fields.Count - 1
        sh.Range("A1").Offset(0, Counter).Value = rs.Fields(Counter).Name
    Next Counter
    
    '=== Tidy up
    rs.Close
    cnn.Close
    Set cnn = Nothing
    Set rs = Nothing
End Sub

Sub AddData()

    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim Connect As String
    Dim SQL As String
    Dim sh As Worksheet
    Dim Counter As Integer

    '=== Initialise
    Set cnn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    '=== Open connection
    Connect = zzConnectSharepoint & _
                "DATABASE=" & Range(zzrngSite).Value & ";" & _
                "LIST=" & Range(zzrngList).Value & ";"
    cnn.Open Connect
    
    '=== Get Data
    SQL = "Select * from List"
    rs.Open SQL, cnn, adOpenDynamic, adLockOptimistic
    
    '=== Add Row
    With rs
        .AddNew
        .Fields("Title").Value = "Another Test"
        .Update
    End With
    
    
    '=== Tidy up
    rs.Close
    cnn.Close
    Set cnn = Nothing
    Set rs = Nothing
    
End Sub

Function FixGUID(PastedGUID As String) As String

    FixGUID = Replace(PastedGUID, "%7B", "{")
    FixGUID = Replace(FixGUID, "%7D", "}")
    FixGUID = Replace(FixGUID, "%2D", "-")

End Function

Sub UploadData(SheetName As String)

    Dim shData As Worksheet
    Dim shLookup As Worksheet
    Dim LookupRow As Long
    Dim URL As String
    Dim GUID As String
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim Connect As String
    Dim SQL As String
    Dim Headings As Range
    Dim Heading As Range
    Dim Counter As Integer
    
    Set shData = Sheets(SheetName)
    Set Headings = Range(shData.Range("A1"), shData.Range("A1").End(xlToRight))
    
    '=== Get info from Lookup Sheet
    Set shLookup = Sheets("WES Lists")
    LookupRow = Application.WorksheetFunction.Match(SheetName, shLookup.Columns(1), 0)
    URL = shLookup.Cells(LookupRow, xxURL).Value
    GUID = shLookup.Cells(LookupRow, xxGUID).Value

    '=== Initialise
    Set cnn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    '=== Open connection
    Connect = zzConnectSharepoint & _
                "DATABASE=" & URL & ";" & _
                "LIST=" & GUID & ";"
    cnn.Open Connect
    
    '=== Get Data
    SQL = "Select * from List"
    rs.Open SQL, cnn, adOpenDynamic, adLockOptimistic
    
    Range("A2").Select
    
    '=== Loop through all data
    Do Until ActiveCell.Value = ""
        With rs
            .AddNew
            For Counter = 0 To Headings.Cells.Count - 1
                .Fields(Headings.Cells(Counter + 1).Value).Value = ActiveCell.Offset(0, Counter).Value
            Next Counter
            .Update
        End With
        '=== Move down one
        ActiveCell.Offset(1, 0).Select
    Loop
    
    '=== Tidy up
    rs.Close
    cnn.Close
    Set cnn = Nothing
    Set rs = Nothing
End Sub
