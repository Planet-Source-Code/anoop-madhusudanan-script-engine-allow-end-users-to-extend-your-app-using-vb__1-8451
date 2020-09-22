Attribute VB_Name = "modDatabase"
'
'This Module Contains Declarations for loading the
'scripts from the database
'
'I am not at all touching ADO here (despite my strong temptation)
'because my purpose is to implement Scripting as easy as possible.
'DAO is much and enough to this purpose





Public Function FillListWithRecords(DBName As String, SQLQuery As String, Lview As ListView, Optional ItemIcon As Integer = 0, Optional FieldToSkip As Integer = 0) As Boolean

'On Error GoTo NoFill

Dim mI, mJ

'=============================================================================================================================
'MODDATABASE
'
'Some DAO Routies for our app
'=============================================================================================================================

Dim mLIt As ListItem

Dim mDB As Database
Dim mRec As Recordset

Set mDB = OpenDatabase(DBName)
Set mRec = mDB.OpenRecordset(SQLQuery)

Lview.ListItems.Clear
Lview.ColumnHeaders.Clear

mRec.MoveLast
mRec.MoveFirst


For mI = 1 To mRec.Fields.Count - 1 - FieldToSkip
    'We are not adding the first field name, assumes that it is index
    Lview.ColumnHeaders.Add , , mRec.Fields(mI).Name
Next mI

For mI = 0 To mRec.RecordCount - 1
    With mRec
    If ItemIcon <> 0 Then
        Set mLIt = Lview.ListItems.Add(, "K" & CStr(.Fields(0).Value), .Fields(1).Value, ItemIcon, ItemIcon)
    Else
        Set mLIt = Lview.ListItems.Add(, "K" & CStr(.Fields(0).Value), .Fields(1).Value)
    End If
        
        For mJ = 2 To Lview.ColumnHeaders.Count
        mLIt.SubItems(mJ - 1) = .Fields(mJ).Value
        Next mJ
        .MoveNext
    End With
Next mI

FillListWithRecords = True


On Error Resume Next
mRec.Close
mDB.Close


Exit Function

NoFill:

On Error Resume Next
Set mRec = Nothing
Set mDB = Nothing

FillListWithRecords = False

End Function

