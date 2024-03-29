VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ConnectionString As String

Public DebugMode As Boolean

Private TableName As String

Private Query As String

Private ExecuteQuery As Boolean

Private FieldsStr As String

Private WhereStr As String

Private OrderByStr As String

Private GroupByStr As String

Private InnerJoinStr As String

Private HavingStr As String

Public Function SelectFields(ParamArray Fields() As Variant) As Database
    Dim Field As Variant

    For Each Field In Fields
        FieldsStr = IIf(FieldsStr = "", Field, FieldsStr & ", " & Field)
    Next Field
  
    Set SelectFields = Me
End Function

Public Function GetAll() As Collection
    AddSelectStatement
    AddInnerJoin
    AddWhere
    AddGroupBy
    AddHaving
    AddOrderBy
    
    If ExecuteQuery Then Set GetAll = GetData
End Function

Public Function Where(ParamArray chunk() As Variant) As Database
    Dim NumberArgs As Integer
    NumberArgs = UBound(chunk) + 1

    Select Case NumberArgs
        Case 1:
            WhereStr = chunk(0)
        Case 2:
            WhereStr = chunk(0) & " = " & FormatValue(chunk(1))
        Case 3:
            WhereStr = chunk(0) & " " & chunk(1) & " " & FormatValue(chunk(2))
    End Select
    
    Set Where = Me
End Function

Public Function OrderBy(ParamArray Fields() As Variant) As Database
    Dim Field As Variant

    For Each Field In Fields
        OrderByStr = IIf(OrderByStr = "", Field, OrderByStr & ", " & Field)
    Next Field

    Set OrderBy = Me
End Function

Public Function GroupBy(ParamArray Fields() As Variant) As Database
    Dim Field As Variant

    For Each Field In Fields
        GroupByStr = IIf(GroupByStr = "", Field, GroupByStr & ", " & Field)
    Next Field

    Set GroupBy = Me
End Function

Public Function Having(Condition As String) As Database
    HavingStr = Condition
    Set Having = Me
End Function

Public Function Join(TableJoin As String, IDTable As String, Equal As String, OtherIDTable As String) As Database
    Dim JoinStr As String
    JoinStr = "INNER JOIN " & TableJoin & " ON " & IDTable & " " & Equal & " " & OtherIDTable
    InnerJoinStr = IIf(InnerJoinStr = "", TableName & " " & JoinStr, "(" & InnerJoinStr & ") " & JoinStr)
    Set Join = Me
End Function

Private Function GetData() As Collection
    Dim Column As Long
    Dim rs As ADODB.Recordset
    Dim Rows As New Collection
    Dim Row As Scripting.Dictionary
    
    Set rs = Recordset
    
    Do While Not rs.EOF
        Set Row = New Scripting.Dictionary

        For Column = 0 To rs.Fields.Count - 1
            Row(rs.Fields(Column).Name) = IIf(IsNull(rs.Fields(Column)), Empty, rs.Fields(Column))
        Next Column
        
        Rows.Add Row
        
        rs.MoveNext
    Loop
    
    Set GetData = Rows
End Function

Private Sub AddSelectStatement()
    If InnerJoinStr = "" Then
        Query = "SELECT " & GetFields & " FROM " & TableName & Query
    Else
        Query = "SELECT " & GetFields & " FROM" & Query
    End If
End Sub

Private Sub AddWhere()
    If WhereStr <> "" Then
        Query = Query & " WHERE " & WhereStr
    End If
End Sub

Private Sub AddOrderBy()
    If OrderByStr <> "" Then
        Query = Query & " ORDER BY " & OrderByStr
    End If
End Sub

Private Sub AddGroupBy()
    If GroupByStr <> "" Then
        Query = Query & " GROUP BY " & GroupByStr
    End If
End Sub

Private Sub AddHaving()
    If HavingStr <> "" Then
        Query = Query & " HAVING " & HavingStr
    End If
End Sub

Private Sub AddInnerJoin()
    If InnerJoinStr <> "" Then
        Query = Query & " " & InnerJoinStr
    End If
End Sub

Private Function GetFields() As String
    If FieldsStr = "" Then
        GetFields = "*"
    Else
        GetFields = FieldsStr
    End If
End Function

Public Function Table(Name As String, Optional Execute As Boolean = True) As Database
    TableName = Name
    ExecuteQuery = Execute
    Set Table = Me
End Function

Public Function Insert(Data As Scripting.Dictionary) As Long
    Dim key As Variant
    Dim Fields As String
    Dim Values As String
    Dim Value As Variant

    For Each key In Data.Keys
        Fields = IIf(Fields = "", key, Fields & ", " & key) ' field1, field2, ...
        Value = FormatValue(Data(key))
        Values = IIf(Values = "", Value, Values & ", " & Value) ' value1, value2, ...
    Next key

    Query = "INSERT INTO " & TableName & " (" & Fields & ") VALUES (" & Values & ")"
    
    ConnectionExecute
    Insert = GetLastId
End Function

Private Function GetLastId() As Long
    Dim rs As ADODB.Recordset
    Query = "SELECT MAX(id) as max_id FROM " & TableName
    Set rs = Recordset
    GetLastId = rs.Fields(0)
End Function

Sub Update(Data As Scripting.Dictionary, Condition As String)
    Dim key As Variant
    Dim FieldsValues As String
    Dim FieldValue As String

    For Each key In Data.Keys
        FieldValue = key & " = " & FormatValue(Data(key)) ' field_name = value
        FieldsValues = IIf(FieldsValues = "", FieldValue, FieldsValues & ", " & FieldValue)
    Next key

    'FieldsValues = FieldsValues & ", updated_at = NOW()"

    Query = "UPDATE " & TableName & " SET " & FieldsValues & " WHERE " & Condition

    ConnectionExecute
End Sub

Sub Delete(Condition As String)
    Query = "DELETE FROM " & TableName & " WHERE " & Condition
    ConnectionExecute
End Sub

Sub DeleteId(Id As Integer)
    Query = "DELETE FROM " & TableName & " WHERE id = " & Id
    ConnectionExecute
End Sub

Public Function Count() As Long
    Dim rs As ADODB.Recordset
    Query = "SELECT COUNT(*) as rows_count FROM " & TableName
    Set rs = Recordset
    Count = rs.Fields(0)
End Function

Private Function FormatValue(Value As Variant) As Variant
    If IsNull(Value) Then
        FormatValue = "Null"
    ElseIf IsDate(Value) Then
        FormatValue = "#" & Format(Value, "yyyy-mm-dd hh:nn:ss") & "#"
    ElseIf WorksheetFunction.IsText(Value) Then
        FormatValue = "'" & Value & "'"
    ElseIf VarType(Value) = vbBoolean Then
        FormatValue = IIf(Value, 1, 0)
    ElseIf IsNumeric(Value) Then
        FormatValue = Value
    End If
End Function

Private Function Recordset() As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    If DebugMode Then Debug.Print Query
    rs.Open Query, Connection, adOpenStatic, adLockOptimistic

    Set Recordset = rs
End Function

Private Function Connection() As ADODB.Connection
    Dim cn As New ADODB.Connection
    cn.Open ConnectionString
    Set Connection = cn
End Function

Private Sub ConnectionExecute()
    If ExecuteQuery Then
        If DebugMode Then Debug.Print Query
        Connection.Execute Query, , adCmdText
    End If
End Sub

'Tests
Public Function GetQuery() As String
    GetQuery = Query
End Function
