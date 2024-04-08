VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCopyEntityName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements LSEXT.IEntityExtension

Option Explicit

Private Const NUMERIC_ARGUMENT = "N"
Private Const TEXT_ARGUMENT = "T"
Private Const DATE_ARGUMENT = "D"
Private Const FILE_EXT = ".txt"

Private aConnection As New ADODB.Connection
Private NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
Private NtlsUser As LSSERVICEPROVIDERLib.NautilusUser
Private sp As LSSERVICEPROVIDERLib.NautilusServiceProvider

Private ArgumentParam As New Scripting.Dictionary
Private OperatorID As String
Private OperatorName As String
Private tableName As String
Private WhereStatement As String
Private FolderID As String
Private FolderName As String
Private FilterID As String
Private SchemaEntityID As String
Private sDay As Variant
Private PageNumber As Integer
Private SelectClause As String
Private FromClause As String
Private WhereClause As String
Private FileName As String
Private ts_result_file As TextStream
Private result_file_machine As New FileSystemObject
Private FileCreated As Boolean
Private strColumnHeaders As String
Public isArgumentNull As Boolean
Private IsExit As Boolean



Private Function IEntityExtension_CanExecute(Parameters As LSEXT.IExtensionParameters) As LSEXT.ExecuteExtension

    'Default to show up on the menu
    IEntityExtension_CanExecute = exEnabled

End Function

Private Sub IEntityExtension_Execute(Parameters As LSEXT.IExtensionParametersEx)

    On Error GoTo ErrEnd
    
    Dim RecordRs As ADODB.Recordset
    Dim aDatasource As String
    Dim aDay As Variant
    Dim StoreArguments As String
    Dim CountArguments As Integer
''    isArgumentNull = False
'    ArgumentParam.RemoveAll
'
    Set sp = Parameters("SERVICE_PROVIDER")

    Set NtlsCon = sp.QueryServiceProvider("DBConnection")
    Set NtlsUser = sp.QueryServiceProvider("User")
    aDatasource = NtlsCon.GetADOConnectionString
    aConnection.Open aDatasource
    aConnection.CursorLocation = adUseClient
'    aConnection.Execute "SET ROLE LIMS_USER"
    Call ConnectSameSession(CDbl(NtlsCon.GetSessionId))
'
'    OperatorID = NtlsUser.GetOperatorId
'    OperatorName = GetOperatorName
'
'    strColumnHeaders = ""
'    FileCreated = False

    Set RecordRs = Parameters("RECORDS")

    If Not RecordRs.EOF Then
        ClipBoard_SetData (RecordRs("NAME"))
    End If
'    aConnection.Close
    Exit Sub

ErrEnd:
    MsgBox "IEntityExtension_Execute... " & vbCrLf & _
            Err.Description
End Sub
'
Private Sub ConnectSameSession(ByVal aSessionID)
    Dim aProc As New ADODB.Command
    Dim aSession As New ADODB.Parameter

    aProc.ActiveConnection = aConnection
    aProc.CommandText = "lims.lims_env.connect_same_session"
    aProc.CommandType = adCmdStoredProc

    aSession.Type = adDouble
    aSession.Direction = adParamInput
    aSession.Value = aSessionID
    aProc.Parameters.Append aSession

    aProc.Execute
    Set aSession = Nothing
    Set aProc = Nothing
End Sub
'
'Private Function nte(e As Variant) As Variant
'    nte = IIf(IsNull(e), "", e)
'End Function
'
'Private Function GetOperatorName() As String
'    On Error GoTo ErrEnd
'    Dim strSQL As String
'    Dim OperatorRS As ADODB.Recordset
'
'    GetOperatorName = ""
'    strSQL = "select " & _
'                "o.full_name " & _
'             "from " & _
'                "lims_sys.operator o " & _
'             "where " & _
'                "o.operator_id = '" & OperatorID & "'"
'    Set OperatorRS = aConnection.Execute(strSQL)
'    If Not OperatorRS.EOF Then
'        GetOperatorName = Trim(nte(OperatorRS("FULL_NAME")))
'    End If
'    OperatorRS.Close
'    Exit Function
'
'ErrEnd:
'    MsgBox "GetOperatorName... " & vbCrLf & _
'            Err.Description
'End Function
'
'Private Function GetFilterID() As String
'    On Error GoTo ErrEnd
'    Dim strSQL As String
'    Dim FolderRS As ADODB.Recordset
'
'    GetFilterID = ""
'    strSQL = "select " & _
'                "f.filter_id, f.name, f.schema_entity_id " & _
'             "from " & _
'                "lims_sys.folder f " & _
'             "where " & _
'                "f.folder_id = '" & FolderID & "'"
'    Set FolderRS = aConnection.Execute(strSQL)
'    If Not FolderRS.EOF Then
'        GetFilterID = Trim(nte(FolderRS("FILTER_ID")))
'        FolderName = Trim(nte(FolderRS("NAME")))
'        SchemaEntityID = Trim(nte(FolderRS("SCHEMA_ENTITY_ID")))
'    End If
'    FolderRS.Close
'    Exit Function
'
'ErrEnd:
'    MsgBox "GetFilterID... " & vbCrLf & _
'            Err.Description
'End Function
'
'Private Sub GetEntityContent()
'    On Error GoTo ErrEnd
'    Dim IndxPos As Integer
''    isArgumentNull = False
'    IsExit = False
'
'    TableName = GetTableName
'    If Trim(TableName) = "" Then
'        MsgBox "Table must be selected in folder properties !", _
'               vbOKOnly + vbCritical, _
'               "Copy To Clipboard"
'        Exit Sub
'    End If
'
'    WhereStatement = GetWhereStatement
'    If Trim(WhereStatement) <> "" Then
'        IndxPos = InStr(1, WhereStatement, "#")
'        If IndxPos > 0 Then
'            GetFilterArgument
'            Unload frmMsg
'        End If
'    End If
'
'    If IsExit Then GoTo ExitSub
'
'    GetSelectClause
'    GetFromClause
'    GetWhereClause
'
'    RunStatement
'ExitSub:
'    Exit Sub
'
'ErrEnd:
'    MsgBox "GetEntityContent... " & vbCrLf & _
'            Err.Description
'End Sub
'
'Private Function GetTableName() As String
'    On Error GoTo ErrEnd
'    Dim strSQL As String
'    Dim SchemaTableRS As ADODB.Recordset
'
'    GetTableName = ""
'    strSQL = "select " & _
'                "st.database_name table_name " & _
'             "from " & _
'                "lims_sys.folder f, " & _
'                "lims_sys.schema_entity se, " & _
'                "lims_sys.schema_table st " & _
'             "where " & _
'                "f.schema_entity_id = se.schema_entity_id and " & _
'                "se.schema_table_id = st.schema_table_id and " & _
'                "f.folder_id = '" & FolderID & "'"
'
'    Set SchemaTableRS = aConnection.Execute(strSQL)
'    If Not SchemaTableRS.EOF Then
'        GetTableName = Trim(nte(SchemaTableRS("table_name")))
'    End If
'    SchemaTableRS.Close
'    Exit Function
'
'ErrEnd:
'    MsgBox "GetTableName... " & vbCrLf & _
'            Err.Description
'End Function
'
'Private Function CheckOperatorFilter() As String
'    On Error GoTo ErrEnd
'    Dim strSQL As String
'    Dim OperArgRS As ADODB.Recordset
'
'    CheckOperatorFilter = ""
'    strSQL = "select " & _
'                "store_arguments " & _
'             "from " & _
'                "lims_sys.operator_filter " & _
'             "where " & _
'                "filter_id = '" & FilterID & "' and " & _
'                "operator_id = '" & OperatorID & "'"
'    Set OperArgRS = aConnection.Execute(strSQL)
'    If Not OperArgRS.EOF Then
'        CheckOperatorFilter = Trim(nte(OperArgRS(0)))
'    End If
'    OperArgRS.Close
'    Exit Function
'
'ErrEnd:
'    MsgBox "CheckOperatorFilter... " & vbCrLf & _
'            Err.Description
'End Function
'
'Private Function CheckCountArguments() As Integer
'    On Error GoTo ErrEnd
'    Dim strSQL As String
'    Dim ArgRS As ADODB.Recordset
'
'    CheckCountArguments = 0
'    strSQL = "select " & _
'                "count(*) " & _
'             "from " & _
'                "lims_sys.filter_argument " & _
'             "where " & _
'                "filter_id = '" & FilterID & "'"
'    Set ArgRS = aConnection.Execute(strSQL)
'    CheckCountArguments = Trim(nte(ArgRS(0)))
'    ArgRS.Close
'    Exit Function
'
'ErrEnd:
'    MsgBox "CheckCountArguments... " & vbCrLf & _
'            Err.Description
'End Function
'
'Private Function GetWhereStatement() As String
'    On Error GoTo ErrEnd
'    Dim strSQL As String
'    Dim FilterRS As ADODB.Recordset
'
'    GetWhereStatement = ""
'    strSQL = "select " & _
'                "f.where_statement where_state " & _
'             "from " & _
'                "lims_sys.filter f " & _
'             "where " & _
'                "f.filter_id = '" & FilterID & "'"
'    Set FilterRS = aConnection.Execute(strSQL)
'    If Not FilterRS.EOF Then
'        GetWhereStatement = Trim(nte(FilterRS("where_state")))
'    End If
'    FilterRS.Close
'    Exit Function
'
'ErrEnd:
'    MsgBox "GetWhereStatement... " & vbCrLf & _
'            Err.Description
'End Function
'
'Private Sub GetFilterArgument()
'    On Error GoTo ErrEnd
'    Dim strSQL As String
'    Dim FilterArgRS As ADODB.Recordset
'    Dim ArgumentType As String
'    Dim ArgumentValue As String
'    Dim i As Integer
'    Dim rsPhrase As Recordset
'    Dim ParentDicPhrase As New Dictionary
'    Dim ChildDicPhrase As Dictionary
'    Dim DicTemp As New Dictionary
'
'    Dim sql As String
'
'
'    strSQL = "select " & _
'                "fa.argument_type arg_type, " & _
'                "fa.name arg_name, " & _
'                "fa.value_number def_val_num, " & _
'                "fa.value_text def_val_text, " & _
'                "to_date(fa.value_date,'dd/mm/yy hh24:mi:ss') def_val_date, " & _
'                "ofa.value_number val_num, " & _
'                "ofa.value_text val_text, " & _
'                "to_date(ofa.value_date,'dd/mm/yy hh24:mi:ss') val_date, " & _
'                " fa.PHRASE_ID " & _
'             "from " & _
'                "lims_sys.filter_argument fa, " & _
'                "lims_sys.operator_filter_argument ofa " & _
'             "where " & _
'                "fa.filter_argument_id = ofa.filter_argument_id(+) and " & _
'                "ofa.operator_id(+) = '" & OperatorID & "' and " & _
'                "fa.filter_id = '" & FilterID & "' " & _
'             "order by fa.order_number"
'    Set FilterArgRS = aConnection.Execute(strSQL)
''    If FilterArgRS.EOF Then isArgumentNull = True
'
'    While Not FilterArgRS.EOF
'
'        ArgumentType = Trim(nte(FilterArgRS("arg_type")))
'        ArgumentValue = ""
'
'        Select Case ArgumentType
'            Case NUMERIC_ARGUMENT
'                ArgumentValue = Trim(nte(FilterArgRS("val_num")))
'                If Trim(ArgumentValue) = "" Then
'                    ArgumentValue = Trim(nte(FilterArgRS("def_val_num")))
'                End If
'            Case TEXT_ARGUMENT
'                ArgumentValue = Trim(nte(FilterArgRS("val_text")))
'                If Trim(ArgumentValue) = "" Then
'                    ArgumentValue = Trim(nte(FilterArgRS("def_val_text")))
'                End If
'            Case DATE_ARGUMENT
'                ArgumentValue = Trim(nte(FilterArgRS("val_date").Value))
'                If Trim(ArgumentValue) = "" Then
'                    ArgumentValue = Trim(nte(FilterArgRS("def_val_date")))
'                End If
'            Case "P" 'Phrase Name ->cmb value=p.name cmb.caption=p.description
'                ArgumentValue = Trim(nte(FilterArgRS("val_text")))
'                If Trim(ArgumentValue) = "" Then
'                    ArgumentValue = Trim(nte(FilterArgRS("def_val_text")))
'                End If
'                sql = " select PHRASE_NAME,PHRASE_DESCRIPTION"
'                sql = sql & "  from lims_sys.PHRASE_ENTRY  where phrase_id=" & nte(FilterArgRS("PHRASE_ID"))
'                sql = sql & "  order by decode(PHRASE_NAME,'" & ArgumentValue & "',1) asc"
'                Set rsPhrase = aConnection.Execute(sql)
'                Set ChildDicPhrase = New Dictionary
'                If Not rsPhrase.EOF Then
'                    ArgumentValue = nte(rsPhrase("PHRASE_DESCRIPTION"))
'                End If
'                While Not rsPhrase.EOF
'                    If Not ChildDicPhrase.Exists(nte(rsPhrase("PHRASE_DESCRIPTION"))) Then
'                        Call ChildDicPhrase.Add(nte(rsPhrase("PHRASE_DESCRIPTION")), nte(rsPhrase("PHRASE_NAME")))
'                    End If
'                rsPhrase.MoveNext
'                Wend
'                Call ParentDicPhrase.Add(CStr(i), ChildDicPhrase)
'                rsPhrase.MoveFirst
'
'
'            Case "E" 'Entity_id
'                ArgumentValue = Trim(nte(FilterArgRS("val_num")))
'                If Trim(ArgumentValue) = "" Then
'                    ArgumentValue = Trim(nte(FilterArgRS("def_val_num")))
'                End If
'            Case Else
'                ArgumentValue = Trim(nte(FilterArgRS("val_num"))) & _
'                Trim(nte(FilterArgRS("val_text").Value)) & _
'                Trim(nte(FilterArgRS("val_date").Value))
'        End Select
'         Call frmMsg.AddArgument(i, Trim(nte(FilterArgRS("arg_name"))), ArgumentValue, _
'                                ArgumentType, rsPhrase)
'
'        i = i + 1
'        FilterArgRS.MoveNext
'    Wend
'    If frmMsg.CostumeMsgbox(".יש להוסיף נתונים לפילטר, " & _
'        vbCrLf & "האם ברצונך להשתמש בנתונים הנ""ל?") = vbNo Then
'            IsExit = True
'
'             Exit Sub
'    End If
'    With frmMsg
'        For i = 0 To .LblArgName.Count - 1
'        Call ArgumentParam.Add(CStr(i), .LblArgName(i) & ", Value = " & .cmbArgVal(i))
'            If Trim(.cmbArgVal(i)) = "" _
'            Or Trim(.cmbArgVal(i)) = "?" Then
'                MsgBox "חסר נתון עבור " & .LblArgName(i) & vbCrLf & _
'                "לא ניתן להשלים את הבקשה."
'            Else
'                If .lblArgType(i) = DATE_ARGUMENT Then
'                           .cmbArgVal(i) = "to_date('" & _
'                            .cmbArgVal(i) & _
'                            "', 'dd/mm/yyyy hh24:mi:ss')"
'                End If
'            If .lblArgType(i).Caption = "P" Then
'                Set DicTemp = ParentDicPhrase(CStr(i))
'                .cmbArgVal(i) = DicTemp.Item(CStr(.cmbArgVal(i)))
'                '= ParentDicPhrase(CStr(i))
'            End If
'            WhereStatement = Replace(WhereStatement, _
'                                     "#" & _
'                                     Trim(.LblArgName(i)) & _
'                                     "#", _
'                                     .cmbArgVal(i))
'            End If
'        Next i
'    End With
'
'    FilterArgRS.Close
'
'    Exit Sub
'
'ErrEnd:
'    MsgBox "GetFilterArgument... " & vbCrLf & _
'            Err.Description
'End Sub
'
'Private Function getNextStr(ByRef s As String, c As String)
'  Dim p
'  Dim res
'  p = InStr(1, s, c)
'  If (p = 0) Then
'    res = s
'    s = ""
'    getNextStr = res
'  Else
'    res = Mid$(s, 1, p - 1)
'    s = Mid$(s, p + Len(c), 255)
'    getNextStr = res
'  End If
'End Function
'Private Sub AddTableNameBeforeVariabls()
'                 'to whereStatement
'     'add "TableName." to every statmant after and
'     'that doesnt already have a "." in them.
'     On Error GoTo err_AddTableNameBeforeVariabls
'
'    Dim sPart As String
'    Dim strLoc As Integer
'    Dim LenAdded As Integer
'    Dim Table As String, TableUser As String
'    Dim TempStr As String
'    Dim BracketsNum As Integer
'    Dim afterAnd As Boolean
'
'    afterAnd = False
'
'    If WhereStatement = "" Then Exit Sub
'    TempStr = Trim(WhereStatement)
'    strLoc = 1
'    Table = TableName & "."
'    TableUser = TableName & "_user."
'    BracketsNum = 0
'
'   While strLoc < Len(TempStr)
'    If BracketsNum > 0 Then GoTo bracketCheck
'        If strLoc = 1 Then
'FirstPlace:
'            If UCase(Mid(TempStr, strLoc, 2)) = "U_" Then
'                TempStr = TableUser & TempStr
'                'sdg.sdg_id=5
'            ElseIf InStr(strLoc, TempStr, ".") > InStr(strLoc, TempStr & " ", " ") _
'                   Or InStr(strLoc, TempStr, ".") = 0 _
'                   Or InStr(strLoc, TempStr, "rownum", vbTextCompare) > InStr(strLoc, TempStr, " ") Then
'                TempStr = Table & TempStr
'            End If
'        ElseIf UCase(Mid(TempStr, strLoc, 5)) = " AND " Then
'            afterAnd = True
'            strLoc = strLoc + 4
'        ElseIf afterAnd = True Then
'            ' if " and    u_...
'            If UCase(Mid(TempStr, strLoc, 2)) = "U_" Then
'                afterAnd = False
'                If Mid(TempStr, strLoc - 1, 1) <> "." Then
'                'don't add when "and    tableuser.u_..."
'                    TempStr = Left(TempStr, strLoc - 1) & TableUser & _
'                                Mid(TempStr, strLoc)
'                End If
'            ElseIf Mid(TempStr, strLoc, 1) <> " " Then
'               'if " and      field.."
'                afterAnd = False
'                If InStr(strLoc, TempStr, ".") > InStr(strLoc, TempStr & " ", " ") _
'                   Or InStr(strLoc, TempStr, ".") = 0 _
'                   Or InStr(strLoc, TempStr, "rownum", vbTextCompare) > InStr(strLoc, TempStr, " ") Then
'                   ' don't add when " table.field"
'                    TempStr = Left(TempStr, strLoc - 1) & Table & _
'                    Mid(TempStr, strLoc)
'                End If
'            End If
'        End If
'
'bracketCheck:
'    If Mid(TempStr, strLoc, 1) = "(" Then BracketsNum = BracketsNum + 1
'    If Mid(TempStr, strLoc, 1) = ")" Then BracketsNum = BracketsNum - 1
'   strLoc = strLoc + 1
'   Wend
'
'
''    while instr(strloc,WhereStatement ,
''
''    if instr(1,WhereStatement," ")<instr(1,WhereStatement,".")
''    strpart=    getNextStr(WhereStatement," "
''    WhereStatement = " " & TempStr & " "
'Exit Sub
'
'err_AddTableNameBeforeVariabls:
'    MsgBox "Error in AddTableNameBeforeVariabls :" & Err.Description
'
'End Sub
'
'Private Sub RunStatement()
'    On Error GoTo ErrEnd
'    Dim strSQL As String
'    Dim strSELECT As String
'    Dim strFROM As String
'    Dim strWHERE As String
'    Dim strORDERBY As String
'    Dim EntityRS As ADODB.Recordset
'    Dim SqlStatement As String
'    Dim strFileds As String
'    Dim fn As Variant
'
'    strWHERE = ""
'    If Trim(WhereStatement) <> "" Then
'        Call AddTableNameBeforeVariabls 'to whereStatement
'        strWHERE = strWHERE & " and " & _
'                     WhereStatement & " "
'    End If
'    'ambiquity problam
'
'    SqlStatement = "SELECT " & SelectClause & " " & _
'                   "FROM " & FromClause & " " & _
'                   "WHERE " & WhereClause & " " & strWHERE & " " & _
'                   "ORDER BY " & TableName & ".name"
'
'    Call WriteResultsToFile(SqlStatement)
'
'    Set EntityRS = aConnection.Execute(SqlStatement)
'
'    If EntityRS.EOF Then
'        MsgBox "You choosed an empty folder (entity)." & vbCrLf & _
'            "No data found for the selected parameters !", _
'            vbCritical, "Nautilus"
'        Exit Sub
'    End If
'
'    PageNumber = 1
'    CopyColumnsHeader
'    strFileds = ""
'    While Not EntityRS.EOF
'        For Each fn In EntityRS.Fields
'            strFileds = strFileds & vbTab & Trim(nte(fn))
'        Next
'        strFileds = strFileds & vbCrLf
'        EntityRS.MoveNext
'    Wend
'    EntityRS.Close
'
'    If FileCreated Then
'        ts_result_file.Close
'    End If
'
'    If Trim(strFileds) <> "" Then
'        strFileds = strColumnHeaders & _
'                    vbCrLf & _
'                    strFileds
'        Call ClipBoard_SetData(strFileds)
'        MsgBox "The entity information has been successfuly copied to the clipboard.", _
'                vbInformation + vbYes, "Nautilus - Copy To Clipboard"
'    End If
'    Exit Sub
'
'ErrEnd:
'    MsgBox "RunStatement... " & vbCrLf & _
'            Err.Description
'End Sub
'
'Private Sub CopyColumnsHeader()
'    On Error GoTo ErrEnd
'    'select the column headers for table name
'    Dim SqlStr  As String
'    Dim Column As ADODB.Recordset
'
'    strColumnHeaders = ""
'    SqlStr = "select sf.name || decode(ec.LOOKUP_SCHEMA_FIELD_ID, null, '', ' - ' || " & _
'                "(select sf1.name from lims_sys.schema_field sf1 where sf1.SCHEMA_FIELD_ID = ec.LOOKUP_SCHEMA_FIELD_ID)), " & _
'                    "ec.column_width " & _
'                    "from lims_sys.schema_field sf, " & _
'                    "lims_sys.entity_column ec " & _
'                    "where ec.SCHEMA_ENTITY_ID = " & SchemaEntityID & " and " & _
'                "sf.SCHEMA_FIELD_ID = ec.SCHEMA_FIELD_ID and " & _
'                "ec.OPERATOR_ID = " & NtlsUser.GetOperatorId & " " & _
'                "order by ec.ORDER_NUMBER"
'    Set Column = aConnection.Execute(SqlStr)
'    While Not Column.EOF
'        strColumnHeaders = strColumnHeaders & vbTab & Trim(nte(Column(0)))
'        Column.MoveNext
'    Wend
'    Column.Close
'    Exit Sub
'
'ErrEnd:
'    MsgBox "CopyColumnsHeader... " & vbCrLf & _
'            Err.Description
'End Sub
'
'Private Sub GetSelectClause()
'    On Error GoTo ErrEnd
'    Dim SqlStr As String
'    Dim Column As ADODB.Recordset
'    'select clause (column database names for table name)
'    SelectClause = ""
'
'    SqlStr = "select decode(ec.LOOKUP_SCHEMA_FIELD_ID, null, st.DATABASE_NAME || '.' ||sf.DATABASE_NAME, " & _
'            "(select st1.DATABASE_NAME || ec.schema_field_id || '.' ||sf1.DATABASE_NAME " & _
'            " from lims_sys.schema_table st1, lims_sys.schema_field sf1 " & _
'            "where st1.SCHEMA_TABLE_ID = sf1.SCHEMA_TABLE_ID and " & _
'            "sf1.SCHEMA_FIELD_ID = ec.LOOKUP_SCHEMA_FIELD_ID)) " & _
'            "from lims_sys.schema_field sf, " & _
'            "lims_sys.entity_column ec, " & _
'            "lims_sys.schema_table st " & _
'            "where ec.SCHEMA_ENTITY_ID = " & SchemaEntityID & " and " & _
'            "sf.SCHEMA_FIELD_ID = ec.SCHEMA_FIELD_ID and " & _
'            "ec.OPERATOR_ID = " & NtlsUser.GetOperatorId & " and " & _
'            "st.SCHEMA_TABLE_ID = sf.SCHEMA_TABLE_ID " & _
'            "order by ec.ORDER_NUMBER"
'    Set Column = aConnection.Execute(SqlStr)
'    While Not Column.EOF
'        SelectClause = SelectClause & Column(0) & ", "
'        Column.MoveNext
'    Wend
'    Column.Close
'    SelectClause = Left(SelectClause, Len(SelectClause) - 2)
'    Exit Sub
'
'ErrEnd:
'    MsgBox "GetSelectClause... " & vbCrLf & _
'            Err.Description
'End Sub
'
'Private Sub GetFromClause()
'    On Error GoTo ErrEnd
'    Dim SqlStr As String
'    Dim Column As ADODB.Recordset
'    'from clause (database table names)
'    FromClause = "lims_sys." & TableName & ", " & _
'                "lims_sys." & TableName & "_USER, "
'
'    SqlStr = "select distinct('lims_sys.' || " & _
'        "(select st1.DATABASE_NAME || ' ' || st1.DATABASE_NAME || ec.schema_field_id " & _
'        "from lims_sys.schema_table st1, lims_sys.schema_field sf1 " & _
'        "where st1.SCHEMA_TABLE_ID = sf1.SCHEMA_TABLE_ID and " & _
'        "sf1.SCHEMA_FIELD_ID = ec.LOOKUP_SCHEMA_FIELD_ID)) " & _
'        "from lims_sys.schema_field sf, " & _
'        "lims_sys.entity_column ec, " & _
'        "lims_sys.schema_table st " & _
'        "where ec.SCHEMA_ENTITY_ID = " & SchemaEntityID & " and " & _
'        "sf.SCHEMA_FIELD_ID = ec.SCHEMA_FIELD_ID and " & _
'        "ec.OPERATOR_ID = " & NtlsUser.GetOperatorId & " and " & _
'        "st.SCHEMA_TABLE_ID = sf.SCHEMA_TABLE_ID and " & _
'        "ec.LOOKUP_SCHEMA_FIELD_ID is not null"
'
'    Set Column = aConnection.Execute(SqlStr)
'    While Not Column.EOF
'        FromClause = FromClause & Column(0) & ", "
'        Column.MoveNext
'    Wend
'    Column.Close
'    FromClause = Left(FromClause, Len(FromClause) - 2)
'    Exit Sub
'
'ErrEnd:
'    MsgBox "GetFromClause... " & vbCrLf & _
'            Err.Description
'End Sub
'
'Private Sub GetWhereClause()
'    On Error GoTo ErrEnd
'    Dim SqlStr As String
'    Dim Column As ADODB.Recordset
'    'where clause
'    '(add hard coded the 'table_name.table_name_id = table_name_user.table_name_id')
'    WhereClause = TableName & "." & TableName & "_ID = " & _
'                TableName & "_USER." & TableName & "_ID and "
'
'    SqlStr = "select distinct(st.DATABASE_NAME || '.' || sf.DATABASE_NAME || ' = ' || " & _
'        "(select st1.DATABASE_NAME || ec.schema_field_id " & _
'        "from lims_sys.schema_table st1, lims_sys.schema_field sf1 " & _
'        "where st1.SCHEMA_TABLE_ID = sf1.SCHEMA_TABLE_ID and " & _
'        "sf1.SCHEMA_FIELD_ID = ec.LOOKUP_SCHEMA_FIELD_ID) || '.' || " & _
'        "lims.lims_function.get_primary_key( " & _
'        "(select st1.DATABASE_NAME " & _
'        "from lims_sys.schema_table st1, lims_sys.schema_field sf1 " & _
'        "where st1.SCHEMA_TABLE_ID = sf1.SCHEMA_TABLE_ID and " & _
'        "sf1.SCHEMA_FIELD_ID = ec.LOOKUP_SCHEMA_FIELD_ID))) || '(+)' " & _
'        "from lims_sys.schema_field sf, " & _
'        "lims_sys.entity_column ec, " & _
'        "lims_sys.schema_table st " & _
'        "where ec.SCHEMA_ENTITY_ID = " & SchemaEntityID & " and " & _
'        "sf.SCHEMA_FIELD_ID = ec.SCHEMA_FIELD_ID and " & _
'        "ec.OPERATOR_ID = " & NtlsUser.GetOperatorId & " and " & _
'        "st.SCHEMA_TABLE_ID = sf.SCHEMA_TABLE_ID and " & _
'        "ec.LOOKUP_SCHEMA_FIELD_ID is not null"
'
'    Set Column = aConnection.Execute(SqlStr)
'    While Not Column.EOF
'        WhereClause = WhereClause & Column(0) & " and "
'        Column.MoveNext
'    Wend
'    Column.Close
'    WhereClause = Left(WhereClause, Len(WhereClause) - 5)
'    Exit Sub
'
'ErrEnd:
'    MsgBox "GetWhereClause... " & vbCrLf & _
'            Err.Description
'End Sub
'
'Private Sub WriteResultsToFile(line As String)
'    On Error GoTo ErrEnd
'    If ts_result_file Is Nothing Then
'        FileCreated = True
'        FileName = "C:\" & Format(Now, "YYYYMMDD") & FILE_EXT
'        Set ts_result_file = result_file_machine.CreateTextFile(FileName, True)
'    End If
'    ts_result_file.WriteLine (line)
'    Exit Sub
'
'ErrEnd:
'    MsgBox "WriteResultsToFile... " & vbCrLf & _
'            Err.Description
'End Sub

