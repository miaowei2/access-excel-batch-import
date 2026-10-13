Option Compare Database
Option Explicit

Private Const lngADO_INTEGER As Long = 3
Private Const lngADO_BOOLEAN As Long = 11
Private Const lngADO_DECIMAL As Long = 14
Private Const lngADO_UNICODE_TEXT As Long = 202
Private Const lngADO_UNICODE_LONG_TEXT As Long = 203
Private Const lngADO_TIMESTAMP As Long = 135
Private Const lngADO_EXECUTE_NO_RECORDS As Long = 128
Public Type typImportColumn
    strFieldName As String
    lngSourceColumnIndex As Long
    lngFieldType As Long
    lngFieldSize As Long
    bytPrecision As Byte
    bytScale As Byte
    strSqlType As String
    blnRequired As Boolean
End Type
Private mobjExcelApplication As Object
Private mobjWorkbook As Object
Private mblnImportBusy As Boolean
Private mblnCancelRequested As Boolean
Private mlngExcelDateOffset As Long
Private mblnMappingReady As Boolean
Private mobjServerConnections As Object
Private mstrLanguage As String

Public Function EI_Language() As String
    If Len(mstrLanguage) = 0 Then
        mstrLanguage = EI_SavedLanguage()
    End If
    EI_Language = mstrLanguage
End Function

Private Function EI_SavedLanguage() As String
    On Error GoTo NotSaved
    EI_SavedLanguage = LCase$(CStr(CurrentProject.Properties("EI_Language").Value))
    If EI_SavedLanguage <> "zh" Then
        EI_SavedLanguage = "en"
    End If
    Exit Function
NotSaved:
    EI_SavedLanguage = "en"
End Function

Public Sub EI_SetLanguage(ByVal strLanguage As String)
    strLanguage = LCase$(Trim$(strLanguage))
    If strLanguage = "zh-cn" Or strLanguage = "cn" Or strLanguage = "chinese" Then
        strLanguage = "zh"
    ElseIf strLanguage = "en-us" Or strLanguage = "english" Then
        strLanguage = "en"
    End If
    If strLanguage <> "en" And strLanguage <> "zh" Then
        Err.Raise vbObjectError + 2142, "ExcelImport.Language", "Language must be ""en"" or ""zh"". / 语言只能是 ""en"" 或 ""zh""。"
    End If
    EI_SaveLanguage strLanguage
    mstrLanguage = strLanguage
    EI_ApplyOpenFormsLanguage
End Sub

Private Sub EI_SaveLanguage(ByVal strLanguage As String)
    On Error GoTo AddProperty
    CurrentProject.Properties("EI_Language").Value = strLanguage
    Exit Sub
AddProperty:
    CurrentProject.Properties.Add "EI_Language", strLanguage
End Sub

Public Function EI_ToggleLanguage() As Boolean
    On Error GoTo Failed
    If mblnImportBusy Then
        Exit Function
    End If
    If EI_Language() = "zh" Then
        EI_SetLanguage "en"
    Else
        EI_SetLanguage "zh"
    End If
    EI_ToggleLanguage = True
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, EI_Text("Language", "语言")
End Function

' Returns the message for the current language; English is the default.
Public Function EI_Text(ByVal strEnglish As String, ByVal strChinese As String) As String
    If EI_Language() = "zh" Then
        EI_Text = strChinese
    Else
        EI_Text = strEnglish
    End If
End Function

Public Function EI_Format(ByVal strTemplate As String, ParamArray varValues() As Variant) As String
    Dim lngValueIndex As Long
    EI_Format = strTemplate
    For lngValueIndex = LBound(varValues) To UBound(varValues)
        EI_Format = Replace(EI_Format, "{" & lngValueIndex & "}", CStr(Nz(varValues(lngValueIndex), "")))
    Next
End Function

' Uses Access's own connection so exclusive-mode databases can still be written.
Public Function EI_LocalConnection() As Object
    Set EI_LocalConnection = CurrentProject.Connection
End Function

Private Function EI_IsProjectConnection(ByVal objConnection As Object) As Boolean
    EI_IsProjectConnection = (objConnection Is CurrentProject.Connection)
End Function

Public Sub EI_RefreshMapForm()
    If CurrentProject.AllForms(strIMPORT_FORM_NAME).IsLoaded Then
        DBEngine.Idle dbRefreshCache
        Forms(strIMPORT_FORM_NAME)!sfrDetail.Form.Requery
    End If
End Sub

' Prevents the same target field from being selected for two Excel columns.
Public Function EI_MapFieldUsed(ByVal frmMap As Access.Form) As Boolean
    Dim rstMap As Object
    Dim strFieldName As String
    strFieldName = CStr(Nz(frmMap!cboAccFld, ""))
    If Len(strFieldName) = 0 Then
        Exit Function
    End If
    Set rstMap = frmMap.RecordsetClone
    If rstMap.RecordCount > 0 Then
        rstMap.MoveFirst
    End If
    Do Until rstMap.EOF
        If rstMap!ColIndex <> frmMap!txtIndex Then
            If StrComp(CStr(Nz(rstMap!AccessField, "")), strFieldName, vbTextCompare) = 0 Then
                EI_MapFieldUsed = True
                MsgBox EI_Format(EI_Text("Target field ""{0}"" is already mapped to column {1}.", "目标字段“{0}”已映射到第 {1} 列。"), strFieldName, rstMap!ColLetter), vbExclamation, EI_Text("Field mapping", "字段映射")
                Exit Function
            End If
        End If
        rstMap.MoveNext
    Loop
End Function

Private Function EI_Command(ByVal objConnection As Object, ByVal strSql As String) As Object
    Dim objCommand As Object
    Set objCommand = CreateObject("ADODB.Command")
    Set objCommand.ActiveConnection = objConnection
    objCommand.CommandType = 1
    objCommand.CommandTimeout = 120
    objCommand.CommandText = strSql
    Set EI_Command = objCommand
End Function

Private Sub EI_AddParameter(ByVal objCommand As Object, ByVal lngAdoDataType As Long, ByVal varParameterValue As Variant, Optional ByVal lngParameterSize As Long = 0, Optional ByVal bytParameterPrecision As Byte = 0, Optional ByVal bytParameterScale As Byte = 0)
    Dim objParameter As Object
    ' Reading Parameters.Count before Append triggers automatic parameter discovery.
    If lngAdoDataType = lngADO_UNICODE_TEXT Or lngAdoDataType = lngADO_UNICODE_LONG_TEXT Then
        If Not IsNull(varParameterValue) Then
            If Len(CStr(varParameterValue)) > lngParameterSize Then lngParameterSize = Len(CStr(varParameterValue))
        End If
        If lngParameterSize < 1 Then lngParameterSize = 1
        Set objParameter = objCommand.CreateParameter(, lngAdoDataType, 1, lngParameterSize)
    Else
        Set objParameter = objCommand.CreateParameter(, lngAdoDataType, 1)
    End If
    If lngAdoDataType = lngADO_DECIMAL Then
        objParameter.Precision = bytParameterPrecision
        objParameter.NumericScale = bytParameterScale
    End If
    objParameter.Value = varParameterValue
    objCommand.Parameters.Append objParameter
End Sub

Public Sub EI_Execute(ByVal objConnection As Object, ByVal strSql As String)
    objConnection.Execute strSql, , lngADO_EXECUTE_NO_RECORDS
End Sub

Public Sub EI_CloseConnection(ByRef objConnection As Object)
    If Not objConnection Is Nothing Then
        If Not EI_IsProjectConnection(objConnection) Then
            If objConnection.State <> 0 Then
                objConnection.Close
            End If
        End If
        Set objConnection = Nothing
    End If
End Sub

Public Sub EI_SetSqlServerConnection(ByVal strAccessTableName As String, ByVal strConnectionString As String)
    If mblnImportBusy Then Err.Raise vbObjectError + 2130, , EI_Text("The database connection cannot be changed during an import.", "导入期间不能更换数据库连接。")
    If mobjServerConnections Is Nothing Then
        Set mobjServerConnections = CreateObject("Scripting.Dictionary")
        mobjServerConnections.CompareMode = vbTextCompare
    End If
    If Len(strConnectionString) = 0 Then
        If mobjServerConnections.Exists(strAccessTableName) Then mobjServerConnections.Remove strAccessTableName
    Else
        mobjServerConnections(strAccessTableName) = strConnectionString
    End If
    mblnMappingReady = False
End Sub

Public Function EI_Q(ByVal strObjectName As String) As String
    If InStr(strObjectName, "]") > 0 Then
        Err.Raise vbObjectError + 2100, "ExcelImport", EI_Text("Object names cannot contain a right square bracket: ", "对象名称不能包含右方括号：") & strObjectName
    End If
    EI_Q = "[" & strObjectName & "]"
End Function

Private Function EI_Supported(ByVal fldField As DAO.Field) As Boolean
    If (fldField.Attributes And dbAutoIncrField) <> 0 Then Exit Function
    If Len(EI_Property(fldField, "Expression")) > 0 Then Exit Function
    Select Case fldField.Type
        Case dbBoolean, dbByte, dbInteger, dbLong, dbCurrency, dbSingle, dbDouble, dbDate, dbText, dbMemo, 16, 20
            EI_Supported = True
    End Select
End Function

Private Function EI_Property(ByVal objAccessObject As Object, ByVal strPropertyName As String) As String
    Dim prpFieldProperty As DAO.Property
    For Each prpFieldProperty In objAccessObject.Properties
        If prpFieldProperty.Name = strPropertyName Then
            EI_Property = Nz(prpFieldProperty.Value, "")
            Exit Function
        End If
    Next
End Function

Private Function EI_ServerTable(ByVal strSourceTableName As String) As String
    Dim varTableNameParts As Variant
    varTableNameParts = Split(strSourceTableName, ".")
    If UBound(varTableNameParts) = 0 Then
        EI_ServerTable = EI_Q(CStr(varTableNameParts(0)))
    ElseIf UBound(varTableNameParts) = 1 Then
        EI_ServerTable = EI_Q(CStr(varTableNameParts(0))) & "." & EI_Q(CStr(varTableNameParts(1)))
    Else
        Err.Raise vbObjectError + 2131, , EI_Text("The SQL Server source table must be schema.table; cross-database and four-part names are not supported.", "SQL Server 源表名应为 schema.table，不支持跨库或四段表名。")
    End If
End Function

Private Function EI_TargetMetadata(ByVal tdfTargetTable As DAO.TableDef, ByRef objTargetConnection As Object, ByRef blnSqlServer As Boolean) As Object
    Dim objFieldMetadata As Object
    Dim objFieldInfo As Object
    Dim objCommand As Object
    Dim rstData As Object
    Dim fldField As DAO.Field
    Dim strConnectionString As String
    Dim strDataType As String
    Dim strSqlType As String
    Dim lngFieldSize As Long
    Dim lngFieldType As Long
    Set objFieldMetadata = CreateObject("Scripting.Dictionary")
    objFieldMetadata.CompareMode = vbTextCompare
    blnSqlServer = (tdfTargetTable.Attributes And dbAttachedODBC) <> 0
    If Not blnSqlServer Then
        If Len(tdfTargetTable.Connect) > 0 And Left$(tdfTargetTable.Connect, 10) <> ";DATABASE=" Then
            Err.Raise vbObjectError + 2113, , EI_Text("Only local Access tables, linked Access back-end tables and SQL Server ODBC linked tables are supported.", "只支持本地 Access 表、Access 后端链接表及 SQL Server ODBC 链接表。")
        End If
        Set objTargetConnection = EI_LocalConnection()
        Set rstData = objTargetConnection.Execute("SELECT * FROM " & EI_Q(tdfTargetTable.Name) & " WHERE 1=0")
        For Each fldField In tdfTargetTable.Fields
            Set objFieldInfo = CreateObject("Scripting.Dictionary")
            objFieldInfo.Add "Name", fldField.Name
            objFieldInfo.Add "FieldType", CLng(fldField.Type)
            objFieldInfo.Add "Size", CLng(fldField.Size)
            objFieldInfo.Add "Precision", 28
            objFieldInfo.Add "Scale", 0
            If fldField.Type = 20 Then
                objFieldInfo("Precision") = CLng(rstData.Fields(fldField.Name).Precision)
                objFieldInfo("Scale") = CLng(rstData.Fields(fldField.Name).NumericScale)
            End If
            objFieldInfo.Add "Required", fldField.Required
            objFieldInfo.Add "Default", Len(Nz(fldField.DefaultValue, "")) > 0
            objFieldInfo.Add "Writable", (fldField.Attributes And dbAutoIncrField) = 0 And Len(EI_Property(fldField, "Expression")) = 0
            objFieldInfo.Add "Supported", EI_Supported(fldField)
            objFieldInfo.Add "SqlType", ""
            objFieldMetadata.Add fldField.Name, objFieldInfo
        Next
        rstData.Close
        Set EI_TargetMetadata = objFieldMetadata
        Exit Function
    End If
    If Not mobjServerConnections Is Nothing Then
        If mobjServerConnections.Exists(tdfTargetTable.Name) Then strConnectionString = mobjServerConnections(tdfTargetTable.Name)
    End If
    If Len(strConnectionString) = 0 Then
        If Left$(tdfTargetTable.Connect, 5) <> "ODBC;" Then Err.Raise vbObjectError + 2132, , EI_Text("The ODBC connection cannot be read. Call EI_SetSqlServerConnection to configure a direct ADO connection.", "无法读取 ODBC 连接，请调用 EI_SetSqlServerConnection 配置直连 ADO 连接。")
        strConnectionString = "Provider=MSDASQL;" & Mid$(tdfTargetTable.Connect, 6)
    End If
    Set objTargetConnection = CreateObject("ADODB.Connection")
    objTargetConnection.ConnectionTimeout = 15
    objTargetConnection.CommandTimeout = 120
    objTargetConnection.Open strConnectionString
    Set rstData = objTargetConnection.Execute("SELECT CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(128)) AS Version")
    If rstData.EOF Then Err.Raise vbObjectError + 2133, , EI_Text("The ODBC back end is not a supported SQL Server.", "ODBC 后端不是支持的 SQL Server。")
    If IsNull(rstData.Fields(0).Value) Then Err.Raise vbObjectError + 2133, , EI_Text("The SQL Server back end cannot be identified.", "无法识别 SQL Server 后端。")
    rstData.Close
    EI_ServerTable tdfTargetTable.SourceTableName
    Set objCommand = EI_Command(objTargetConnection, "SELECT c.name,t.name AS TypeName,c.max_length,c.precision,c.scale,c.is_nullable,c.is_identity,c.is_computed,c.default_object_id,COLUMNPROPERTY(c.object_id,c.name,'GeneratedAlwaysType') AS GeneratedType FROM sys.columns AS c INNER JOIN sys.types AS t ON c.system_type_id=t.user_type_id WHERE c.object_id=OBJECT_ID(?) AND OBJECTPROPERTY(c.object_id,'IsUserTable')=1 ORDER BY c.column_id")
    EI_AddParameter objCommand, lngADO_UNICODE_TEXT, tdfTargetTable.SourceTableName, 776
    Set rstData = objCommand.Execute
    If rstData.EOF Then Err.Raise vbObjectError + 2134, , EI_Text("The SQL Server user table was not found, or you cannot read its structure.", "未找到 SQL Server 用户表或没有读取表结构的权限。")
    Do Until rstData.EOF
        strDataType = LCase$(CStr(rstData!TypeName))
        lngFieldSize = CLng(rstData!max_length)
        lngFieldType = 0
        strSqlType = strDataType
        Select Case strDataType
            Case "bit": lngFieldType = dbBoolean
            Case "tinyint": lngFieldType = dbByte
            Case "smallint": lngFieldType = dbInteger
            Case "int": lngFieldType = dbLong
            Case "bigint": lngFieldType = 16
            Case "money", "smallmoney": lngFieldType = dbCurrency
            Case "real": lngFieldType = dbSingle
            Case "float": lngFieldType = dbDouble
            Case "decimal", "numeric"
                lngFieldType = 20
                strSqlType = strDataType & "(" & rstData!precision & "," & rstData!scale & ")"
            Case "date", "datetime", "smalldatetime", "datetime2"
                lngFieldType = dbDate
                If strDataType = "datetime2" Then strSqlType = strDataType & "(" & rstData!scale & ")"
            Case "nvarchar", "nchar", "varchar", "char"
                lngFieldType = dbText
                If strDataType = "nvarchar" Or strDataType = "nchar" Then
                    If lngFieldSize > 0 Then lngFieldSize = lngFieldSize \ 2
                End If
                If lngFieldSize = -1 Then
                    lngFieldType = dbMemo
                    strSqlType = strDataType & "(max)"
                    lngFieldSize = 0
                Else
                    strSqlType = strDataType & "(" & lngFieldSize & ")"
                End If
            Case "text", "ntext"
                lngFieldType = dbMemo
                lngFieldSize = 0
                If strDataType = "text" Then
                    strSqlType = "varchar(max)"
                Else
                    strSqlType = "nvarchar(max)"
                End If
        End Select
        Set objFieldInfo = CreateObject("Scripting.Dictionary")
        objFieldInfo.Add "Name", CStr(rstData!name)
        objFieldInfo.Add "FieldType", lngFieldType
        objFieldInfo.Add "Size", lngFieldSize
        objFieldInfo.Add "Precision", CLng(rstData!precision)
        objFieldInfo.Add "Scale", CLng(rstData!scale)
        objFieldInfo.Add "Required", Not CBool(rstData!is_nullable)
        objFieldInfo.Add "Default", CLng(rstData!default_object_id) <> 0
        objFieldInfo.Add "Writable", Not CBool(rstData!is_identity) And Not CBool(rstData!is_computed) And Nz(rstData!GeneratedType, 0) = 0 And strDataType <> "timestamp" And strDataType <> "rowversion"
        objFieldInfo.Add "Supported", lngFieldType <> 0 And CBool(objFieldInfo("Writable"))
        objFieldInfo.Add "SqlType", strSqlType
        objFieldMetadata.Add CStr(rstData!name), objFieldInfo
        rstData.MoveNext
    Loop
    rstData.Close
    Set EI_TargetMetadata = objFieldMetadata
End Function

Private Function EI_MetadataAllows(ByVal objFieldMetadata As Object, ByVal strFieldName As String) As Boolean
    If objFieldMetadata.Exists(strFieldName) Then EI_MetadataAllows = CBool(objFieldMetadata(strFieldName)("Supported"))
End Function

Private Function EI_ColumnLetter(ByVal lngColumnIndex As Long) As String
    Do While lngColumnIndex > 0
        lngColumnIndex = lngColumnIndex - 1
        EI_ColumnLetter = Chr$(65 + lngColumnIndex Mod 26) & EI_ColumnLetter
        lngColumnIndex = lngColumnIndex \ 26
    Loop
End Function

Private Sub EI_Status(ByVal strMessage As String)
    If CurrentProject.AllForms(strIMPORT_FORM_NAME).IsLoaded Then
        Forms(strIMPORT_FORM_NAME)!lblStatus.Caption = strMessage
        Forms(strIMPORT_FORM_NAME).Repaint
    End If
End Sub

Private Sub EI_CheckCancel()
    DoEvents
    If mblnCancelRequested Then Err.Raise vbObjectError + 2101, "ExcelImport", EI_Text("Import cancelled; nothing was written to the target table.", "已取消；本批次未写入目标表。")
End Sub

Public Function EI_IsBusy() As Boolean
    EI_IsBusy = mblnImportBusy
End Function

Public Function EI_Cancel() As Boolean
    mblnCancelRequested = True
    EI_Status EI_Text("Cancelling; please wait for the current step to finish...", "正在取消，请等待当前步骤结束...")
    EI_Cancel = True
End Function

Public Function EI_Close() As Boolean
    On Error GoTo Failed
    If mblnImportBusy Then
        EI_Cancel
        Exit Function
    End If
    EI_ReleaseExcel
    DoCmd.Close acForm, strIMPORT_FORM_NAME
    EI_Close = True
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, EI_Text("Close import form", "关闭导入窗体")
End Function

Public Sub EI_ReleaseExcel()
    If Not mobjWorkbook Is Nothing Then
        mobjWorkbook.Close False
        Set mobjWorkbook = Nothing
    End If
    If Not mobjExcelApplication Is Nothing Then
        mobjExcelApplication.Quit
        Set mobjExcelApplication = Nothing
    End If
End Sub

Public Function EI_Load() As Boolean
    On Error GoTo Failed
    Dim dbsDatabase As DAO.Database
    Dim tdfTargetTable As DAO.TableDef
    Dim frmImport As Access.Form
    Dim objConnection As Object
    Dim strMessage As String
    Set frmImport = Forms(strIMPORT_FORM_NAME)
    Set dbsDatabase = CurrentDb
    mblnMappingReady = False
    EI_ApplyImportFormLanguage frmImport
    Set objConnection = EI_LocalConnection()
    EI_RequireImportTables objConnection
    EI_Execute objConnection, "DELETE FROM tblEI_Map"
    frmImport!cboTableName.RowSource = ""
    frmImport!cboSheetName.RowSource = ""
    frmImport!cboSheetName.Enabled = False
    For Each tdfTargetTable In dbsDatabase.TableDefs
        If Left$(tdfTargetTable.Name, 4) <> "MSys" And Left$(tdfTargetTable.Name, 5) <> "tblEI" And Left$(tdfTargetTable.Name, 4) <> "USys" And Left$(tdfTargetTable.Name, 1) <> "~" Then
            If (tdfTargetTable.Attributes And dbHiddenObject) = 0 Then
                frmImport!cboTableName.AddItem """" & Replace(tdfTargetTable.Name, """", """""") & """"
            End If
        End If
    Next
    If Len(Nz(frmImport.OpenArgs, "")) > 0 Then
        If Not EI_TableExists(objConnection, CStr(frmImport.OpenArgs)) Then
            Err.Raise vbObjectError + 2102, , EI_Text("The target table passed in OpenArgs does not exist.", "OpenArgs 指定的目标表不存在。")
        End If
        frmImport!cboTableName = frmImport.OpenArgs
        frmImport!cboTableName.Enabled = False
    End If
    EI_CloseConnection objConnection
    EI_RefreshMapForm
    EI_Status EI_Text("Select a target table and an Excel file, then confirm the field mapping.", "请选择目标表和 Excel 文件，然后确认字段映射。")
    EI_Load = True
    If Len(Nz(frmImport.OpenArgs, "")) > 0 Then
        ' Same as the old form: a preset target table goes straight to file selection.
        frmImport.Visible = True
        EI_Browse
    End If
    Exit Function
Failed:
    strMessage = Err.Description
    On Error GoTo CleanupFailed
    EI_CloseConnection objConnection
    MsgBox strMessage, vbCritical, EI_Text("Initialize import", "初始化导入")
    Exit Function
CleanupFailed:
    MsgBox strMessage & vbCrLf & EI_Text("Releasing the connection also failed: ", "释放连接失败：") & Err.Description, vbCritical, EI_Text("Initialize import", "初始化导入")
End Function

Public Function EI_Browse() As Boolean
    On Error GoTo Failed
    Dim strFileName As String
    Dim strMessage As String
    If mblnImportBusy Then Exit Function
    With Application.FileDialog(3)
        .AllowMultiSelect = False
        .Title = EI_Text("Choose Excel file", "选择 Excel 文件")
        .Filters.Clear
        .Filters.Add EI_Text("Excel workbooks", "Excel 工作簿"), "*.xls;*.xlsx;*.xlsm;*.xlsb"
        If .Show = 0 Then Exit Function
        strFileName = .SelectedItems(1)
    End With
    EI_OpenWorkbook strFileName
    EI_Map
    EI_Browse = True
    Exit Function
Failed:
    strMessage = Err.Description
    EI_ReportCleanup
    MsgBox strMessage, vbCritical, EI_Text("Open Excel", "打开 Excel 失败")
End Function

Public Sub EI_OpenWorkbook(ByVal strFileName As String)
    Dim frmImport As Access.Form
    Dim objWorksheet As Object
    Set frmImport = Forms(strIMPORT_FORM_NAME)
    If mblnImportBusy Then Err.Raise vbObjectError + 2123, , EI_Text("An import is running; the file cannot be changed.", "正在导入，不能更换文件。")
    EI_ReleaseExcel
    mblnMappingReady = False
    frmImport!txtFileName = Null
    frmImport!cboSheetName = Null
    frmImport!cboSheetName.RowSource = ""
    Set mobjExcelApplication = CreateObject("Excel.Application")
    mobjExcelApplication.DisplayAlerts = False
    mobjExcelApplication.AutomationSecurity = 3
    Set mobjWorkbook = mobjExcelApplication.Workbooks.Open(strFileName, 0, True)
    mlngExcelDateOffset = 0
    If mobjWorkbook.Date1904 Then mlngExcelDateOffset = 1462
    For Each objWorksheet In mobjWorkbook.Worksheets
        ' Like the old form, only worksheets that contain data can be selected.
        If mobjExcelApplication.WorksheetFunction.CountA(objWorksheet.UsedRange) > 0 Then
            frmImport!cboSheetName.AddItem """" & Replace(objWorksheet.Name, """", """""") & """"
        End If
    Next
    If frmImport!cboSheetName.ListCount = 0 Then
        EI_ReleaseExcel
        Err.Raise vbObjectError + 2103, , EI_Text("The workbook has no worksheet containing data.", "工作簿中没有包含数据的工作表。")
    End If
    frmImport!txtFileName = strFileName
    frmImport!cboSheetName = frmImport!cboSheetName.ItemData(0)
    frmImport!cboSheetName.Enabled = (frmImport!cboSheetName.ListCount > 1)
End Sub

Private Sub EI_ReportCleanup()
    On Error GoTo Failed
    EI_ReleaseExcel
    Exit Sub
Failed:
    MsgBox EI_Text("Closing the Excel instance failed: ", "清理 Excel 实例失败：") & Err.Description, vbExclamation, EI_Text("Release resources", "资源清理")
End Sub

Private Sub EI_Bounds(ByVal objWorksheet As Object, ByRef lngLastRow As Long, ByRef lngLastColumn As Long)
    Dim objFoundCell As Object
    Set objFoundCell = objWorksheet.Cells.Find("*", objWorksheet.Cells(1, 1), -4123, 2, 1, 2, False, False, False)
    If objFoundCell Is Nothing Then Err.Raise vbObjectError + 2103, , EI_Text("The worksheet has no data.", "工作表没有数据。")
    lngLastRow = objFoundCell.Row
    Set objFoundCell = objWorksheet.Cells.Find("*", objWorksheet.Cells(1, 1), -4123, 2, 2, 2, False, False, False)
    lngLastColumn = objFoundCell.Column
End Sub

Private Function EI_HeaderRow(ByVal frmImport As Access.Form) As Long
    Dim dblHeaderNumber As Double
    If Not IsNumeric(Nz(frmImport!txtHeaderRow, "")) Then
        Err.Raise vbObjectError + 2104, , EI_Text("The header row must be a positive whole number.", "标题行必须是正整数。")
    End If
    dblHeaderNumber = CDbl(frmImport!txtHeaderRow)
    If dblHeaderNumber < 1 Or dblHeaderNumber > 1048576 Or dblHeaderNumber <> Fix(dblHeaderNumber) Then
        Err.Raise vbObjectError + 2104, , EI_Text("The header row must be a valid Excel row number.", "标题行必须是有效的 Excel 行号。")
    End If
    EI_HeaderRow = CLng(dblHeaderNumber)
End Function

Public Function EI_Map() As Boolean
    On Error GoTo Failed
    Dim frmImport As Access.Form
    Dim dbsDatabase As DAO.Database
    Dim tdfTargetTable As DAO.TableDef
    Dim fldField As DAO.Field
    Dim objConnection As Object
    Dim objCommand As Object
    Dim objTargetConnection As Object
    Dim objTargetMetadata As Object
    Dim objWorksheet As Object
    Dim varHeaders As Variant
    Dim varValue As Variant
    Dim strHeaderText As String
    Dim strMatchedField As String
    Dim objUsedFields As Object
    Dim lngExcelRow As Long
    Dim lngLastRow As Long
    Dim lngLastColumn As Long
    Dim lngColumnIndex As Long
    Dim blnSqlServer As Boolean
    Dim blnInTransaction As Boolean
    Dim strMessage As String
    Set frmImport = Forms(strIMPORT_FORM_NAME)
    If mblnImportBusy Then Exit Function
    mblnMappingReady = False
    If mobjWorkbook Is Nothing Then Exit Function
    If Len(Nz(frmImport!cboTableName, "")) = 0 Or Len(Nz(frmImport!cboSheetName, "")) = 0 Then Exit Function
    Set dbsDatabase = CurrentDb
    Set tdfTargetTable = dbsDatabase.TableDefs(CStr(frmImport!cboTableName))
    Set objTargetMetadata = EI_TargetMetadata(tdfTargetTable, objTargetConnection, blnSqlServer)
    Set objWorksheet = mobjWorkbook.Worksheets(CStr(frmImport!cboSheetName))
    lngExcelRow = EI_HeaderRow(frmImport)
    EI_Bounds objWorksheet, lngLastRow, lngLastColumn
    If lngExcelRow >= lngLastRow Then Err.Raise vbObjectError + 2105, , EI_Text("There is no data below the header row.", "标题行下方没有数据。")
    If frmImport!sfrDetail.Form.Dirty Then frmImport!sfrDetail.Form.Dirty = False
    Set objConnection = EI_LocalConnection()
    objConnection.BeginTrans
    blnInTransaction = True
    EI_Execute objConnection, "DELETE FROM tblEI_Map"
    frmImport!sfrDetail.Form!cboAccFld.RowSource = ""
    For Each fldField In tdfTargetTable.Fields
        If EI_MetadataAllows(objTargetMetadata, fldField.Name) Then
            ' Second column shows the field caption, as in the old import form.
            frmImport!sfrDetail.Form!cboAccFld.AddItem """" & Replace(fldField.Name, """", """""") & """;""" & Replace(EI_Property(fldField, "Caption"), """", """""") & """"
        End If
    Next
    varHeaders = objWorksheet.Range(objWorksheet.Cells(lngExcelRow, 1), objWorksheet.Cells(lngExcelRow, lngLastColumn)).Value2
    Set objUsedFields = CreateObject("Scripting.Dictionary")
    objUsedFields.CompareMode = vbTextCompare
    Set objCommand = EI_Command(objConnection, "INSERT INTO tblEI_Map (ColIndex,ColLetter,ExcelField,AccessField,SkipColumn) VALUES (?,?,?,?,?)")
    EI_AddParameter objCommand, lngADO_INTEGER, 0
    EI_AddParameter objCommand, lngADO_UNICODE_TEXT, "", 8
    EI_AddParameter objCommand, lngADO_UNICODE_TEXT, "", 255
    EI_AddParameter objCommand, lngADO_UNICODE_TEXT, Null, 64
    EI_AddParameter objCommand, lngADO_INTEGER, -1
    For lngColumnIndex = 1 To lngLastColumn
        If lngLastColumn = 1 Then
            varValue = varHeaders
        Else
            varValue = varHeaders(1, lngColumnIndex)
        End If
        strHeaderText = ""
        If Not IsError(varValue) And Not IsEmpty(varValue) Then strHeaderText = Trim$(CStr(varValue))
        strMatchedField = ""
        For Each fldField In tdfTargetTable.Fields
            If EI_MetadataAllows(objTargetMetadata, fldField.Name) And Len(strHeaderText) > 0 Then
                If StrComp(strHeaderText, fldField.Name, vbTextCompare) = 0 Or StrComp(strHeaderText, EI_Property(fldField, "Caption"), vbTextCompare) = 0 Then
                    If Not objUsedFields.Exists(fldField.Name) Then
                        strMatchedField = fldField.Name
                        objUsedFields.Add strMatchedField, True
                    End If
                    Exit For
                End If
            End If
        Next
        objCommand.Parameters(0).Value = lngColumnIndex
        objCommand.Parameters(1).Value = EI_ColumnLetter(lngColumnIndex)
        objCommand.Parameters(2).Value = Left$(strHeaderText, 255)
        objCommand.Parameters(3).Value = Null
        If Len(strMatchedField) > 0 Then objCommand.Parameters(3).Value = strMatchedField
        objCommand.Parameters(4).Value = CLng(Len(strMatchedField) = 0)
        objCommand.Execute , , lngADO_EXECUTE_NO_RECORDS
    Next
    objConnection.CommitTrans
    blnInTransaction = False
    EI_CloseConnection objConnection
    EI_CloseConnection objTargetConnection
    EI_RefreshMapForm
    mblnMappingReady = True
    EI_Status EI_Format(EI_Text("{0} data rows found. Confirm the mapping; unmatched columns are skipped.", "数据范围：{0} 行；请确认映射，未匹配的列默认跳过。"), lngLastRow - lngExcelRow)
    EI_Map = True
    Exit Function
Failed:
    strMessage = Err.Description
    On Error GoTo CleanupFailed
    If blnInTransaction Then objConnection.RollbackTrans
    EI_CloseConnection objConnection
    EI_CloseConnection objTargetConnection
    MsgBox strMessage, vbExclamation, EI_Text("Field mapping", "生成字段映射")
    Exit Function
CleanupFailed:
    MsgBox strMessage & vbCrLf & EI_Text("Releasing the connection also failed: ", "释放连接时又发生错误：") & Err.Description, vbExclamation, EI_Text("Field mapping", "生成字段映射")
End Function

Private Function EI_CellText(ByVal varValue As Variant) As String
    If IsError(varValue) Then
        EI_CellText = "#EI_EXCEL_ERROR#"
    ElseIf Not IsEmpty(varValue) And Not IsNull(varValue) Then
        EI_CellText = CStr(varValue)
    End If
End Function

Private Function EI_StageBuffer(ByVal objConnection As Object, ByVal strTableName As String) As Object
    Dim rstData As Object
    Set rstData = CreateObject("ADODB.Recordset")
    rstData.CursorLocation = 3
    rstData.Open "SELECT * FROM " & EI_Q(strTableName) & " WHERE 1=0", objConnection, 3, 4
    If Not rstData.Supports(65536) Then Err.Raise vbObjectError + 2140, , EI_Text("The current ADO provider does not support batch updates of staging tables.", "当前 ADO 提供程序不支持暂存表批量更新。")
    Set EI_StageBuffer = rstData
End Function

Private Function EI_WriteStage(ByVal objConnection As Object, ByVal objWorksheet As Object, ByVal lngHeaderRow As Long, ByRef udtColumns() As typImportColumn) As Long
    Dim rstData As Object
    Dim lngSourceColumnIndices() As Long
    Dim varExcelData As Variant
    Dim varCellValues() As Variant
    Dim strValue As String
    Dim lngLastRow As Long
    Dim lngLastColumn As Long
    Dim lngMaxColumn As Long
    Dim lngColumnCount As Long
    Dim lngColumnIndex As Long
    Dim lngExcelRow As Long
    Dim lngStartRow As Long
    Dim lngEndRow As Long
    Dim lngChunkRows As Long
    Dim blnHasData As Boolean
    Dim blnInTransaction As Boolean
    Dim lngErrorNumber As Long
    Dim strErrorText As String
    On Error GoTo Failed
    EI_Bounds objWorksheet, lngLastRow, lngLastColumn
    lngColumnCount = UBound(udtColumns)
    If lngColumnCount > 254 Then Err.Raise vbObjectError + 2106, , EI_Text("At most 254 columns can be mapped at once (one staging column holds the Excel row number).", "一次最多映射 254 列（暂存表保留一列用于 Excel 行号）。")
    ReDim lngSourceColumnIndices(1 To lngColumnCount)
    For lngColumnIndex = 1 To lngColumnCount
        lngSourceColumnIndices(lngColumnIndex) = udtColumns(lngColumnIndex).lngSourceColumnIndex
        If lngSourceColumnIndices(lngColumnIndex) > lngMaxColumn Then lngMaxColumn = lngSourceColumnIndices(lngColumnIndex)
    Next
    EI_DropStage objConnection
    EI_CreateStageTable objConnection, udtColumns, False
    ReDim varCellValues(1 To lngColumnCount)
    lngChunkRows = 100000 \ lngMaxColumn
    If lngChunkRows < 1 Then lngChunkRows = 1
    If lngChunkRows > 2000 Then lngChunkRows = 2000
    objConnection.BeginTrans
    blnInTransaction = True
    For lngStartRow = lngHeaderRow + 1 To lngLastRow Step lngChunkRows
        EI_CheckCancel
        lngEndRow = lngStartRow + lngChunkRows - 1
        If lngEndRow > lngLastRow Then lngEndRow = lngLastRow
        varExcelData = objWorksheet.Range(objWorksheet.Cells(lngStartRow, 1), objWorksheet.Cells(lngEndRow, lngMaxColumn)).Value2
        Set rstData = EI_StageBuffer(objConnection, strRAW_STAGE_TABLE_NAME)
        For lngExcelRow = 1 To lngEndRow - lngStartRow + 1
            blnHasData = False
            For lngColumnIndex = 1 To lngColumnCount
                If lngStartRow = lngEndRow And lngMaxColumn = 1 Then
                    strValue = EI_CellText(varExcelData)
                Else
                    strValue = EI_CellText(varExcelData(lngExcelRow, lngSourceColumnIndices(lngColumnIndex)))
                End If
                If Len(strValue) > 0 Then blnHasData = True
                varCellValues(lngColumnIndex) = Null
                If Len(strValue) > 0 Then varCellValues(lngColumnIndex) = strValue
            Next
            If blnHasData Then
                rstData.AddNew
                rstData.Fields(0).Value = lngStartRow + lngExcelRow - 1
                For lngColumnIndex = 1 To lngColumnCount
                    rstData.Fields(lngColumnIndex).Value = varCellValues(lngColumnIndex)
                Next
                rstData.Update
                EI_WriteStage = EI_WriteStage + 1
            End If
        Next
        rstData.UpdateBatch 3
        rstData.Close
        Set rstData = Nothing
        EI_Status EI_Format(EI_Text("Reading and staging in blocks: {0} / {1} rows", "正在分块读取及批量暂存：{0} / {1} 行"), lngEndRow - lngHeaderRow, lngLastRow - lngHeaderRow)
    Next
    EI_CheckCancel
    objConnection.CommitTrans
    blnInTransaction = False
    Exit Function
Failed:
    lngErrorNumber = Err.Number
    strErrorText = Err.Description
    On Error GoTo CleanupFailed
    If Not rstData Is Nothing Then
        If rstData.State <> 0 Then rstData.Close
    End If
    If blnInTransaction Then objConnection.RollbackTrans
    On Error GoTo 0
    Err.Raise lngErrorNumber, "ExcelImport.Stage", strErrorText
CleanupFailed:
    Err.Raise lngErrorNumber, "ExcelImport.Stage", strErrorText & EI_Text("; closing or rolling back staging failed: ", "；关闭或回滚暂存失败：") & Err.Description
End Function

Public Function EI_Convert(ByVal varValue As Variant, ByVal lngFieldType As Long, ByVal lngDateOffset As Long) As Variant
    Dim strCellText As String
    Dim dblNumber As Double
    EI_Convert = Null
    If IsNull(varValue) Or IsEmpty(varValue) Then Exit Function
    strCellText = CStr(varValue)
    If Len(strCellText) = 0 Then Exit Function
    If strCellText = "#EI_EXCEL_ERROR#" Then Err.Raise vbObjectError + 2107, , EI_Text("The Excel cell contains a formula error.", "Excel 单元格含公式错误。")
    Select Case lngFieldType
        Case dbText, dbMemo
            EI_Convert = strCellText
        Case dbBoolean
            Select Case LCase$(Trim$(strCellText))
                Case "true", "yes", "是", "真", "1", "-1"
                    EI_Convert = True
                Case "false", "no", "否", "假", "0"
                    EI_Convert = False
                Case Else
                    Err.Raise vbObjectError + 2108, , EI_Text("Invalid Yes/No value.", "无效的是/否值。")
            End Select
        Case dbDate
            If IsNumeric(strCellText) Then
                dblNumber = CDbl(strCellText) + lngDateOffset
                If dblNumber < 1 Or dblNumber >= 2958466# Then Err.Raise vbObjectError + 2109, , EI_Text("The Excel date serial number is out of range.", "Excel 日期序号超出范围。")
                If lngDateOffset = 0 And Fix(dblNumber) = 60 Then Err.Raise vbObjectError + 2109, , EI_Text("Excel's 1900-02-29 is not a valid date.", "Excel 的 1900-02-29 并非有效日期。")
                If lngDateOffset = 0 And dblNumber < 60 Then dblNumber = dblNumber + 1
                EI_Convert = CDate(dblNumber)
            Else
                EI_Convert = CDate(Replace(strCellText, ".", "-"))
            End If
        Case dbByte, dbInteger, dbLong
            dblNumber = CDbl(strCellText)
            If dblNumber <> Fix(dblNumber) Then Err.Raise vbObjectError + 2110, , EI_Text("Whole numbers cannot contain decimals.", "整数不能包含小数。")
            Select Case lngFieldType
                Case dbByte: EI_Convert = CByte(dblNumber)
                Case dbInteger: EI_Convert = CInt(dblNumber)
                Case dbLong: EI_Convert = CLng(dblNumber)
            End Select
        Case dbCurrency: EI_Convert = CCur(strCellText)
        Case dbSingle: EI_Convert = CSng(strCellText)
        Case dbDouble: EI_Convert = CDbl(strCellText)
        Case 20: EI_Convert = CDec(strCellText)
        Case 16
            EI_Convert = CDec(strCellText)
            If EI_Convert <> Fix(EI_Convert) Then Err.Raise vbObjectError + 2110, , EI_Text("Whole numbers cannot contain decimals.", "整数不能包含小数。")
            If EI_Convert < CDec("-9223372036854775808") Or EI_Convert > CDec("9223372036854775807") Then
                Err.Raise vbObjectError + 2110, , EI_Text("The value is outside the bigint range.", "数值超出 bigint 范围。")
            End If
        Case Else
            Err.Raise vbObjectError + 2111, , EI_Text("Unsupported target field type.", "不支持的目标字段类型。")
    End Select
End Function

Public Function EI_Reason(ByVal varValue As Variant, ByVal lngFieldType As Long, ByVal lngFieldSize As Long, ByVal blnRequired As Boolean, ByVal lngDateOffset As Long) As String
    On Error GoTo InvalidValue
    Dim varConvertedValue As Variant
    varConvertedValue = EI_Convert(varValue, lngFieldType, lngDateOffset)
    If IsNull(varConvertedValue) Then
        If blnRequired Then EI_Reason = EI_Text("Required field is empty.", "必填字段为空。")
    ElseIf lngFieldType = dbText Then
        If Len(CStr(varConvertedValue)) > lngFieldSize Then EI_Reason = EI_Format(EI_Text("Text is longer than the field size ({0}).", "文本长度超过字段限制（{0}）。"), lngFieldSize)
    End If
    Exit Function
InvalidValue:
    EI_Reason = EI_Text("Type conversion failed: ", "类型转换失败：") & Err.Description
End Function

Private Function EI_StartBatch(ByVal objConnection As Object, ByVal frmImport As Access.Form) As Long
    Dim objCommand As Object
    Dim rstData As Object
    Set objCommand = EI_Command(objConnection, "INSERT INTO tblEI_Batches (StartedAt,FileName,SheetName,TargetTable,Status) VALUES (?,?,?,?,?)")
    EI_AddParameter objCommand, lngADO_TIMESTAMP, Now()
    EI_AddParameter objCommand, lngADO_UNICODE_LONG_TEXT, CStr(frmImport!txtFileName)
    EI_AddParameter objCommand, lngADO_UNICODE_TEXT, CStr(frmImport!cboSheetName), 255
    EI_AddParameter objCommand, lngADO_UNICODE_TEXT, CStr(frmImport!cboTableName), 64
    EI_AddParameter objCommand, lngADO_UNICODE_TEXT, EI_Text("Running", "进行中"), 32
    objCommand.Execute , , lngADO_EXECUTE_NO_RECORDS
    Set rstData = objConnection.Execute("SELECT @@IDENTITY AS BatchID")
    EI_StartBatch = CLng(rstData!BatchID)
    rstData.Close
End Function

Private Sub EI_FinishBatch(ByVal objConnection As Object, ByVal lngBatchID As Long, ByVal strStatus As String, ByVal lngSourceRows As Long, ByVal varImportedRows As Variant, ByVal lngRejectedRows As Long, ByVal strDetails As String)
    Dim objCommand As Object
    Dim varAffectedRows As Variant
    Set objCommand = EI_Command(objConnection, "UPDATE tblEI_Batches SET FinishedAt=?,Status=?,SourceRows=?,ImportedRows=?,RejectedRows=?,Details=? WHERE BatchID=?")
    EI_AddParameter objCommand, lngADO_TIMESTAMP, Now()
    EI_AddParameter objCommand, lngADO_UNICODE_TEXT, strStatus, 32
    EI_AddParameter objCommand, lngADO_INTEGER, lngSourceRows
    EI_AddParameter objCommand, lngADO_INTEGER, varImportedRows
    EI_AddParameter objCommand, lngADO_INTEGER, lngRejectedRows
    EI_AddParameter objCommand, lngADO_UNICODE_LONG_TEXT, strDetails
    EI_AddParameter objCommand, lngADO_INTEGER, lngBatchID
    objCommand.Execute varAffectedRows, , lngADO_EXECUTE_NO_RECORDS
    If CLng(varAffectedRows) <> 1 Then Err.Raise vbObjectError + 2135, , EI_Text("Unexpected row count while updating the batch log.", "批次日志更新行数异常。")
End Sub

Private Sub EI_LogError(ByVal objConnection As Object, ByVal lngBatchID As Long, ByVal lngExcelRow As Long, ByVal varFieldName As Variant, ByVal varValue As Variant, ByVal strReason As String)
    Dim objCommand As Object
    Set objCommand = EI_Command(objConnection, "INSERT INTO tblEI_Errors (BatchID,ExcelRow,FieldName,SourceValue,Reason) VALUES (?,?,?,?,?)")
    EI_AddParameter objCommand, lngADO_INTEGER, lngBatchID
    EI_AddParameter objCommand, lngADO_INTEGER, lngExcelRow
    EI_AddParameter objCommand, lngADO_UNICODE_TEXT, varFieldName, 64
    EI_AddParameter objCommand, lngADO_UNICODE_LONG_TEXT, varValue
    EI_AddParameter objCommand, lngADO_UNICODE_LONG_TEXT, strReason
    objCommand.Execute , , lngADO_EXECUTE_NO_RECORDS
End Sub

Private Sub EI_SetBusy(ByVal blnBusy As Boolean)
    Dim frmImport As Access.Form
    mblnImportBusy = blnBusy
    DoCmd.Hourglass blnBusy
    If Not CurrentProject.AllForms(strIMPORT_FORM_NAME).IsLoaded Then Exit Sub
    Set frmImport = Forms(strIMPORT_FORM_NAME)
    frmImport!cmdBrowse.Enabled = Not blnBusy
    frmImport!cboSheetName.Enabled = Not blnBusy And frmImport!cboSheetName.ListCount > 1
    frmImport!cboTableName.Enabled = Not blnBusy And Len(Nz(frmImport.OpenArgs, "")) = 0
    frmImport!txtHeaderRow.Enabled = Not blnBusy
    frmImport!chkSkipInvalid.Enabled = Not blnBusy
    frmImport!cmdOK.Enabled = Not blnBusy
    frmImport!cmdErrors.Enabled = Not blnBusy
    frmImport!cmdHistory.Enabled = Not blnBusy
    frmImport!cmdLanguage.Enabled = Not blnBusy
    frmImport!sfrDetail.Enabled = Not blnBusy
    frmImport!cmdCancel.Caption = IIf(blnBusy, EI_Text("Cancel import", "取消导入"), EI_Text("Close", "关闭"))
End Sub

Private Sub EI_BuildColumns(ByVal varFieldMappings As Variant, ByVal objTargetMetadata As Object, ByRef udtColumns() As typImportColumn)
    Dim objUsedFields As Object
    Dim objFieldInfo As Object
    Dim varFieldKey As Variant
    Dim strFieldName As String
    Dim lngColumnIndex As Long
    Dim lngColumnCount As Long
    lngColumnCount = UBound(varFieldMappings, 2) + 1
    If lngColumnCount > 254 Then Err.Raise vbObjectError + 2106, , EI_Text("At most 254 columns can be mapped at once.", "一次最多映射 254 列。")
    ReDim udtColumns(1 To lngColumnCount)
    Set objUsedFields = CreateObject("Scripting.Dictionary")
    objUsedFields.CompareMode = vbTextCompare
    For lngColumnIndex = 1 To lngColumnCount
        strFieldName = CStr(Nz(varFieldMappings(1, lngColumnIndex - 1), ""))
        If Len(strFieldName) = 0 Then Err.Raise vbObjectError + 2116, , EI_Text("Every selected column must have a target field.", "选中的列必须指定目标字段。")
        If Not EI_MetadataAllows(objTargetMetadata, strFieldName) Then Err.Raise vbObjectError + 2117, , EI_Text("This field cannot be imported or is generated by the database: ", "字段不支持导入或是数据库自动生成字段：") & strFieldName
        If objUsedFields.Exists(strFieldName) Then Err.Raise vbObjectError + 2118, , EI_Text("A target field cannot be mapped more than once: ", "目标字段不能重复映射：") & strFieldName
        objUsedFields.Add strFieldName, True
        Set objFieldInfo = objTargetMetadata(strFieldName)
        With udtColumns(lngColumnIndex)
            .strFieldName = CStr(objFieldInfo("Name"))
            .lngSourceColumnIndex = CLng(varFieldMappings(0, lngColumnIndex - 1))
            If .lngSourceColumnIndex < 1 Or .lngSourceColumnIndex > 16384 Then Err.Raise vbObjectError + 2136, , EI_Text("Invalid source column number.", "源列号无效。")
            .lngFieldType = CLng(objFieldInfo("FieldType"))
            .lngFieldSize = CLng(objFieldInfo("Size"))
            .bytPrecision = CByte(objFieldInfo("Precision"))
            .bytScale = CByte(objFieldInfo("Scale"))
            .strSqlType = CStr(objFieldInfo("SqlType"))
            .blnRequired = CBool(objFieldInfo("Required"))
        End With
    Next
    For Each varFieldKey In objTargetMetadata.Keys
        Set objFieldInfo = objTargetMetadata(varFieldKey)
        If CBool(objFieldInfo("Writable")) And CBool(objFieldInfo("Required")) And Not CBool(objFieldInfo("Default")) Then
            If Not objUsedFields.Exists(CStr(varFieldKey)) Then Err.Raise vbObjectError + 2119, , EI_Text("A required field without a default value is not mapped: ", "必填且没有默认值的字段未映射：") & CStr(varFieldKey)
        End If
    Next
End Sub

Private Function EI_ValidateStage(ByVal objConnection As Object, ByVal lngBatchID As Long, ByRef udtColumns() As typImportColumn) As Long
    On Error GoTo Failed
    Dim rstData As Object
    Dim varStageRows As Variant
    Dim lngColumnIndex As Long
    Dim lngRowIndex As Long
    Dim lngCheckedRows As Long
    Dim lngErrorNumber As Long
    Dim strReason As String
    Dim strMessage As String
    Dim blnInvalidRow As Boolean
    Dim blnInTransaction As Boolean
    Set rstData = objConnection.Execute("SELECT * FROM tblEI_Stage ORDER BY EI_Row")
    objConnection.BeginTrans
    blnInTransaction = True
    Do Until rstData.EOF
        EI_CheckCancel
        varStageRows = rstData.GetRows(1000)
        For lngRowIndex = 0 To UBound(varStageRows, 2)
            blnInvalidRow = False
            For lngColumnIndex = 1 To UBound(udtColumns)
                strReason = EI_Reason(varStageRows(lngColumnIndex, lngRowIndex), udtColumns(lngColumnIndex).lngFieldType, udtColumns(lngColumnIndex).lngFieldSize, udtColumns(lngColumnIndex).blnRequired, mlngExcelDateOffset)
                If Len(strReason) > 0 Then
                    EI_LogError objConnection, lngBatchID, CLng(varStageRows(0, lngRowIndex)), udtColumns(lngColumnIndex).strFieldName, varStageRows(lngColumnIndex, lngRowIndex), strReason
                    blnInvalidRow = True
                End If
            Next
            If blnInvalidRow Then EI_ValidateStage = EI_ValidateStage + 1
            lngCheckedRows = lngCheckedRows + 1
        Next
        EI_Status EI_Format(EI_Text("Validating: {0} rows; {1} rows with errors", "正在校验：{0} 行；发现 {1} 行错误"), lngCheckedRows, EI_ValidateStage)
    Loop
    rstData.Close
    objConnection.CommitTrans
    blnInTransaction = False
    Exit Function
Failed:
    lngErrorNumber = Err.Number
    strMessage = Err.Description
    On Error GoTo RecoveryFailed
    If Not rstData Is Nothing Then
        If rstData.State <> 0 Then rstData.Close
    End If
    If blnInTransaction Then objConnection.RollbackTrans
    On Error GoTo 0
    Err.Raise lngErrorNumber, "ExcelImport.Validation", strMessage
RecoveryFailed:
    Err.Raise lngErrorNumber, "ExcelImport.Validation", strMessage & EI_Text("; restoring the validation transaction failed: ", "；恢复校验事务失败：") & Err.Description
End Function

Private Function EI_ValidRows(ByVal objConnection As Object, ByVal lngBatchID As Long) As Object
    Set EI_ValidRows = objConnection.Execute("SELECT s.* FROM tblEI_Stage AS s WHERE NOT EXISTS (SELECT * FROM tblEI_Errors AS e WHERE e.BatchID=" & lngBatchID & " AND e.ExcelRow=s.EI_Row) ORDER BY s.EI_Row")
End Function

Private Sub EI_WriteTypedStage(ByVal objConnection As Object, ByVal lngBatchID As Long, ByRef udtColumns() As typImportColumn)
    On Error GoTo Failed
    Dim rstData As Object
    Dim rstStageBuffer As Object
    Dim varStageRows As Variant
    Dim lngColumnIndex As Long
    Dim lngRowIndex As Long
    Dim lngCheckedRows As Long
    Dim lngErrorNumber As Long
    Dim strMessage As String
    Dim blnInTransaction As Boolean
    EI_CreateStageTable objConnection, udtColumns, True
    Set rstData = EI_ValidRows(objConnection, lngBatchID)
    objConnection.BeginTrans
    blnInTransaction = True
    Do Until rstData.EOF
        EI_CheckCancel
        varStageRows = rstData.GetRows(1000)
        Set rstStageBuffer = EI_StageBuffer(objConnection, strTYPED_STAGE_TABLE_NAME)
        For lngRowIndex = 0 To UBound(varStageRows, 2)
            rstStageBuffer.AddNew
            rstStageBuffer.Fields(0).Value = CLng(varStageRows(0, lngRowIndex))
            For lngColumnIndex = 1 To UBound(udtColumns)
                rstStageBuffer.Fields(lngColumnIndex).Value = EI_Convert(varStageRows(lngColumnIndex, lngRowIndex), udtColumns(lngColumnIndex).lngFieldType, mlngExcelDateOffset)
            Next
            rstStageBuffer.Update
            lngCheckedRows = lngCheckedRows + 1
        Next
        rstStageBuffer.UpdateBatch 3
        rstStageBuffer.Close
        Set rstStageBuffer = Nothing
        EI_Status EI_Format(EI_Text("Building typed staging data: {0} rows", "正在生成类型化暂存数据：{0} 行"), lngCheckedRows)
    Loop
    rstData.Close
    Set rstData = Nothing
    EI_CheckCancel
    objConnection.CommitTrans
    blnInTransaction = False
    Exit Sub
Failed:
    lngErrorNumber = Err.Number
    strMessage = Err.Description
    On Error GoTo CleanupFailed
    If Not rstData Is Nothing Then
        If rstData.State <> 0 Then rstData.Close
    End If
    If Not rstStageBuffer Is Nothing Then
        If rstStageBuffer.State <> 0 Then rstStageBuffer.Close
    End If
    If blnInTransaction Then objConnection.RollbackTrans
    On Error GoTo 0
    Err.Raise lngErrorNumber, "ExcelImport.TypedStage", strMessage
CleanupFailed:
    Err.Raise lngErrorNumber, "ExcelImport.TypedStage", strMessage & EI_Text("; releasing staging resources failed: ", "；释放暂存资源失败：") & Err.Description
End Sub

Private Function EI_AdoType(ByVal lngFieldType As Long) As Long
    Select Case lngFieldType
        Case dbBoolean: EI_AdoType = lngADO_BOOLEAN
        Case dbByte: EI_AdoType = 17
        Case dbInteger: EI_AdoType = 2
        Case dbLong: EI_AdoType = lngADO_INTEGER
        Case dbCurrency: EI_AdoType = 6
        Case dbSingle: EI_AdoType = 4
        Case dbDouble: EI_AdoType = 5
        Case dbDate: EI_AdoType = lngADO_TIMESTAMP
        Case dbText: EI_AdoType = lngADO_UNICODE_TEXT
        Case dbMemo: EI_AdoType = lngADO_UNICODE_LONG_TEXT
        Case 16: EI_AdoType = 20
        Case 20: EI_AdoType = lngADO_DECIMAL
        Case Else: Err.Raise vbObjectError + 2111, , EI_Text("Unsupported target field type.", "不支持的目标字段类型。")
    End Select
End Function

Private Sub EI_ServerUpload(ByVal objLocalConnection As Object, ByVal objTargetConnection As Object, ByVal lngBatchID As Long, ByRef udtColumns() As typImportColumn)
    Dim rstData As Object
    Dim objCommand As Object
    Dim varStageRows As Variant
    Dim varValue As Variant
    Dim strTargetFields As String
    Dim strParameterMarkers As String
    Dim strSql As String
    Dim lngColumnIndex As Long
    Dim lngRowIndex As Long
    Dim lngBatchSize As Long
    Dim lngUploadedRows As Long
    strTargetFields = "EI_Row"
    strParameterMarkers = "?"
    For lngColumnIndex = 1 To UBound(udtColumns)
        strTargetFields = strTargetFields & ",T" & lngColumnIndex
        strParameterMarkers = strParameterMarkers & ",?"
    Next
    EI_CreateServerStage objTargetConnection, udtColumns
    lngBatchSize = 2000 \ (UBound(udtColumns) + 1)
    If lngBatchSize > 500 Then lngBatchSize = 500
    Set rstData = EI_ValidRows(objLocalConnection, lngBatchID)
    Do Until rstData.EOF
        EI_CheckCancel
        varStageRows = rstData.GetRows(lngBatchSize)
        Set objCommand = EI_Command(objTargetConnection, "")
        strSql = "INSERT INTO #EI_Stage (" & strTargetFields & ") VALUES "
        For lngRowIndex = 0 To UBound(varStageRows, 2)
            If lngRowIndex > 0 Then strSql = strSql & ","
            strSql = strSql & "(" & strParameterMarkers & ")"
            EI_AddParameter objCommand, lngADO_INTEGER, CLng(varStageRows(0, lngRowIndex))
            For lngColumnIndex = 1 To UBound(udtColumns)
                varValue = EI_Convert(varStageRows(lngColumnIndex, lngRowIndex), udtColumns(lngColumnIndex).lngFieldType, mlngExcelDateOffset)
                EI_AddParameter objCommand, EI_AdoType(udtColumns(lngColumnIndex).lngFieldType), varValue, udtColumns(lngColumnIndex).lngFieldSize, udtColumns(lngColumnIndex).bytPrecision, udtColumns(lngColumnIndex).bytScale
            Next
            lngUploadedRows = lngUploadedRows + 1
        Next
        objCommand.CommandText = strSql
        objCommand.Execute , , lngADO_EXECUTE_NO_RECORDS
        EI_Status EI_Format(EI_Text("Uploading SQL Server staging rows in batches: {0}", "正在分批上传 SQL Server 暂存表：{0} 行"), lngUploadedRows)
    Loop
    rstData.Close
End Sub

Private Function EI_ServerAppend(ByVal objTargetConnection As Object, ByVal strTableName As String, ByVal strTargetFields As String, ByVal strStageFields As String) As Long
    Dim rstData As Object
    Dim fldField As Object
    EI_ServerAppend = -1
    Set rstData = objTargetConnection.Execute("SET NOCOUNT ON; DECLARE @EI_Count int; INSERT INTO " & strTableName & " (" & strTargetFields & ") SELECT " & strStageFields & " FROM #EI_Stage; SET @EI_Count=@@ROWCOUNT; SELECT @EI_Count AS EI_Appended;")
    Do Until rstData Is Nothing
        If rstData.State <> 0 Then
            For Each fldField In rstData.Fields
                If fldField.Name = "EI_Appended" Then
                    If Not rstData.EOF Then EI_ServerAppend = CLng(fldField.Value)
                End If
            Next
        End If
        Set rstData = rstData.NextRecordset
    Loop
    If EI_ServerAppend < 0 Then Err.Raise vbObjectError + 2137, , EI_Text("SQL Server did not return the appended row count; the batch was not committed.", "SQL Server 未返回追加行数，拒绝提交。")
End Function

Private Function EI_AdoErrors(ByVal objConnection As Object) As String
    Dim objAdoError As Object
    If objConnection Is Nothing Then Exit Function
    For Each objAdoError In objConnection.Errors
        EI_AdoErrors = EI_AdoErrors & vbCrLf & "ADO " & objAdoError.NativeError & " / " & objAdoError.SQLState & "：" & objAdoError.Description
    Next
End Function

Private Function EI_RollbackError(ByVal objConnection As Object) As String
    On Error GoTo Failed
    objConnection.RollbackTrans
    Exit Function
Failed:
    EI_RollbackError = EI_Text("Rollback failed: ", "回滚失败：") & Err.Description & EI_AdoErrors(objConnection)
End Function

Public Function EI_Run(Optional ByVal blnSilent As Boolean = False) As Boolean
    On Error GoTo Failed
    Dim dbsDatabase As DAO.Database
    Dim tdfTargetTable As DAO.TableDef
    Dim objLocalConnection As Object
    Dim objTargetConnection As Object
    Dim objTargetMetadata As Object
    Dim rstData As Object
    Dim varFieldMappings As Variant
    Dim varAffectedRows As Variant
    Dim udtColumns() As typImportColumn
    Dim frmImport As Access.Form
    Dim objWorksheet As Object
    Dim strTargetFields As String
    Dim strStageFields As String
    Dim strSql As String
    Dim strStatus As String
    Dim strMessage As String
    Dim strRollbackError As String
    Dim lngBatchID As Long
    Dim lngSourceRows As Long
    Dim lngImportedRows As Long
    Dim lngRejectedRows As Long
    Dim lngColumnIndex As Long
    Dim lngHeaderRow As Long
    Dim lngErrorNumber As Long
    Dim blnInTransaction As Boolean
    Dim blnSqlServer As Boolean
    Dim blnCommitStarted As Boolean
    Dim blnCompleted As Boolean
    Dim blnCommitUnknown As Boolean
    Dim sngStartTime As Single
    If mblnImportBusy Then Exit Function
    sngStartTime = Timer()
    Set frmImport = Forms(strIMPORT_FORM_NAME)
    If mobjWorkbook Is Nothing Then Err.Raise vbObjectError + 2112, , EI_Text("Select an Excel file first.", "请先选择 Excel 文件。")
    If Not mblnMappingReady Then Err.Raise vbObjectError + 2138, , EI_Text("Select a valid target table and worksheet, then rebuild the field mapping.", "请先选择有效的目标表、工作表并重新生成字段映射。")
    If frmImport!sfrDetail.Form.Dirty Then frmImport!sfrDetail.Form.Dirty = False
    Set dbsDatabase = CurrentDb
    Set tdfTargetTable = dbsDatabase.TableDefs(CStr(Nz(frmImport!cboTableName, "")))
    Set objLocalConnection = EI_LocalConnection()
    Set objTargetMetadata = EI_TargetMetadata(tdfTargetTable, objTargetConnection, blnSqlServer)
    Set rstData = objLocalConnection.Execute("SELECT ColIndex,AccessField FROM tblEI_Map WHERE SkipColumn=False ORDER BY ColIndex")
    If rstData.EOF Then Err.Raise vbObjectError + 2115, , EI_Text("Select at least one column to import.", "至少需要选择一列。")
    varFieldMappings = rstData.GetRows()
    rstData.Close
    Set rstData = Nothing
    EI_BuildColumns varFieldMappings, objTargetMetadata, udtColumns
    If Not blnSqlServer Then EI_CloseConnection objTargetConnection
    lngHeaderRow = EI_HeaderRow(frmImport)
    Set objWorksheet = mobjWorkbook.Worksheets(CStr(frmImport!cboSheetName))
    mblnCancelRequested = False
    EI_SetBusy True
    lngBatchID = EI_StartBatch(objLocalConnection, frmImport)
    frmImport!txtBatchID = lngBatchID
    lngSourceRows = EI_WriteStage(objLocalConnection, objWorksheet, lngHeaderRow, udtColumns)
    If lngSourceRows = 0 Then Err.Raise vbObjectError + 2120, , EI_Text("The mapped columns contain no non-empty data rows.", "映射列中没有非空数据行。")
    Set rstData = objLocalConnection.Execute("SELECT Count(*) AS N FROM tblEI_Stage")
    If CLng(rstData!N) <> lngSourceRows Then Err.Raise vbObjectError + 2139, , EI_Text("The raw staging row count is inconsistent; import stopped.", "原始暂存行数异常，已停止导入。")
    rstData.Close
    Set rstData = Nothing
    lngRejectedRows = EI_ValidateStage(objLocalConnection, lngBatchID, udtColumns)
    For lngColumnIndex = 1 To UBound(udtColumns)
        If lngColumnIndex > 1 Then
            strTargetFields = strTargetFields & ","
            strStageFields = strStageFields & ","
        End If
        strTargetFields = strTargetFields & EI_Q(udtColumns(lngColumnIndex).strFieldName)
        strStageFields = strStageFields & "T" & lngColumnIndex
    Next
    If lngRejectedRows > 0 And Not Nz(frmImport!chkSkipInvalid, False) Then
        strStatus = EI_Text("Validation failed", "校验未通过")
        strMessage = EI_Format(EI_Text("{0} rows failed validation; nothing was written to the target table. See Error details.", "有 {0} 行未通过校验；目标表未写入任何数据。请查看错误明细。"), lngRejectedRows)
    Else
        EI_CheckCancel
        If Not blnSqlServer Then
            EI_WriteTypedStage objLocalConnection, lngBatchID, udtColumns
            ' ACE must read the newly created staging table on its creating connection.
            Set objTargetConnection = objLocalConnection
        End If
        EI_CheckCancel
        objTargetConnection.BeginTrans
        blnInTransaction = True
        If blnSqlServer Then
            EI_ServerUpload objLocalConnection, objTargetConnection, lngBatchID, udtColumns
            EI_CheckCancel
            EI_Status EI_Text("Appending on SQL Server in one batch; this step cannot be cancelled...", "正在 SQL Server 上执行批量追加；此步骤结束前无法取消...")
            lngImportedRows = EI_ServerAppend(objTargetConnection, EI_ServerTable(tdfTargetTable.SourceTableName), strTargetFields, strStageFields)
        Else
            EI_Status EI_Text("Appending to Access in one ADO batch; this step cannot be cancelled...", "正在通过 ADO 执行 Access 批量追加；此步骤结束前无法取消...")
            strSql = "INSERT INTO " & EI_Q(tdfTargetTable.Name) & " (" & strTargetFields & ") SELECT " & strStageFields & " FROM tblEI_TypedStage"
            objTargetConnection.Execute strSql, varAffectedRows, lngADO_EXECUTE_NO_RECORDS
            lngImportedRows = CLng(varAffectedRows)
        End If
        If lngImportedRows <> lngSourceRows - lngRejectedRows Then Err.Raise vbObjectError + 2121, , EI_Text("The appended row count does not match validation; the batch was not committed.", "追加行数与校验结果不符，已拒绝提交。")
        blnCommitStarted = True
        objTargetConnection.CommitTrans
        blnCommitStarted = False
        blnInTransaction = False
        blnCompleted = True
        strStatus = EI_Text("Completed", "完成")
        strMessage = EI_Format(EI_Text("Appended {0} rows; skipped {1} rows. Batch: {2}. Elapsed: {3} s.", "批量追加 {0} 行；跳过 {1} 行。批次：{2}。用时 {3} 秒。"), lngImportedRows, lngRejectedRows, lngBatchID, Format$(Timer() - sngStartTime, "0.00"))
    End If
    EI_FinishBatch objLocalConnection, lngBatchID, strStatus, lngSourceRows, lngImportedRows, lngRejectedRows, strMessage
    EI_DropStage objLocalConnection
    If blnSqlServer Then EI_CloseConnection objTargetConnection
    EI_CloseConnection objLocalConnection
    EI_SetBusy False
    EI_Status strMessage
    EI_Run = blnCompleted
    If Not blnSilent Then MsgBox strMessage, vbInformation, EI_Text("Excel import", "Excel 导入")
    Exit Function
Failed:
    lngErrorNumber = Err.Number
    strMessage = "#" & lngErrorNumber & ": " & Err.Description & EI_AdoErrors(objTargetConnection) & EI_AdoErrors(objLocalConnection)
    On Error GoTo RecoveryFailed
    If blnInTransaction Then
        strRollbackError = EI_RollbackError(objTargetConnection)
        If Not blnCommitStarted And Len(strRollbackError) = 0 Then lngImportedRows = 0
        blnInTransaction = False
    End If
    If lngBatchID > 0 Then
        strStatus = EI_Text("Failed", "失败")
        If lngErrorNumber = vbObjectError + 2101 Then strStatus = EI_Text("Cancelled", "已取消")
        If blnCommitStarted Or Len(strRollbackError) > 0 Then
            blnCommitUnknown = True
            strStatus = EI_Text("Commit status unknown", "提交状态待核实")
            strMessage = EI_Text("Commit or rollback failed, so it is unknown whether the target table was committed. Check it before retrying.", "提交或回滚异常，不能确认目标表是否已提交；请先核实，勿直接重试。") & vbCrLf & strMessage
            If Len(strRollbackError) > 0 Then strMessage = strMessage & vbCrLf & strRollbackError
        ElseIf lngImportedRows > 0 Then
            strStatus = EI_Text("Committed; cleanup failed", "已提交，后处理失败")
            strMessage = EI_Format(EI_Text("{0} rows were committed to the target table; do not import again. Post-processing error: ", "目标表已提交 {0} 行，请勿重复导入。后处理错误："), lngImportedRows) & strMessage
        Else
            strMessage = strMessage & vbCrLf & EI_Text("Nothing from this batch was committed. Unique index, validation rule and other constraint errors roll back the whole batch.", "本批次未向目标表提交数据。唯一索引、表验证规则等约束错误会使整批回滚。")
        End If
        EI_LogError objLocalConnection, lngBatchID, 0, Null, Null, strMessage
        If blnCommitUnknown Then
            EI_FinishBatch objLocalConnection, lngBatchID, strStatus, lngSourceRows, Null, lngRejectedRows, strMessage
        Else
            EI_FinishBatch objLocalConnection, lngBatchID, strStatus, lngSourceRows, lngImportedRows, lngRejectedRows, strMessage
        End If
    End If
    If Not rstData Is Nothing Then
        If rstData.State <> 0 Then rstData.Close
    End If
    If Not objLocalConnection Is Nothing Then EI_DropStage objLocalConnection
    If blnSqlServer Then EI_CloseConnection objTargetConnection
    EI_CloseConnection objLocalConnection
    EI_SetBusy False
    EI_Status strMessage
    If Not blnSilent Then MsgBox strMessage, vbExclamation, EI_Text("Import not completed", "导入未完成")
    Exit Function
RecoveryFailed:
    EI_SetBusy False
    If blnSilent Then
        Err.Raise lngErrorNumber, "ExcelImport.Recovery", strMessage & EI_Text("; recovery failed: ", "；恢复失败：") & Err.Description
    Else
        MsgBox strMessage & vbCrLf & EI_Text("Rollback or cleanup also failed: ", "回滚或清理时又发生错误：") & Err.Description & vbCrLf & EI_Text("Check the target table and import history to confirm the commit status before retrying.", "请检查目标表及导入历史，确认提交状态后再重试。"), vbCritical, EI_Text("Check the database", "需要检查数据库")
    End If
End Function

Public Function EI_Errors() As Boolean
    On Error GoTo Failed
    Dim lngBatchID As Long
    lngBatchID = Nz(Forms(strIMPORT_FORM_NAME)!txtBatchID, 0)
    If lngBatchID = 0 Then Err.Raise vbObjectError + 2122, , EI_Text("There is no import batch yet.", "尚无导入批次。")
    If CurrentProject.AllForms("frmExcelImportErrors").IsLoaded Then DoCmd.Close acForm, "frmExcelImportErrors"
    DBEngine.Idle dbRefreshCache
    DoCmd.OpenForm "frmExcelImportErrors", acFormDS, , "BatchID=" & lngBatchID, acFormReadOnly
    EI_Errors = True
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, EI_Text("Error details", "错误明细")
End Function

Public Function EI_History() As Boolean
    On Error GoTo Failed
    If CurrentProject.AllForms("frmExcelImportHistory").IsLoaded Then DoCmd.Close acForm, "frmExcelImportHistory"
    DBEngine.Idle dbRefreshCache
    DoCmd.OpenForm "frmExcelImportHistory", acFormDS, , , acFormReadOnly
    EI_History = True
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, EI_Text("Import history", "导入历史")
End Function
