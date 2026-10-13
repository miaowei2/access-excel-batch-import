Option Compare Database
Option Explicit

Public Const strIMPORT_FORM_NAME As String = "frmExcelImport"
Public Const strIMPORT_MAP_FORM_NAME As String = "frmExcelImportMap"
Public Const strMAP_TABLE_NAME As String = "tblEI_Map"
Public Const strRAW_STAGE_TABLE_NAME As String = "tblEI_Stage"
Public Const strBATCH_TABLE_NAME As String = "tblEI_Batches"
Public Const strERROR_TABLE_NAME As String = "tblEI_Errors"
Public Const strTYPED_STAGE_TABLE_NAME As String = "tblEI_TypedStage"
Private Const lngFORM_SECTION_HEADER As Long = 1
Private Const lngFORM_SECTION_FOOTER As Long = 2

Private mlngFormBackColor As Long
Private mlngBlueColor As Long
Private mlngLightBlueColor As Long
Private mlngTextColor As Long
Private mlngHeadingColor As Long
Private mlngMutedTextColor As Long
Private mlngBorderColor As Long
Private mlngImportButtonBackColor As Long

Private Sub EI_InitializeColors()
    mlngFormBackColor = RGB(247, 247, 255)
    mlngBlueColor = RGB(13, 110, 253)
    mlngLightBlueColor = RGB(231, 241, 255)
    mlngTextColor = RGB(76, 82, 88)
    mlngHeadingColor = RGB(33, 37, 41)
    mlngMutedTextColor = RGB(108, 117, 125)
    mlngBorderColor = RGB(222, 226, 230)
    mlngImportButtonBackColor = RGB(0, 120, 110)
End Sub

' Single source for generated captions, so the forms can switch language at run time.
Public Function EI_Caption(ByVal strControlName As String) As String
    Select Case strControlName
        Case strIMPORT_FORM_NAME: EI_Caption = EI_Text("Excel Data Import · Batch Append", "通用 Excel 数据导入 · 批量追加")
        Case "frmExcelImportErrors": EI_Caption = EI_Text("Batch error details (ExcelRow = 0 means a batch-level error)", "本批次错误明细（ExcelRow=0 表示整批错误）")
        Case "frmExcelImportHistory": EI_Caption = EI_Text("Excel import history", "Excel 导入历史")
        Case "lblTitle": EI_Caption = EI_Text("Excel Data Import", "通用 Excel 数据导入")
        Case "lblHelp": EI_Caption = EI_Text("Select source → confirm field mapping → validate and append in batches", "选择数据源 → 确认字段映射 → 校验并批量追加")
        Case "lblSourceTitle": EI_Caption = EI_Text("Data source", "数据源设置")
        Case "lblReadOnly": EI_Caption = EI_Text("The source file is opened read-only", "只读打开源文件，不修改原始数据")
        Case "lblTable": EI_Caption = EI_Text("Target table", "目标表")
        Case "lblSheet": EI_Caption = EI_Text("Worksheet", "工作表")
        Case "lblFile": EI_Caption = EI_Text("Excel file", "Excel 文件")
        Case "cmdBrowse": EI_Caption = EI_Text("Choose Excel file", "选择 Excel 文件")
        Case "lblHeader": EI_Caption = EI_Text("Header row", "标题行")
        Case "lblHeaderHelp": EI_Caption = EI_Text("Default is row 1; the data starts below the header row", "默认第 1 行；标题行下方为待导入数据")
        Case "lblMap": EI_Caption = EI_Text("Field mapping", "字段映射")
        Case "lblMapHelp": EI_Caption = EI_Text("Check ""Skip"" to exclude a column; empty cells are imported as Null", "勾选“跳过”则不导入该列；空值按 Null 处理")
        Case "lblSkip": EI_Caption = EI_Text("Skip rows that fail validation (unique index and other constraint errors still roll back the whole batch)", "跳过校验失败的行（唯一索引等约束错误仍整批回滚）")
        Case "lblStatus": EI_Caption = EI_Text("Ready", "准备就绪")
        Case "cmdErrors": EI_Caption = EI_Text("Error details", "错误明细")
        Case "cmdHistory": EI_Caption = EI_Text("Import history", "导入历史")
        Case "cmdLanguage": EI_Caption = EI_Text("中文", "English")
        Case "cmdCancel": EI_Caption = EI_Text("Close", "关闭")
        Case "cmdOK": EI_Caption = EI_Text("Start import", "开始导入")
        Case "lblExcelColumn": EI_Caption = EI_Text("Excel column", "Excel 列")
        Case "lblExcelField": EI_Caption = EI_Text("Excel header", "Excel 标题")
        Case "lblTargetField": EI_Caption = EI_Text("Target field", "目标字段")
        Case "lblSkipColumn": EI_Caption = EI_Text("Skip", "跳过")
    End Select
End Function

Public Sub EI_ApplyImportFormLanguage(ByVal frmImport As Access.Form)
    Dim ctlControl As Access.Control
    Dim strCaption As String
    frmImport.Caption = EI_Caption(strIMPORT_FORM_NAME)
    For Each ctlControl In frmImport.Controls
        If ctlControl.ControlType = acLabel Or ctlControl.ControlType = acCommandButton Then
            ' Keep the latest progress or result message when switching language.
            If ctlControl.Name <> "lblStatus" Then
                strCaption = EI_Caption(ctlControl.Name)
                If Len(strCaption) > 0 Then
                    ctlControl.Caption = strCaption
                End If
            End If
        End If
    Next
    If EI_IsBusy() Then
        frmImport!cmdCancel.Caption = EI_Text("Cancel import", "取消导入")
    End If
    If frmImport.CurrentView <> 0 Then
        EI_ApplyMapFormLanguage frmImport!sfrDetail.Form
    End If
End Sub

Public Sub EI_ApplyMapFormLanguage(ByVal frmMap As Access.Form)
    Dim ctlControl As Access.Control
    Dim strCaption As String
    For Each ctlControl In frmMap.Controls
        If ctlControl.ControlType = acLabel Then
            strCaption = EI_Caption(ctlControl.Name)
            If Len(strCaption) > 0 Then
                ctlControl.Caption = strCaption
            End If
        End If
    Next
End Sub

Public Sub EI_ApplyLogFormLanguage(ByVal frmLog As Access.Form)
    Dim ctlControl As Access.Control
    frmLog.Caption = EI_Caption(frmLog.Name)
    For Each ctlControl In frmLog.Controls
        If ctlControl.ControlType = acLabel And Left$(ctlControl.Name, 3) = "lbl" Then
            ctlControl.Caption = EI_LogCaption(Mid$(ctlControl.Name, 4))
        End If
    Next
End Sub

Public Sub EI_ApplyOpenFormsLanguage()
    If CurrentProject.AllForms(strIMPORT_FORM_NAME).IsLoaded Then
        EI_ApplyImportFormLanguage Forms(strIMPORT_FORM_NAME)
    End If
    If CurrentProject.AllForms("frmExcelImportErrors").IsLoaded Then
        EI_ApplyLogFormLanguage Forms("frmExcelImportErrors")
    End If
    If CurrentProject.AllForms("frmExcelImportHistory").IsLoaded Then
        EI_ApplyLogFormLanguage Forms("frmExcelImportHistory")
    End If
End Sub

Public Sub EI_Install()
    On Error GoTo Failed
    Dim objConnection As Object
    Dim lngErrorNumber As Long
    Dim strMessage As String
    Set objConnection = EI_LocalConnection()
    If Not EI_TableExists(objConnection, strMAP_TABLE_NAME) Then
        EI_Execute objConnection, "CREATE TABLE tblEI_Map (ColIndex LONG CONSTRAINT PK_EI_Map PRIMARY KEY, ColLetter TEXT(8), ExcelField TEXT(255), AccessField TEXT(64), SkipColumn YESNO)"
    End If
    If Not EI_TableExists(objConnection, strBATCH_TABLE_NAME) Then
        EI_Execute objConnection, "CREATE TABLE tblEI_Batches (BatchID COUNTER CONSTRAINT PK_EI_Batches PRIMARY KEY, StartedAt DATETIME, FinishedAt DATETIME, FileName LONGTEXT, SheetName TEXT(255), TargetTable TEXT(64), SourceRows LONG, ImportedRows LONG, RejectedRows LONG, Status TEXT(32), Details LONGTEXT)"
    End If
    If Not EI_TableExists(objConnection, strERROR_TABLE_NAME) Then
        EI_Execute objConnection, "CREATE TABLE tblEI_Errors (ErrorID COUNTER CONSTRAINT PK_EI_Errors PRIMARY KEY, BatchID LONG, ExcelRow LONG, FieldName TEXT(64), SourceValue LONGTEXT, Reason LONGTEXT)"
        EI_Execute objConnection, "CREATE INDEX IX_EI_ErrorRow ON tblEI_Errors (BatchID, ExcelRow)"
    End If
    EI_CloseConnection objConnection
    Exit Sub
Failed:
    lngErrorNumber = Err.Number
    strMessage = Err.Description
    On Error GoTo CleanupFailed
    EI_CloseConnection objConnection
    On Error GoTo 0
    Err.Raise lngErrorNumber, "ExcelImport.Install", strMessage
CleanupFailed:
    Err.Raise lngErrorNumber, "ExcelImport.Install", strMessage & EI_Text("; closing the setup connection failed: ", "；关闭建表连接失败：") & Err.Description
End Sub

Public Function EI_TableExists(ByVal objConnection As Object, ByVal strTableName As String) As Boolean
    Dim rstData As Object
    Set rstData = objConnection.OpenSchema(20, Array(Empty, Empty, strTableName, Empty))
    EI_TableExists = Not rstData.EOF
    rstData.Close
End Function

Public Sub EI_RequireImportTables(ByVal objConnection As Object)
    Dim varTableName As Variant
    For Each varTableName In Array(strMAP_TABLE_NAME, strBATCH_TABLE_NAME, strERROR_TABLE_NAME)
        If Not EI_TableExists(objConnection, CStr(varTableName)) Then
            Err.Raise vbObjectError + 2141, "ExcelImport.Setup", EI_Format(EI_Text("Import setup table {0} is missing. Run EI_Install or CreateExcelImportForm from the form generator module first.", "缺少导入配置表：{0}。请先运行窗体生成模块中的 EI_Install 或 CreateExcelImportForm。"), varTableName)
        End If
    Next
End Sub

Public Sub EI_DropStage(ByVal objConnection As Object)
    If EI_TableExists(objConnection, strTYPED_STAGE_TABLE_NAME) Then EI_Execute objConnection, "DROP TABLE tblEI_TypedStage"
    If EI_TableExists(objConnection, strRAW_STAGE_TABLE_NAME) Then EI_Execute objConnection, "DROP TABLE tblEI_Stage"
End Sub

Public Sub EI_CreateStageTable(ByVal objConnection As Object, ByRef udtColumns() As typImportColumn, ByVal blnTypedStage As Boolean)
    Dim strCreateTableSql As String
    Dim lngColumnIndex As Long
    If UBound(udtColumns) > 254 Then Err.Raise vbObjectError + 2106, , EI_Text("At most 254 columns can be mapped at once.", "一次最多映射 254 列。")
    If blnTypedStage Then
        strCreateTableSql = "CREATE TABLE tblEI_TypedStage (EI_Row LONG CONSTRAINT PK_EI_TypedStage PRIMARY KEY"
    Else
        strCreateTableSql = "CREATE TABLE tblEI_Stage (EI_Row LONG CONSTRAINT PK_EI_Stage PRIMARY KEY"
    End If
    For lngColumnIndex = 1 To UBound(udtColumns)
        If blnTypedStage Then
            strCreateTableSql = strCreateTableSql & ",T" & lngColumnIndex & " " & EI_AccessType(udtColumns(lngColumnIndex))
        Else
            strCreateTableSql = strCreateTableSql & ",F" & lngColumnIndex & " LONGTEXT"
        End If
    Next
    EI_Execute objConnection, strCreateTableSql & ")"
End Sub

Private Function EI_AccessType(ByRef udtColumn As typImportColumn) As String
    Select Case udtColumn.lngFieldType
        Case dbBoolean: EI_AccessType = "YESNO"
        Case dbByte: EI_AccessType = "BYTE"
        Case dbInteger: EI_AccessType = "SHORT"
        Case dbLong: EI_AccessType = "LONG"
        Case dbCurrency: EI_AccessType = "CURRENCY"
        Case dbSingle: EI_AccessType = "SINGLE"
        Case dbDouble: EI_AccessType = "DOUBLE"
        Case dbDate: EI_AccessType = "DATETIME"
        Case dbText: EI_AccessType = "TEXT(" & udtColumn.lngFieldSize & ")"
        Case dbMemo: EI_AccessType = "LONGTEXT"
        Case 16: EI_AccessType = "BIGINT"
        Case 20: EI_AccessType = "DECIMAL(" & udtColumn.bytPrecision & "," & udtColumn.bytScale & ")"
        Case Else: Err.Raise vbObjectError + 2111, , EI_Text("Unsupported staging field type.", "不支持的暂存字段类型。")
    End Select
End Function

Public Sub EI_CreateServerStage(ByVal objConnection As Object, ByRef udtColumns() As typImportColumn)
    Dim strCreateTableSql As String
    Dim lngColumnIndex As Long
    strCreateTableSql = "CREATE TABLE #EI_Stage (EI_Row int NOT NULL PRIMARY KEY"
    For lngColumnIndex = 1 To UBound(udtColumns)
        strCreateTableSql = strCreateTableSql & ",T" & lngColumnIndex & " " & udtColumns(lngColumnIndex).strSqlType & " NULL"
    Next
    EI_Execute objConnection, strCreateTableSql & ")"
End Sub

Public Sub CreateExcelImportForm(Optional ByVal blnSilent As Boolean = False)
    On Error GoTo Failed
    Dim frmImport As Access.Form
    Dim ctlControl As Access.Control
    Dim objAccessObject As Access.AccessObject
    Dim strFormName As String
    EI_InitializeColors
    EI_Install
    EI_RequireClosed strIMPORT_FORM_NAME
    EI_RequireClosed strIMPORT_MAP_FORM_NAME
    EI_RequireClosed "frmExcelImportErrors"
    EI_RequireClosed "frmExcelImportHistory"
    EI_CreateMapForm
    EI_CreateLogForm "frmExcelImportErrors", "tblEI_Errors", "ExcelRow,ErrorID"
    EI_CreateLogForm "frmExcelImportHistory", "tblEI_Batches", "BatchID DESC"
    Set frmImport = CreateForm
    strFormName = frmImport.Name
    frmImport.Caption = EI_Caption(strIMPORT_FORM_NAME)
    DoCmd.SelectObject acForm, strFormName, False
    DoCmd.RunCommand acCmdFormHdrFtr
    frmImport.Width = 13546
    frmImport.Section(lngFORM_SECTION_HEADER).Height = 3514
    frmImport.Section(acDetail).Height = 3565
    frmImport.Section(lngFORM_SECTION_FOOTER).Height = 1870
    EI_StyleForm frmImport
    frmImport.Section(acDetail).BackColor = vbWhite
    frmImport.Section(lngFORM_SECTION_HEADER).BackColor = vbWhite
    frmImport.Section(lngFORM_SECTION_FOOTER).BackColor = vbWhite
    frmImport.DefaultView = 0
    frmImport.ViewsAllowed = 1
    frmImport.OnLoad = "[Event Procedure]"
    frmImport.OnUnload = "[Event Procedure]"
    EI_Label strFormName, "lblTitle", 56, 56, 11000, 520, lngFORM_SECTION_HEADER
    frmImport!lblTitle.FontSize = 20
    frmImport!lblTitle.FontBold = True
    frmImport!lblTitle.ForeColor = mlngHeadingColor
    EI_Label strFormName, "lblHelp", 60, 615, 13400, 360, lngFORM_SECTION_HEADER
    frmImport!lblHelp.ForeColor = mlngMutedTextColor
    EI_Label strFormName, "lblSourceTitle", 60, 1050, 4800, 400, lngFORM_SECTION_HEADER
    frmImport!lblSourceTitle.FontBold = True
    frmImport!lblSourceTitle.FontSize = 11
    frmImport!lblSourceTitle.ForeColor = mlngHeadingColor
    EI_Label strFormName, "lblReadOnly", 8360, 1090, 5100, 330, lngFORM_SECTION_HEADER
    frmImport!lblReadOnly.ForeColor = mlngMutedTextColor
    frmImport!lblReadOnly.FontSize = 9
    EI_Label strFormName, "lblTable", 60, 1510, 1100, 386, lngFORM_SECTION_HEADER
    Set ctlControl = EI_Control(strFormName, acComboBox, "cboTableName", 1360, 1440, 5450, 386, lngFORM_SECTION_HEADER)
    ctlControl.RowSourceType = "Value List"
    ctlControl.LimitToList = True
    ctlControl.AfterUpdate = "=EI_Map()"
    EI_Label strFormName, "lblSheet", 7610, 1510, 1100, 386, lngFORM_SECTION_HEADER
    Set ctlControl = EI_Control(strFormName, acComboBox, "cboSheetName", 8810, 1440, 4650, 386, lngFORM_SECTION_HEADER)
    ctlControl.RowSourceType = "Value List"
    ctlControl.LimitToList = True
    ctlControl.AfterUpdate = "=EI_Map()"
    EI_Label strFormName, "lblFile", 60, 2035, 1100, 386, lngFORM_SECTION_HEADER
    Set ctlControl = EI_Control(strFormName, acTextBox, "txtFileName", 1360, 1965, 9100, 386, lngFORM_SECTION_HEADER)
    ctlControl.Locked = True
    ctlControl.BackColor = mlngFormBackColor
    EI_Button strFormName, "cmdBrowse", "=EI_Browse()", 10760, 1965, 2700, 386, False, lngFORM_SECTION_HEADER
    EI_Label strFormName, "lblHeader", 60, 2530, 1100, 386, lngFORM_SECTION_HEADER
    Set ctlControl = EI_Control(strFormName, acTextBox, "txtHeaderRow", 1360, 2460, 950, 386, lngFORM_SECTION_HEADER)
    ctlControl.DefaultValue = "1"
    ctlControl.TextAlign = 2
    ctlControl.AfterUpdate = "=EI_Map()"
    EI_Label strFormName, "lblHeaderHelp", 2660, 2530, 10800, 386, lngFORM_SECTION_HEADER
    frmImport!lblHeaderHelp.ForeColor = mlngMutedTextColor
    frmImport!lblHeaderHelp.FontSize = 9
    EI_Label strFormName, "lblMap", 60, 3000, 2200, 400, lngFORM_SECTION_HEADER
    frmImport!lblMap.FontSize = 11
    frmImport!lblMap.FontBold = True
    frmImport!lblMap.ForeColor = mlngHeadingColor
    EI_Label strFormName, "lblMapHelp", 2560, 3020, 10900, 330, lngFORM_SECTION_HEADER
    frmImport!lblMapHelp.ForeColor = mlngMutedTextColor
    frmImport!lblMapHelp.FontSize = 9
    Set ctlControl = EI_Control(strFormName, acSubform, "sfrDetail", 60, 0, 13430, 3565, acDetail)
    ctlControl.SourceObject = "Form." & strIMPORT_MAP_FORM_NAME
    ctlControl.HorizontalAnchor = 2
    ctlControl.VerticalAnchor = 2
    ctlControl.SpecialEffect = 0
    ctlControl.BorderStyle = 1
    ctlControl.BorderColor = mlngBorderColor
    Set ctlControl = EI_Control(strFormName, acCheckBox, "chkSkipInvalid", 60, 105, 300, 300, lngFORM_SECTION_FOOTER)
    ctlControl.DefaultValue = "False"
    EI_Label strFormName, "lblSkip", 480, 105, 12980, 360, lngFORM_SECTION_FOOTER
    frmImport!lblSkip.FontSize = 9
    EI_Label strFormName, "lblStatus", 60, 565, 13400, 450, lngFORM_SECTION_FOOTER
    frmImport!lblStatus.ForeColor = mlngBlueColor
    EI_Button strFormName, "cmdErrors", "=EI_Errors()", 60, 1225, 1700, 500, False, lngFORM_SECTION_FOOTER
    EI_Button strFormName, "cmdHistory", "=EI_History()", 2000, 1225, 1700, 500, False, lngFORM_SECTION_FOOTER
    EI_Button strFormName, "cmdLanguage", "=EI_ToggleLanguage()", 3940, 1225, 1500, 500, False, lngFORM_SECTION_FOOTER
    EI_Button strFormName, "cmdCancel", "=EI_Close()", 9340, 1225, 1700, 500, False, lngFORM_SECTION_FOOTER
    EI_Button strFormName, "cmdOK", "=EI_Run()", 11240, 1225, 2220, 500, True, lngFORM_SECTION_FOOTER
    Set ctlControl = EI_Control(strFormName, acTextBox, "txtBatchID", 5960, 1225, 800, 400, lngFORM_SECTION_FOOTER)
    ctlControl.Visible = False
    EI_ResetModule frmImport
    EI_AddLine frmImport, "Option Compare Database"
    EI_AddLine frmImport, "Option Explicit"
    EI_AddLine frmImport, "Private Sub Form_Load()"
    EI_AddLine frmImport, "    EI_Load"
    EI_AddLine frmImport, "End Sub"
    EI_AddLine frmImport, "Private Sub Form_Unload(intCancel As Integer)"
    EI_AddLine frmImport, "    On Error GoTo Failed"
    EI_AddLine frmImport, "    If EI_IsBusy() Then"
    EI_AddLine frmImport, "        intCancel = True"
    EI_AddLine frmImport, "        EI_Cancel"
    EI_AddLine frmImport, "    Else"
    EI_AddLine frmImport, "        EI_ReleaseExcel"
    EI_AddLine frmImport, "    End If"
    EI_AddLine frmImport, "    Exit Sub"
    EI_AddLine frmImport, "Failed:"
    EI_AddLine frmImport, "    intCancel = True"
    EI_AddLine frmImport, "    MsgBox Err.Description, vbExclamation, EI_Text(""Close import form"", ""关闭导入窗体"")"
    EI_AddLine frmImport, "End Sub"
    EI_SaveForm strFormName, strIMPORT_FORM_NAME
    If Not blnSilent Then
        DoCmd.OpenForm strIMPORT_FORM_NAME
        MsgBox EI_Text("The import forms were generated.", "导入窗体已生成。"), vbInformation, EI_Text("Excel import", "Excel 导入")
    End If
    Exit Sub
Failed:
    Dim strMessage As String
    strMessage = Err.Description
    On Error GoTo CleanupFailed
    If Len(strFormName) > 0 Then
        For Each objAccessObject In CurrentProject.AllForms
            If objAccessObject.Name = strFormName Then
                If objAccessObject.IsLoaded Then DoCmd.Close acForm, strFormName, acSaveNo
                Exit For
            End If
        Next
    End If
    If Not blnSilent Then MsgBox strMessage, vbCritical, EI_Text("Generate forms", "生成窗体失败")
    On Error GoTo 0
    Err.Raise vbObjectError + 2200, "CreateExcelImportForm", strMessage
CleanupFailed:
    Err.Raise vbObjectError + 2200, "CreateExcelImportForm", strMessage & EI_Text("; closing the design form failed: ", "；关闭设计窗体失败：") & Err.Description
End Sub

Private Sub EI_StyleForm(ByVal frmImport As Access.Form)
    frmImport.ScrollBars = 0
    frmImport.PopUp = True
    frmImport.Modal = True
    frmImport.BorderStyle = 3
    frmImport.MinMaxButtons = 0
    frmImport.AutoCenter = True
    frmImport.AutoResize = True
    frmImport.RecordSelectors = False
    frmImport.NavigationButtons = False
    frmImport.DividingLines = False
    frmImport.Section(acDetail).BackColor = mlngFormBackColor
    frmImport.DatasheetBackColor = vbWhite
    frmImport.DatasheetAlternateBackColor = mlngFormBackColor
    frmImport.DatasheetForeColor = mlngTextColor
    frmImport.DatasheetGridlinesColor = mlngBorderColor
    frmImport.DatasheetFontName = "微软雅黑"
    frmImport.DatasheetFontHeight = 10
End Sub

Private Sub EI_RequireClosed(ByVal strFormName As String)
    Dim objAccessObject As Access.AccessObject
    For Each objAccessObject In CurrentProject.AllForms
        If objAccessObject.Name = strFormName Then
            If objAccessObject.IsLoaded Then Err.Raise vbObjectError + 2201, , EI_Text("Close this form first: ", "请先关闭窗体：") & strFormName
        End If
    Next
End Sub

Private Function EI_Control(ByVal strFormName As String, ByVal lngDataType As AcControlType, ByVal strControlName As String, ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngWidth As Long, ByVal lngHeight As Long, Optional ByVal lngSection As Long = 0) As Access.Control
    Dim ctlControl As Access.Control
    Set ctlControl = CreateControl(strFormName, lngDataType, lngSection, , , lngLeft, lngTop, lngWidth, lngHeight)
    ctlControl.Name = strControlName
    Select Case lngDataType
        Case acTextBox, acComboBox, acCommandButton, acLabel
            ctlControl.FontName = "微软雅黑"
            ctlControl.FontSize = 10
            ctlControl.ForeColor = mlngTextColor
    End Select
    Select Case lngDataType
        Case acTextBox, acComboBox
            ctlControl.SpecialEffect = 0
            ctlControl.BorderStyle = 1
            ctlControl.BorderColor = mlngBorderColor
            ctlControl.BackColor = vbWhite
        Case acLabel
            ctlControl.BackStyle = 0
            ctlControl.BorderStyle = 0
    End Select
    If lngDataType = acComboBox Then
        ctlControl.AllowValueListEdits = False
    End If
    Set EI_Control = ctlControl
End Function

Private Sub EI_Label(ByVal strFormName As String, ByVal strControlName As String, ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngWidth As Long, ByVal lngHeight As Long, Optional ByVal lngSection As Long = 0)
    Dim ctlControl As Access.Control
    Set ctlControl = EI_Control(strFormName, acLabel, strControlName, lngLeft, lngTop, lngWidth, lngHeight, lngSection)
    ctlControl.Caption = EI_Caption(strControlName)
End Sub

Private Sub EI_Button(ByVal strFormName As String, ByVal strControlName As String, ByVal strClickAction As String, ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngWidth As Long, ByVal lngHeight As Long, Optional ByVal blnPrimaryButton As Boolean = False, Optional ByVal lngSection As Long = 0)
    Dim ctlControl As Access.Control
    Set ctlControl = EI_Control(strFormName, acCommandButton, strControlName, lngLeft, lngTop, lngWidth, lngHeight, lngSection)
    ctlControl.Caption = EI_Caption(strControlName)
    ctlControl.OnClick = strClickAction
    ctlControl.UseTheme = False
    ctlControl.Gradient = 0
    ctlControl.FontBold = True
    ctlControl.BorderStyle = 1
    ctlControl.BorderColor = mlngBlueColor
    ctlControl.ForeColor = mlngBlueColor
    ctlControl.BackColor = vbWhite
    ctlControl.HoverColor = mlngLightBlueColor
    ctlControl.HoverForeColor = mlngBlueColor
    ctlControl.PressedColor = RGB(207, 226, 255)
    ctlControl.PressedForeColor = mlngBlueColor
    If blnPrimaryButton Then
        ctlControl.BackThemeColorIndex = -1
        ctlControl.ForeThemeColorIndex = -1
        ctlControl.BorderThemeColorIndex = -1
        ctlControl.Gradient = 0
        ctlControl.BorderColor = mlngImportButtonBackColor
        ctlControl.BackColor = mlngImportButtonBackColor
        ctlControl.ForeColor = mlngBlueColor
        ctlControl.HoverColor = RGB(0, 105, 96)
        ctlControl.HoverForeColor = mlngBlueColor
        ctlControl.PressedColor = RGB(0, 90, 82)
        ctlControl.PressedForeColor = mlngBlueColor
    End If
End Sub

Private Sub EI_ResetModule(ByVal frmTarget As Access.Form)
    frmTarget.HasModule = True
    If frmTarget.Module.CountOfLines > 0 Then
        frmTarget.Module.DeleteLines 1, frmTarget.Module.CountOfLines
    End If
End Sub

Private Sub EI_AddLine(ByVal frmImport As Access.Form, ByVal strCodeLine As String)
    frmImport.HasModule = True
    frmImport.Module.InsertLines frmImport.Module.CountOfLines + 1, strCodeLine
End Sub

Private Sub EI_SaveForm(ByVal strTemporaryFormName As String, ByVal strFinalFormName As String)
    Dim objAccessObject As Access.AccessObject
    DoCmd.Save acForm, strTemporaryFormName
    DoCmd.Close acForm, strTemporaryFormName, acSaveYes
    For Each objAccessObject In CurrentProject.AllForms
        If objAccessObject.Name = strFinalFormName Then
            DoCmd.DeleteObject acForm, strFinalFormName
            Exit For
        End If
    Next
    DoCmd.Rename strFinalFormName, acForm, strTemporaryFormName
End Sub

Private Sub EI_CreateMapForm()
    Dim frmImport As Access.Form
    Dim ctlControl As Access.Control
    Dim strFormName As String
    Set frmImport = CreateForm
    strFormName = frmImport.Name
    EI_StyleForm frmImport
    frmImport.Width = 11100
    frmImport.Section(acDetail).Height = 5952
    frmImport.RecordSource = "SELECT * FROM tblEI_Map ORDER BY ColIndex"
    frmImport.DefaultView = 2
    frmImport.ViewsAllowed = 2
    frmImport.AllowAdditions = False
    frmImport.AllowDeletions = False
    frmImport.NavigationButtons = False
    frmImport.OnLoad = "[Event Procedure]"
    Set ctlControl = EI_Control(strFormName, acTextBox, "txtIndex", 100, 100, 900, 300)
    ctlControl.ControlSource = "ColIndex"
    ctlControl.ColumnHidden = True
    ctlControl.Locked = True
    Set ctlControl = EI_Control(strFormName, acTextBox, "txtLetter", 1100, 100, 900, 300)
    ctlControl.ControlSource = "ColLetter"
    ctlControl.ColumnWidth = 1100
    ctlControl.Locked = True
    EI_FieldLabel strFormName, ctlControl.Name, "lblExcelColumn", EI_Caption("lblExcelColumn")
    Set ctlControl = EI_Control(strFormName, acTextBox, "txtExcelField", 2100, 100, 3500, 300)
    ctlControl.ControlSource = "ExcelField"
    ctlControl.ColumnWidth = 3800
    ctlControl.Locked = True
    EI_FieldLabel strFormName, ctlControl.Name, "lblExcelField", EI_Caption("lblExcelField")
    Set ctlControl = EI_Control(strFormName, acComboBox, "cboAccFld", 5700, 100, 3500, 300)
    ctlControl.ControlSource = "AccessField"
    ctlControl.RowSourceType = "Value List"
    ctlControl.LimitToList = True
    ' Field name is stored; the second list column shows its caption.
    ctlControl.ColumnCount = 2
    ctlControl.BoundColumn = 1
    ctlControl.ColumnWidths = "2268;2268"
    ctlControl.ListWidth = 4536
    ctlControl.BeforeUpdate = "[Event Procedure]"
    ctlControl.AfterUpdate = "[Event Procedure]"
    ctlControl.ColumnWidth = 3800
    EI_FieldLabel strFormName, ctlControl.Name, "lblTargetField", EI_Caption("lblTargetField")
    Set ctlControl = EI_Control(strFormName, acCheckBox, "chkSkip", 9300, 100, 400, 300)
    ctlControl.ControlSource = "SkipColumn"
    ctlControl.ColumnWidth = 1000
    EI_FieldLabel strFormName, ctlControl.Name, "lblSkipColumn", EI_Caption("lblSkipColumn")
    EI_ResetModule frmImport
    EI_AddLine frmImport, "Option Compare Database"
    EI_AddLine frmImport, "Option Explicit"
    EI_AddLine frmImport, "Private Sub Form_Load()"
    EI_AddLine frmImport, "    EI_ApplyMapFormLanguage Me"
    EI_AddLine frmImport, "End Sub"
    EI_AddLine frmImport, "Private Sub cboAccFld_BeforeUpdate(intCancel As Integer)"
    EI_AddLine frmImport, "    If EI_MapFieldUsed(Me) Then"
    EI_AddLine frmImport, "        intCancel = True"
    EI_AddLine frmImport, "        Me!cboAccFld.Undo"
    EI_AddLine frmImport, "    End If"
    EI_AddLine frmImport, "End Sub"
    EI_AddLine frmImport, "Private Sub cboAccFld_AfterUpdate()"
    EI_AddLine frmImport, "    Me!chkSkip = IsNull(Me!cboAccFld)"
    EI_AddLine frmImport, "End Sub"
    EI_SaveForm strFormName, strIMPORT_MAP_FORM_NAME
End Sub

Private Sub EI_CreateLogForm(ByVal strFinalFormName As String, ByVal strTableName As String, ByVal strOrderBy As String)
    Dim frmImport As Access.Form
    Dim ctlControl As Access.Control
    Dim dbsDatabase As DAO.Database
    Dim tdfTargetTable As DAO.TableDef
    Dim fldField As DAO.Field
    Dim strFormName As String
    Dim lngLeft As Long
    Dim lngTop As Long
    Set dbsDatabase = CurrentDb
    Set tdfTargetTable = dbsDatabase.TableDefs(strTableName)
    Set frmImport = CreateForm
    strFormName = frmImport.Name
    EI_StyleForm frmImport
    frmImport.Width = 13200
    frmImport.Section(acDetail).Height = 5952
    frmImport.Caption = EI_Caption(strFinalFormName)
    frmImport.OnLoad = "[Event Procedure]"
    frmImport.RecordSource = "SELECT * FROM " & EI_Q(strTableName) & " ORDER BY " & strOrderBy
    frmImport.DefaultView = 2
    frmImport.ViewsAllowed = 2
    frmImport.AllowAdditions = False
    frmImport.AllowDeletions = False
    frmImport.AllowEdits = False
    lngLeft = 100
    lngTop = 100
    For Each fldField In tdfTargetTable.Fields
        If lngLeft + 1800 > 12800 Then
            lngLeft = 100
            lngTop = lngTop + 850
        End If
        Set ctlControl = EI_Control(strFormName, acTextBox, "txt" & fldField.Name, lngLeft, lngTop, 1800, 300)
        ctlControl.ControlSource = fldField.Name
        ctlControl.ColumnWidth = 2200
        If fldField.Type = dbDate Then ctlControl.Format = "yyyy-mm-dd hh:nn:ss"
        EI_FieldLabel strFormName, ctlControl.Name, "lbl" & fldField.Name, EI_LogCaption(fldField.Name)
        lngLeft = lngLeft + 1900
    Next
    EI_ResetModule frmImport
    EI_AddLine frmImport, "Option Compare Database"
    EI_AddLine frmImport, "Option Explicit"
    EI_AddLine frmImport, "Private Sub Form_Load()"
    EI_AddLine frmImport, "    EI_ApplyLogFormLanguage Me"
    EI_AddLine frmImport, "End Sub"
    EI_SaveForm strFormName, strFinalFormName
End Sub

Private Sub EI_FieldLabel(ByVal strFormName As String, ByVal strParentControlName As String, ByVal strLabelName As String, ByVal strCaption As String)
    Dim ctlControl As Access.Control
    Dim ctlFieldControl As Access.Control
    Set ctlFieldControl = Forms(strFormName).Controls(strParentControlName)
    Set ctlControl = CreateControl(strFormName, acLabel, acDetail, strParentControlName, , ctlFieldControl.Left, ctlFieldControl.Top + 400, 1800, 300)
    ctlControl.Name = strLabelName
    ctlControl.Caption = strCaption
    ctlControl.FontName = "微软雅黑"
    ctlControl.FontSize = 10
    ctlControl.ForeColor = mlngTextColor
    ctlControl.BackStyle = 0
End Sub

Public Function EI_LogCaption(ByVal strFieldName As String) As String
    Select Case strFieldName
        Case "ErrorID": EI_LogCaption = EI_Text("Error ID", "错误编号")
        Case "BatchID": EI_LogCaption = EI_Text("Batch ID", "批次编号")
        Case "ExcelRow": EI_LogCaption = EI_Text("Excel row", "Excel 行号")
        Case "FieldName": EI_LogCaption = EI_Text("Target field", "目标字段")
        Case "SourceValue": EI_LogCaption = EI_Text("Source value", "原始值")
        Case "Reason": EI_LogCaption = EI_Text("Reason", "错误原因")
        Case "StartedAt": EI_LogCaption = EI_Text("Started", "开始时间")
        Case "FinishedAt": EI_LogCaption = EI_Text("Finished", "结束时间")
        Case "FileName": EI_LogCaption = EI_Text("Source file", "源文件")
        Case "SheetName": EI_LogCaption = EI_Text("Worksheet", "工作表")
        Case "TargetTable": EI_LogCaption = EI_Text("Target table", "目标表")
        Case "SourceRows": EI_LogCaption = EI_Text("Non-empty rows", "非空数据行")
        Case "ImportedRows": EI_LogCaption = EI_Text("Appended rows", "已追加行")
        Case "RejectedRows": EI_LogCaption = EI_Text("Rejected rows", "校验失败行")
        Case "Status": EI_LogCaption = EI_Text("Status", "状态")
        Case "Details": EI_LogCaption = EI_Text("Details", "说明")
        Case Else: EI_LogCaption = strFieldName
    End Select
End Function
