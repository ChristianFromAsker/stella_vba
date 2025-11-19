Attribute VB_Name = "utilities"
Option Compare Database
Option Explicit

'used for get_event_id
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef GUID As GUID) As LongPtr
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Type typ_log_object
    change_source As String
    changer_id As String
    comment As String
    data_set As String
    deal_id As Long
    event_id As String
    executed_sql As String
    field_name As String
    new_value As Variant
    operation_type As String
    policy_id As Long
    record_id As Long
    security_id As Long
End Type

Public Function log_change__create_command() As ADODB.Command
    Const proc_name As String = "utilities.log_change__create_command"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    cmd.CommandText = _
    "INSERT INTO log_data_t (" _
        & "app_continent" _
        & ", app_name" _
        & ", change_source" _
        & ", changer_id" _
        & ", comment" _
        & ", data_set_id" _
        & ", deal_id" _
        & ", event_id" _
        & ", executed_sql" _
        & ", field_name" _
        & ", new_value" _
        & ", operation_type" _
        & ", policy_id" _
        & ", record_id" _
    & ") VALUES (" _
        & " ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?" _
    & ")"
    
    'Parameters must be added in the same order as the ? placeholders
    With cmd.Parameters
        .Append cmd.CreateParameter("p_app_continent", adVarChar, adParamInput, 50)
        .Append cmd.CreateParameter("p_app_name", adVarChar, adParamInput, 50)
        .Append cmd.CreateParameter("p_change_source", adVarChar, adParamInput, 255)
        .Append cmd.CreateParameter("p_changer_id", adVarChar, adParamInput, 100)
        .Append cmd.CreateParameter("p_comment", adLongVarChar, adParamInput, -1)
        .Append cmd.CreateParameter("p_data_set_id", adVarChar, adParamInput, 100)
        .Append cmd.CreateParameter("p_deal_id", adInteger, adParamInput)
        .Append cmd.CreateParameter("p_event_id", adVarChar, adParamInput, 36)
        .Append cmd.CreateParameter("p_executed_sql", adLongVarChar, adParamInput, -1)
        .Append cmd.CreateParameter("p_field_name", adVarChar, adParamInput, 100)
        .Append cmd.CreateParameter("p_new_value", adVarChar, adParamInput, 255)
        .Append cmd.CreateParameter("p_operation_type", adVarChar, adParamInput, 20)
        .Append cmd.CreateParameter("p_policy_id", adInteger, adParamInput)
        .Append cmd.CreateParameter("p_record_id", adInteger, adParamInput)
    End With
    
    Set log_change__create_command = cmd

outro:
    utilities.call_stack_remove_last_item
    Exit Function
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "", "global_vars.deal.deal_id = " & global_vars.deal.deal_id, "", True
    Resume outro
End Function
Public Sub log_change__field_change( _
ByRef log_obj_base As utilities.typ_log_object _
, ByVal field_name As String _
, ByVal new_val As Variant _
, ByVal data_set As String)

    Const proc_name As String = "utilities.log_change__field_change"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0

    Dim log_obj As utilities.typ_log_object

    log_obj = log_obj_base   ' copy base metadata (event_id, deal_id, etc.)

    log_obj.field_name = field_name
    log_obj.new_value = new_val
    log_obj.data_set = data_set

    log_change log_obj
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
        
Public Sub log_change(ByRef log_object As utilities.typ_log_object)

    Const proc_name As String = "full address of routine"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim log_cmd As ADODB.Command
    
    Set log_cmd = utilities.log_change__create_command
    log_cmd.ActiveConnection = load.conn
    
    With log_cmd
        .Parameters("p_app_continent").value = load.system_info.app_continent
        .Parameters("p_app_name").value = load.system_info.app_name
        .Parameters("p_change_source").value = log_object.change_source
        
        'changer_id
        If log_object.changer_id = "" Then
            .Parameters("p_changer_id").value = Environ("username")
        Else
            .Parameters("p_changer_id").value = log_object.changer_id
        End If
        
        .Parameters("p_comment").value = log_object.comment
        .Parameters("p_data_set_id").value = log_object.data_set
        
        'deal_id
        .Parameters("p_deal_id").value = Null
        If log_object.deal_id <> 0 Then .Parameters("p_deal_id").value = log_object.deal_id
        
        .Parameters("p_event_id").value = log_object.event_id
        .Parameters("p_executed_sql").value = log_object.executed_sql
        .Parameters("p_field_name").value = log_object.field_name
        
        'policy_id
        .Parameters("p_policy_id").value = Null
        If log_object.policy_id <> 0 Then .Parameters("p_policy_id").value = log_object.policy_id
        
        .Parameters("p_record_id").value = log_object.record_id
        .Parameters("p_new_value").value = CStr(log_object.new_value)
        .Parameters("p_operation_type").value = log_object.operation_type
        
        .Execute , , adExecuteNoRecords
        
    End With
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "", "deal_id = " & Nz(log_object.deal_id, 0) & ", event_id = " & log_object.event_id, "", True
    Resume outro
End Sub
Function get_event_id() As String
    Dim g As GUID
    Dim ret As LongPtr
    
    ret = CoCreateGuid(g)
    
    If ret = 0 Then
        get_event_id = _
            LCase$( _
            Right$("00000000" & Hex$(g.Data1), 8) & "-" & _
            Right$("0000" & Hex$(g.Data2), 4) & "-" & _
            Right$("0000" & Hex$(g.Data3), 4) & "-" & _
            Right$("00" & Hex$(g.Data4(0)), 2) & _
            Right$("00" & Hex$(g.Data4(1)), 2) & "-" & _
            Right$("00" & Hex$(g.Data4(2)), 2) & _
            Right$("00" & Hex$(g.Data4(3)), 2) & _
            Right$("00" & Hex$(g.Data4(4)), 2) & _
            Right$("00" & Hex$(g.Data4(5)), 2) & _
            Right$("00" & Hex$(g.Data4(6)), 2) & _
            Right$("00" & Hex$(g.Data4(7)), 2))
    Else
        get_event_id = ""
    End If
End Function

Public Sub call_stack_add_item(ByVal input_proc_name As String)
    load.call_stack = load.call_stack & vbNewLine & Time & " " & input_proc_name
End Sub
Public Sub call_stack_remove_last_item()
    Dim pos As Long
    pos = InStrRev(load.call_stack, vbNewLine)
    If pos > 0 Then load.call_stack = Left(load.call_stack, pos - 1)
End Sub
Public Sub col_policies_print(ByVal print_to_file As Boolean)
    Dim proc_name As String
    proc_name = "global_vars.col_policies_print"
    load.call_stack = load.call_stack & vbNewLine & proc_name
        
    Dim binder As Scripting.Dictionary
    Dim policy As Scripting.Dictionary
    
    For Each policy In global_vars.col_policies
        Debug.Print vbNewLine & "policy id: " & vbTab & vbTab & policy(policy_data.policy_id)
        Debug.Print "layer: " & vbTab & vbTab & vbTab & policy(policy_data.layer_no_text)
        Debug.Print "policy limit: " & vbTab & policy(policy_data.policy_limit)
        Debug.Print "policy no: " & vbTab & vbTab & policy(policy_data.policy_no)
        Debug.Print "binders: "
        If policy(policy_data.binders).Count = 0 Then
            Debug.Print vbTab & "No binders"
        Else
            For Each binder In policy(policy_data.binders)
                Debug.Print vbTab & "binder id: " & vbTab & vbTab & binder(binder_data.binder_id)
                Debug.Print vbTab & "binder name: " & vbTab & binder(binder_data.binder_name)
                Debug.Print vbTab & "binder limit: " & vbTab & binder(binder_data.binder_limit)
                Debug.Print vbTab & "max_limit_deal_currency: " & binder(binder_data.max_limit_deal_currency)
                Debug.Print vbTab & "default binder currency: " & binder(binder_data.default_binder_currency)
            Next binder
        End If
    Next policy
End Sub

Public Function get_policy_object(ByVal policy_id As Long, ByVal policy_list As Collection) As Scripting.Dictionary
    Dim policy As Scripting.Dictionary
    For Each policy In policy_list
        If policy(policy_data.policy_id) = policy_id Then
            Set get_policy_object = policy
            Exit Function
        End If
    Next policy
End Function



Public Function get_extra_limit_from_binders_on_deal(ByVal rs As ADODB.Recordset) As Collection
    Dim proc_name As String
    proc_name = "utilities.get_extra_limit_from_binders_on_deal"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim output As New Collection
    Set output = Nothing
    With output
        If rs.RecordCount = 0 Then
            GoTo outro
        End If
        rs.MoveFirst
        Do Until rs.EOF
            
            On Error Resume Next
            .Add rs!extra_limit.value, CStr(rs!binder_id.value)
            On Error GoTo err_handler
            If load.is_debugging = True Then On Error GoTo 0
            
            rs.MoveNext
        Loop
    End With
    
    Set get_extra_limit_from_binders_on_deal = output
outro:
    Exit Function

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "rs!deal_id = " & rs!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Function

Public Sub paint_control(ByVal form_name As String, ByVal col_input_controls As Collection)
    Dim proc_name As String
    proc_name = "utilities.paint_control_name"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0

    If CurrentProject.AllForms(form_name).IsLoaded = False Then
        GoTo outro
    End If
    
    Dim input_control As cls_field
    For Each input_control In col_input_controls
        If input_control.field_name = "-1" Then
            GoTo next_iteration
        End If
        With Forms(form_name).Controls(input_control.field_name)
            
            'bg_color
            If input_control.field_bg_color <> -1 Then
                .BackColor = input_control.field_bg_color
            End If
            
            'caption
            If input_control.field_caption <> "-1" Then
                .Caption = input_control.field_caption
            End If
            
            'font color
            If input_control.font_color <> -1 Then
                .ForeColor = input_control.font_color
            End If
            
            'height
            If input_control.field_height <> -1 Then
                .Height = input_control.field_height
            End If
            
            'left
            If input_control.field_left <> -1 Then
                .Left = input_control.field_left
            End If
            
            'top
            If input_control.field_top <> -1 Then
                .Top = input_control.field_top
            End If
            
            'value
            If input_control.field_value <> "-1" And input_control.field_type = control_types.txt_box Then
                .value = input_control.field_value
            End If
            
            'visible
            .Visible = input_control.field_visible
            
            'width
            If input_control.field_width <> -1 Then
                .Width = input_control.field_width
            End If
        End With
next_iteration:
    Next input_control
    
outro:
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "form_name = " & form_name _
        & vbNewLine & "input_control.field_name = " & input_control.field_name
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Public Function create_conn(ByVal conn_info As cls_conn_info) As ADODB.Connection
    Dim proc_name As String
    proc_name = "utilities.create_conn"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim str_conn As String
    str_conn = "Driver={MySQL ODBC 8.0 Unicode Driver};" _
        & "Server=" & conn_info.ip_address & ";" _
        & "DATABASE=" & conn_info.database_name & ";" _
        & "UID=" & conn_info.user_name & ";" _
        & "PWD=" & CStr(conn_info.get_password)
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = str_conn
    Set create_conn = conn
End Function

Public Function decrypt_string(ByVal input_string As String, ByVal salt_factor As Integer, ByVal offset_factor As Integer) As String
    Dim proc_name As String
    proc_name = "utilities.decrypt_string"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim pw_int As String, i As Integer, pw As String
    pw_int = pw_int & Mid(input_string, 1, 1)
    For i = 2 + salt_factor To Len(input_string) Step salt_factor + 1
        pw_int = pw_int & Mid(input_string, i, 1)
    Next i
    Dim int_letter As Integer
    pw = ""
    For i = 1 To Len(pw_int)
        int_letter = Asc(Mid(pw_int, i, 1))
        pw = pw & ChrW(int_letter - offset_factor)
    Next i
    decrypt_string = pw
outro:
End Function

Public Function convert_layer_to_words(ByVal int_layer As Integer) As String
    Dim proc_name As String
    proc_name = "utilities.convert_layer_to_words"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim str_return As String
    If int_layer = 0 Then
        str_return = "Primary"
    ElseIf int_layer = 1 Then
        str_return = "1st xs"
    ElseIf int_layer = 2 Then
        str_return = "2nd xs"
    ElseIf int_layer = 3 Then
        str_return = "3rd xs"
    ElseIf int_layer = 4 Then
        str_return = "4th xs"
    ElseIf int_layer = 5 Then
        str_return = "5th xs"
    ElseIf int_layer = 6 Then
        str_return = "6th xs"
    ElseIf int_layer = 7 Then
        str_return = "7th xs"
    ElseIf int_layer = 8 Then
        str_return = "8th xs"
    ElseIf int_layer = 9 Then
        str_return = "9th xs"
    ElseIf int_layer = 10 Then
        str_return = "10th xs"
    ElseIf int_layer = 11 Then
        str_return = "11th xs"
    ElseIf int_layer = 12 Then
        str_return = "12th xs"
    End If
    convert_layer_to_words = str_return
End Function
Public Function create_adodb_rs(ByVal conn As ADODB.Connection, ByVal str_sql As String) As ADODB.Recordset
    Const proc_name As String = "utilities.create_adodb_rs"
    utilities.call_stack_add_item proc_name
    
    Dim rs As New ADODB.Recordset
        With rs
            Set .ActiveConnection = conn
            .Source = str_sql
            .LockType = adOpenStatic
            .CursorType = adLockReadOnly
            .CursorLocation = adUseClient
        End With
    Set create_adodb_rs = rs
    Set rs = Nothing
    load.rs_counter = load.rs_counter + 1
    
outro:
    utilities.call_stack_remove_last_item
    Exit Function
End Function
Public Function get_binder_list_for_question(ByVal template_question_id) As Variant()
    Dim proc_name As String
    proc_name = "utilities.get_binder_list_for_question"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim binder_list()
    Dim str_sql As String, rs As ADODB.Recordset
    
    str_sql = "SELECT id, binder_id FROM " & load.sources.binder_questions_view _
    & " WHERE template_question_id = " & template_question_id
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        If rs.EOF And rs.BOF Then
            ReDim binder_list(0)
            binder_list(0) = -1
            get_binder_list_for_question = binder_list
            GoTo outro
        End If
        Dim i As Integer
        i = 1
        ReDim binder_list(0 To CLng(rs.RecordCount))
        Do Until rs.EOF
            binder_list(i) = rs!binder_id.value
            i = i + 1
            rs.MoveNext
        Loop
        binder_list(0) = rs.RecordCount
    rs.Close
    get_binder_list_for_question = binder_list
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Function
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: utilities.get_actual_binders_for_policy" & vbNewLine _
        & "Parameters: template_question_id = " & template_question_id & vbNewLine _
        & "App: CM UW", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Function
Public Function get_actual_binders_for_policy(ByVal policy_id As Long, ByVal col_input_policies As Collection) As Collection
    Dim proc_name As String
    proc_name = "utilities.get_actual_binders"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim binder As Scripting.Dictionary
    Dim policy As Scripting.Dictionary
    Dim col_output As Collection
    
    Set col_output = Nothing
    Set col_output = New Collection
    For Each policy In col_input_policies
        If policy(policy_data.policy_id) = policy_id Then
            For Each binder In policy(policy_data.binders)
                col_output.Add binder
            Next binder
        End If
    Next policy
    
    Set get_actual_binders_for_policy = col_output
    
outro:
    Exit Function

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "policy_id = " & policy_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Function
Public Function get_actual_questions_for_deal(ByVal deal_id) As Variant()
    Dim proc_name As String
    proc_name = "utilities.get_actual_questions_for_deal"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    Dim str_sql As String, rs As ADODB.Recordset, i As Integer

    Dim question_list() As Variant
    
    str_sql = "SELECT * FROM " & sources.deal_questions_view _
    & " WHERE deal_id = " & deal_id _
    & " AND (" _
        & " question_type_id = " & question_types.binder _
        & " OR question_type_id = " & question_types.auto_referral _
    & ") ORDER BY template_question_id ASC"
    
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        'exit sub if no binders on policy
        If rs.EOF And rs.BOF Then
            ReDim question_list(0)
            question_list(0) = -1
            get_actual_questions_for_deal = question_list
            GoTo outro
        End If
        ReDim question_list(0 To CLng(rs.RecordCount))
        rs.MoveFirst
        For i = 1 To CLng(rs.RecordCount)
            question_list(i) = rs!template_question_id.value
            rs.MoveNext
        Next i
        question_list(0) = rs.RecordCount
    rs.Close
    get_actual_questions_for_deal = question_list
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Function
err_handler:
    MsgBox load.system_info.error_instruction & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: utilities.get_actual_questions_for_deal" & vbNewLine _
        & "Parameters: deal_id = " & deal_id & vbNewLine _
        & "App: " & load.system_info.app_name, , load.system_info.error_msg_heading
    GoTo outro
End Function
Public Function get_binder_list_for_deal(ByVal deal_id) As Variant()
    'Purpose: get a list of binders on a deal
    Dim str_sql As String
    Dim binder_list() As Variant
    
    str_sql = "SELECT DISTINCT(binder_id) FROM " & load.sources.policy_binders_view & " WHERE deal_id = " & deal_id
    Dim rs As ADODB.Recordset
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        'exit sub if no binders on policy
        If rs.EOF And rs.BOF Then
            ReDim binder_list(0)
            binder_list(0) = -1
            get_binder_list_for_deal = binder_list
            GoTo outro
        End If
        ReDim binder_list(0 To CLng(rs.RecordCount))
        rs.MoveFirst
        Dim i As Integer
        For i = 1 To CLng(rs.RecordCount)
            binder_list(i) = rs!binder_id.value
            rs.MoveNext
        Next i
        binder_list(0) = rs.RecordCount
    rs.Close
    
    get_binder_list_for_deal = binder_list
outro:
    If Not rs Is Nothing Then Set rs = Nothing
End Function
Public Function get_policy_list_for_deal(ByVal deal_id) As Variant()
    Dim str_sql As String, output_list() As Variant
    str_sql = "SELECT DISTINCT(policy_id) FROM " & load.sources.policy_binders_view & " WHERE deal_id = " & deal_id
    Dim rs As ADODB.Recordset
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        'exit sub if no binders on policy
        If rs.EOF And rs.BOF Then
            output_list(0) = 666
            get_policy_list_for_deal = output_list
            GoTo outro
        End If
        ReDim output_list(1 To CLng(rs.RecordCount))
        rs.MoveFirst
        Dim i As Integer
        i = 0
        For i = 1 To CLng(rs.RecordCount)
            output_list(i) = rs!policy_id.value
            rs.MoveNext
        Next i
    rs.Close
    get_policy_list_for_deal = output_list
outro:
    Set rs = Nothing
End Function
Public Function twips_converter(ByVal input_number, ByVal inch_or_cm) As Long
    If inch_or_cm = "inch" Then
        twips_converter = input_number * 1440
    ElseIf inch_or_cm = "cm" Then
        twips_converter = input_number * 1440 / 2.54
    Else
        twips_converter = 666
    End If
End Function
Public Function generate_sql_date(ByVal input_date As Date) As String
    Dim str_day As String, str_month As String
    If Day(input_date) < 10 Then
        str_day = "0" & CStr(Day(input_date))
    Else
        str_day = CStr(Day(input_date))
    End If
    If Month(input_date) < 10 Then
        str_month = "0" & CStr(Month(input_date))
    Else
        str_month = CStr(Month(input_date))
    End If
    generate_sql_date = CStr(Year(input_date)) & "-" & str_month & "-" & str_day
End Function
Function convert_long_color_to_rgb(color_value As Long) As String
    Dim Red As Long, Green As Long, Blue As Long
    Red = color_value Mod 256
    Green = ((color_value - Red) / 256) Mod 256
    Blue = ((color_value - Red - (Green * 256)) / 256 / 256) Mod 256
    
    convert_long_color_to_rgb = "RGB(" & _
                    Red & ", " & _
                    Green & ", " & _
                    Blue & ")"
End Function

Function is_in_array(item_id As Long, item_list) As Boolean
    '6 January 2025, CK: Can only do one dimmensional arrays
    If Not (IsArray(item_list)) Then Exit Function
    If InStr(1, "'" & Join(item_list, "'") & "'", "'" & item_id & "'") > 0 Then
        is_in_array = True
    End If
End Function

Public Function get_actual_inar_referrals_for_deal(ByVal deal_id) As Variant()
    'latest review or change: 1 Apr 2023 by CK
    'intro
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    Dim str_sql As String, rs As ADODB.Recordset, i As Integer
    
    Dim inar_referrals() As Variant
    str_sql = "SELECT template_question_id FROM " & sources.referrals_view _
    & " WHERE question_type = " & question_types.inar & " AND deal_id = " & deal_id
    
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        If rs.EOF And rs.BOF Then
            ReDim inar_referrals(0)
            inar_referrals(0) = -1
            get_actual_inar_referrals_for_deal = inar_referrals
            GoTo outro
        End If
        ReDim inar_referrals(0 To CLng(rs.RecordCount))
        rs.MoveFirst
        For i = 1 To CLng(rs.RecordCount)
            inar_referrals(i) = rs!template_question_id.value
            rs.MoveNext
        Next i
        inar_referrals(0) = rs.RecordCount
    rs.Close
    get_actual_inar_referrals_for_deal = inar_referrals
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Function
err_handler:
    MsgBox load.system_info.error_instruction & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: utilities.get_actual_inar_referrals_for_deal" & vbNewLine _
        & "Parameters: deal_id = " & deal_id & vbNewLine _
        & "App: " & load.system_info.app_name, , load.system_info.error_msg_heading
    GoTo outro
End Function
Public Function get_actual_binder_referrals_for_deal(ByVal deal_id) As Variant()
    Dim proc_name As String
    proc_name = "utilities.get_actual_binder_referrals_for_deal"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String, rs As ADODB.Recordset, i As Integer
    
    Dim binder_referrals() As Variant
    str_sql = "SELECT id, deal_question_id FROM " & sources.referrals_view _
    & " WHERE question_type_id = " & question_types.binder & " AND deal_id = " & deal_id
    
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        If rs.EOF And rs.BOF Then
            ReDim binder_referrals(0)
            binder_referrals(0) = -1
            get_actual_binder_referrals_for_deal = binder_referrals
            GoTo outro
        End If
        ReDim binder_referrals(1 To CLng(rs.RecordCount))
        rs.MoveFirst
        For i = 1 To CLng(rs.RecordCount)
            binder_referrals(i) = rs!id.value
            rs.MoveNext
        Next i
    rs.Close
    get_actual_binder_referrals_for_deal = binder_referrals
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Function
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Function
Public Function get_default_answer_for_template_question(ByVal template_question_id) As Long
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim counter As Integer
    With load.col_template_questions
        For counter = 1 To .Count
            If load.col_template_questions(counter)("id") = template_question_id Then
                get_default_answer_for_template_question = load.col_template_questions(counter)("default_answer_id")
                GoTo outro
            End If
        Next counter
    End With
    
outro:
    Exit Function
    
err_handler:
    get_default_answer_for_template_question = -1
    
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "utilities.get_default_answer_for_template_question"
        .milestone = ""
        .params = "template_question_id = " & template_question_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
End Function

Public Function get_public_ip()
    Dim url As String, ip_adr As String
    With CreateObject("MSXML2.XMLHTTP.6.0")
        url = "https://checkip.amazonaws.com/"
        .Open "GET", url, False
        .Send
        ip_adr = .responseText
    Dim reg_exp As Object

    Set reg_exp = CreateObject("vbscript.regexp")
        If .status = 200 Then
            With reg_exp
                .Pattern = "\s"
                .MultiLine = True
                .Global = True
                get_public_ip = .Replace(ip_adr, vbNullString)
            End With
        Else
            get_public_ip = "-1"
        End If
    End With
End Function
