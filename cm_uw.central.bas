Attribute VB_Name = "central"
Option Compare Database
Option Explicit

Public Sub err_handler(ByVal input_proc_name As String _
, vba_err_no As Long _
, vba_err_desc As String _
, str_milestone As String _
, str_params As String _
, stella_err_desc As String _
, show_err As Boolean)
    Const proc_name As String = "central.err_handler"
    utilities.call_stack_add_item proc_name
    On Error Resume Next
    
    Dim cmd As ADODB.Command
    Dim app_name As String
    Dim app_continent As String
    
    If load.is_debugging = True Then
        GoTo outro
    End If
    
    app_name = "-1"
    app_name = load.system_info.app_name
    app_continent = "-1"
    app_continent = load.system_info.app_continent
    
    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = load.conn
        
        .CommandText = "INSERT INTO log_errors_t " & _
        "(system_error_text, system_error_code, stella_error_text, routine_name, call_stack, params, milestone, uw_name, app_name, file_path, app_continent) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        
        .CommandType = adCmdText

        ' Append parameters in the same order as the placeholders
        .Parameters.Append .CreateParameter("pSystemErrorText", adVarWChar, adParamInput, 255, vba_err_desc)
        .Parameters.Append .CreateParameter("pSystemErrorCode", adInteger, adParamInput, , vba_err_no)
        .Parameters.Append .CreateParameter("pStellaErrorText", adVarWChar, adParamInput, 255, stella_err_desc)
        .Parameters.Append .CreateParameter("pRoutineName", adVarWChar, adParamInput, 255, input_proc_name)
        .Parameters.Append .CreateParameter("pCallStack", adLongVarWChar, adParamInput, -1, load.call_stack)
        .Parameters.Append .CreateParameter("pParams", adVarWChar, adParamInput, 255, str_params)
        .Parameters.Append .CreateParameter("pMilestone", adVarWChar, adParamInput, 255, str_milestone)
        .Parameters.Append .CreateParameter("pUwName", adVarWChar, adParamInput, 255, Environ("Username"))
        .Parameters.Append .CreateParameter("pAppName", adVarWChar, adParamInput, 255, load.system_info.app_name)
        .Parameters.Append .CreateParameter("pFilePath", adVarWChar, adParamInput, 255, CurrentProject.FullName)
        .Parameters.Append .CreateParameter("pAppContinent", adVarWChar, adParamInput, 255, load.system_info.app_continent)
        .Execute
    End With

    Set cmd = Nothing
    
    Err.Clear
    
    '16 October 2025, CK: I don't think the below is require anymore. _
    However, due to recent issues with this in the US, I am leaving it for now in case it is needed. But to be removed next time I read this :)
    
    If show_err = True Then
        MsgBox load.system_info.error_instruction & vbNewLine & vbNewLine _
        & "Error description: " & vba_err_desc & vbNewLine _
        & "Where: " & input_proc_name & vbNewLine _
        & "Parameters: " & str_params & vbNewLine _
        & "App: " & load.system_info.app_name _
        & vbNewLine & "Call stack: " & Right(call_stack, 500) _
        , , load.system_info.error_msg_heading
    End If

outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Debug.Print "Error in " & proc_name & " - err.number: " & Err.Number & " - err.description: " & Err.Description
    Resume outro
End Sub
Public Sub check_deal(ByVal deal_id As Long)
    If global_vars.deal Is Nothing Then
        Set global_vars.deal = New cls_deal
    End If
    
    If global_vars.deal.is_init = False Then
        global_vars.deal.init deal_id
    End If
End Sub
Public Sub change_policy_from_stella_uw(ByVal deal_id As Long, ByVal policy_id As Long, ByVal app_continent As String)
    Dim proc_name As String
    proc_name = "central.change_policy_from_stella_uw"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    central.load_cm deal_id, app_continent
    open_forms.policies_f deal_id
    load.policies_f.change_policy deal_id, policy_id
End Sub
Public Sub country_procedures_add_update(ByVal deal_id As Long, ByVal first_run As Boolean)
    Dim proc_name As String
    proc_name = "central.country_procedures_add_update"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim col_target_procs As Collection
    Dim procedure As Scripting.Dictionary
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    Dim str_target_procedures
    Dim str_actual_procedures
    
    'add target procedures to collection
    Set col_target_procs = Nothing
    
    str_sql = "SELECT template_id, jurisdiction_id__jurisdictions_t, country_procedure, jurisdiction" _
    & " FROM " & load.sources.cm_country_procedures_templates_view _
    & " WHERE jurisdiction_id__jurisdictions_t = " & global_vars.deal.insured_registered_country_id
    
    Set col_target_procs = Nothing
    Set col_target_procs = New Collection
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        If CLng(.RecordCount) = 0 And first_run = False Then
            'no target procedures, and deal has been run before, so need to remove potential procedures
            str_sql = "UPDATE " & load.sources.cm_country_procedures_deals_table & " SET is_deleted = 1 WHERE deal_id__deals_t = " & deal_id
            conn.Execute str_sql
            load.rs_counter = load.rs_counter + 1
        End If
        
        If CLng(.RecordCount) = 0 Then
            'no target procedures, and deal has not been run before, so nothing to do
            GoTo outro
        End If
        
        .MoveFirst
        Do Until .EOF
            Set procedure = New Scripting.Dictionary
            With procedure
                .Add "template_procedure_id", rs!template_id.value
                .Add "country_procedure", rs!country_procedure.value
                .Add "jurisdiction_id", rs!jurisdiction_id__jurisdictions_t.value
                .Add "jurisdiction", rs!jurisdiction.value
            End With
            col_target_procs.Add procedure
            .MoveNext
        Loop
        
        .Close
    End With
    
    str_target_procedures = ""
    For Each procedure In col_target_procs
        str_target_procedures = str_target_procedures & procedure("country_procedure")
    Next procedure
        
    'prepare collection of actual procedures
    If first_run = False Then
        str_sql = "SELECT deal_procedure_id, country_procedure" _
        & " FROM " & load.sources.cm_country_procedures_deals_view _
        & " WHERE deal_id__deals_t = " & deal_id
        
        Set rs = utilities.create_adodb_rs(conn, str_sql)
        str_actual_procedures = ""
        With rs
            .Open
            Do Until .EOF
                str_actual_procedures = str_actual_procedures & !country_procedure
                .MoveNext
            Loop
            .Close
        End With
            
        If str_actual_procedures <> str_target_procedures Then
            If str_actual_procedures <> "" Then
                str_sql = "UPDATE " & load.sources.cm_country_procedures_deals_table & " SET is_deleted = 1 WHERE deal_id__deals_t = " & deal_id
                conn.Execute str_sql
                load.rs_counter = load.rs_counter + 1
            End If
            
            For Each procedure In col_target_procs
                str_sql = "INSERT INTO " & load.sources.cm_country_procedures_deals_table _
                & " (deal_id__deals_t, template_id__cm_country_procedures_templates_t)" _
                & " VALUES(" & deal_id & ", " & procedure("template_procedure_id") & ")"
                
                conn.Execute str_sql
                load.rs_counter = load.rs_counter + 1
            Next procedure
        End If
    Else
        For Each procedure In col_target_procs
            str_sql = "INSERT INTO " & load.sources.cm_country_procedures_deals_table _
            & " (deal_id__deals_t, template_id__cm_country_procedures_templates_t)" _
            & " VALUES(" & deal_id & ", " & procedure("template_procedure_id") & ")"
            
            conn.Execute str_sql
            load.rs_counter = load.rs_counter + 1
        Next procedure
    End If
    
outro:
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then rs.Close
        Set rs = Nothing
    End If
    
    Exit Sub
    
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

End Sub
Public Sub limit_premium_and_summary_central(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "central.limimt_premium_and_summary_central"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_secondary_access_app
    With load.secondary_access_app
        .Visible = False
        On Error Resume Next
        .CloseCurrentDatabase
        On Error GoTo err_handler
        If load.is_debugging = True Then On Error GoTo 0
        
        .OpenCurrentDatabase load.system_info.system_paths.scripts_path, False
        .Visible = False
        .Run "limit_premium_and_summary_central", deal_id, load.system_info.app_continent
        .CloseCurrentDatabase
        .OpenCurrentDatabase load.system_info.system_paths.common_path & "placeholder.accdb", False
        .Visible = False
    End With
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub search_in_deal_questions_f(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "central.search_in_deal_questions_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim order_by() As String
    Dim rs As ADODB.Recordset
    Dim str_form As String
    Dim str_recordsource() As String
    Dim str_sql As String
    
    str_form = "deal_questions_f"
    If IsNull(Forms(str_form)!header_free_text_search) = True Or Forms(str_form)!header_free_text_search = "" Then
        fix_rs.deal_questions_f Forms(load.control_panel.form_name)!deal_id
        GoTo outro
    End If
    
    order_by = Split(Forms(str_form).RecordSource, "ORDER BY")
    str_recordsource = Split(Forms(str_form).RecordSource, "WHERE")
    str_sql = str_recordsource(0) _
    & " WHERE deal_id = " & deal_id _
    & " AND question LIKE '%" & Forms(str_form)!header_free_text_search & "%' " _
    & " AND (" _
        & "question_type_id = " & question_types.binder _
        & " OR question_type_id = " & question_types.confirmation _
        & " OR question_type_id = " & question_types.inar _
    & ")" _
    & " ORDER BY " & order_by(1)
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        fix_rs.deal_questions_f Forms(str_form)!header_deal_id, rs
        .Close
    End With
    Set rs = Nothing
    
outro:
    load.has_searched = True
    Exit Sub

End Sub

Public Sub load_cm(ByVal deal_id As Long, Optional ByVal input_app_continent As String)
    Const proc_name As String = "central.load_cm"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    load.event_id = utilities.get_event_id
    
    Dim app_continent As String
    
    If input_app_continent <> "" Then load.system_info.app_continent = input_app_continent
    
    load.init_conn
    load.check_conn_and_variables
    
    central.load_cm_internal deal_id
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "", "deal_id = " & deal_id, "", True
    Resume outro
End Sub

Public Sub load_cm_internal(ByVal deal_id As Long)
    Const proc_name As String = "central.load_cm_internal"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim first_run As Boolean
    Dim log_object As utilities.typ_log_object
    Dim log_object_blank As utilities.typ_log_object
    Dim policy As cls_policy
    Dim re_run_cm As Boolean
    Dim risk_key_data As cls_deal
    Dim rs As ADODB.Recordset
    Dim str_milestone As String
    Dim str_sql As String
    Dim timer_start As Single
    
    timer_start = Timer
    
    'Check if cm has been run for this deal and/or shall be re-inited
    re_run_cm = False
    first_run = False
    Set risk_key_data = New cls_deal
    risk_key_data.populate_key_data deal_id
    
    If risk_key_data.is_deal_found = False Then
        re_run_cm = True
        first_run = True
    End If
    
    '28 March 2025, CK: Need to decide whether risk_object is created here and then passed around, or whether it's a global object
    '29 April 2025, CK: Decided to implement a new deal_object built from scratch that is a global variable
    str_milestone = "before risk_object.init deal_id"
    Set global_vars.deal = New cls_deal
    With global_vars.deal
        .init_deal_data deal_id
        .init_policy_data deal_id
    End With
    
    If risk_key_data.deal_currency <> global_vars.deal.deal_currency Then
        global_vars.deal.has_currency_changed = True
    End If
    
    If first_run = False Then
        For Each policy In global_vars.deal.col_policies
            policy.init_binder_data
        Next policy
    End If
    
    'check if deal is locked.
    If global_vars.deal.is_locked = True Then
        GoTo paint_control_panel
    End If
    
    'Check if cm shall be re-run for risk due to changes centrally
    If risk_key_data.cm_shall_re_run = 1 Then
        re_run_cm = True
        
        MsgBox "Due to changes from RPMA centrally, the CM for this risk needs re-running." _
        & " This will happen now.", , "CM will re-run now"
        
        str_sql = "UPDATE " & load.sources.cm_key_data_table & " SET shall_re_run = 0 WHERE deal_id = " & deal_id
        
        log_object = log_object_blank
        With log_object
            .change_source = proc_name
            .changer_id = Environ("username")
            .data_set = load.sources.cm_key_data_table
            .deal_id = deal_id
            .event_id = load.event_id
        End With
        utilities.log_change log_object
        
        conn.Execute str_sql
        load.rs_counter = load.rs_counter + 1
    End If
        
    'CM has been run for this risk before. But have something changed?
    Dim update_cm As Integer
    Dim text_for_box As String
    If central.has_key_deal_data_changed(deal, risk_key_data) = True And re_run_cm = False Then
        
        text_for_box = "Something has changed for this risk. Re-run the CM?"
        update_cm = MsgBox(text_for_box, vbQuestion + vbYesNo + vbDefaultButton2, "Re-run CM?")
        If update_cm = vbYes Then
            re_run_cm = True
        End If
    End If
    
    If re_run_cm = True Then
        central.init_deal global_vars.deal, first_run
    End If
    
paint_control_panel:
    
    load.control_panel.init
    open_forms.control_panel_f deal_id
    
    'Check if there are binders on the del
    central.check_status_area_on_control_panel_f deal_id
    If global_vars.deal.binder_count = 0 Then
        MsgBox "The Compliance Module did not find any binders for this deal." _
        & vbNewLine & vbNewLine & "There might be some though - check with your regional lead before declining.", , "No binders found"
    End If
    
    open_forms.deal_questions_f deal_id
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Debug.Print "load.rs_counter at end of central.load_cm_internal: " & load.rs_counter
    Debug.Print "Timer at end of central.load_cm_internal: " & Timer - timer_start
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, str_milestone, "deal_id = " & deal_id, "", True
    Resume outro
End Sub
Public Sub remove_binders_from_col_policy_binders(ByVal policy_id As Long)
    Dim proc_name As String
    proc_name = "central.remove_binders_from_col_policy_binders"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim binder As Scripting.Dictionary
    Dim i As Long
    
    i = 1
    For Each binder In global_vars.col_policy_binders
        If binder(binder_data.policy_id) = policy_id Then
            global_vars.col_policy_binders.Remove i
            i = i - 1
        End If
        i = i + 1
    Next binder

outro:
    Exit Sub

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
End Sub
Public Sub add_refresh_binder_list(ByRef policy As cls_policy, ByVal first_run As Boolean)
    Const proc_name As String = "central.add_refresh_binder_list"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    'purpose: check if binders are added to policy. If not, add the template selection
    
    Dim actual_binders As Collection
    Dim binder As cls_binder
    Dim binder_actual As Scripting.Dictionary
    Dim binder_target As cls_binder
    Dim log_object As utilities.typ_log_object
    Dim log_object_blank As utilities.typ_log_object
    Dim missing_binders As Collection
    Dim policy_id As Long
    Dim policy_test As cls_policy
    Dim redundant_binders As Collection
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    Dim str_actual_binders As String
    Dim str_target_binders As String
    Dim target_binders As Collection
        
    policy_id = policy.policy_id
    
    'check if atual binders match target binders
    Set actual_binders = New Collection
    For Each binder In policy.col_binders
        actual_binders.Add binder
    Next binder
    
    Set target_binders = central.get_target_binders_for_policy(policy)
    
    If target_binders.Count = 0 Then
        Set policy_test = global_vars.deal.get_policy(policy.policy_id)
        
        '18 November 2025, CK: If there are no target binders, but one or more are added manually, the manually added binders should remain.
        str_sql = ""
        For Each binder In policy.col_binders
            If binder.manually_added = "no" Then
                If str_sql <> "" Then str_sql = str_sql & ", "
                str_sql = str_sql & binder.binder_id
            End If
        Next binder
        
        If str_sql <> "" Then
            str_sql = "UPDATE " & sources.policy_binders_table & " SET is_deleted = 1 WHERE binder_id IN (" & str_sql & ") AND policy_id = " & policy_id
            conn.Execute str_sql
            load.rs_counter = load.rs_counter + 1
            
            '18 November 2025, CK: Ideally, each removed binder is logged individually.
            log_object = log_object_blank
            With log_object
                .change_source = proc_name
                .changer_id = Environ("username")
                .data_set = load.sources.policy_binders_table
                .deal_id = global_vars.deal.deal_id
                .event_id = load.event_id
                .executed_sql = str_sql
                .field_name = "is_deleted"
                .new_value = 1
                .operation_type = "update"
                .policy_id = policy.policy_id
            End With
            utilities.log_change log_object
            
            global_vars.deal.init_policy_data global_vars.deal.deal_id
            policy_test.init_binder_data
        End If
        
        If policy_test.col_binders.Count = 0 Then
            GoTo outro
        End If
    End If
    
    'add new binders
    Set missing_binders = New Collection
    For Each binder In target_binders
        If policy.is_binder_on_policy(binder.binder_id) = False Then
            binder.policy_id = policy.policy_id
            binder.on_policy_id = menu_list.yes
            policy.col_binders.Add binder
            missing_binders.Add binder
        End If
    Next binder
    
    Dim binder_found As Boolean
    Set redundant_binders = New Collection
    
    For Each binder In policy.col_binders
        binder_found = False
        For Each binder_target In target_binders
            If binder.binder_id = binder_target.binder_id Then
                binder_found = True
            End If
        Next
        If binder_found = False And binder.manually_added = "no" Then
            redundant_binders.Add binder
        End If
    Next binder
    
    'remove redundant binders from MySQL database
    str_sql = ""
    For Each binder In redundant_binders
        If str_sql = "" Then
            str_sql = "binder_id = " & binder.binder_id
        Else
            str_sql = str_sql & " OR binder_id = " & binder.binder_id
        End If
    Next binder
    
    If str_sql <> "" Then
        str_sql = "UPDATE " & sources.policy_binders_table & " SET is_deleted = 1 WHERE policy_id = " & policy_id & " AND (" & str_sql & ")"
        conn.Execute str_sql
        load.rs_counter = load.rs_counter + 1
        
        log_object = log_object_blank
        With log_object
            .changer_id = Environ("username")
            .change_source = proc_name
            .comment = "logging policy_id as record_id even though binder removal is on security_t."
            .data_set = load.sources.policy_binders_table
            .deal_id = global_vars.deal.deal_id
            .event_id = load.event_id
            .executed_sql = str_sql
            .field_name = "security_t.is_deleted"
            .new_value = "1"
            .policy_id = policy.policy_id
            .operation_type = "update"
        End With
        utilities.log_change log_object
    End If
        
    'add missing binders to MySQL database
    str_sql = ""
    For Each binder In missing_binders
        If str_sql <> "" Then
            str_sql = str_sql & ", "
        End If
        
        str_sql = str_sql & "(" & policy_id _
        & ", " & binder.binder_id & ", 'stella.birkelund', 'no')"
    Next binder
    
    If str_sql <> "" Then
        str_sql = "INSERT INTO " & load.sources.policy_binders_table & " (policy_id, binder_id, added_by, manually_added)" _
        & " VALUES " & str_sql
    
        conn.Execute str_sql
        load.rs_counter = load.rs_counter + 1
        
        log_object = log_object_blank
        With log_object
            .changer_id = Environ("username")
            .change_source = proc_name
            .comment = "logging policy_id as record_id even though the data entry is on binder level."
            .data_set = load.sources.policy_binders_table
            .deal_id = global_vars.deal.deal_id
            .event_id = load.event_id
            .executed_sql = str_sql
            .field_name = ""
            .new_value = ""
            .operation_type = "insert"
            .policy_id = policy.policy_id
            .record_id = policy.policy_id
        End With
        utilities.log_change log_object
    End If
    
    If global_vars.deal.has_currency_changed = True Then
        For Each binder In policy.col_binders
            central.add_refresh_binder_limits binder
        Next binder
    Else
        For Each binder In missing_binders
            central.add_refresh_binder_limits binder
        Next binder
    End If
    
    central.limit_distribution policy
    fix_rs.binder_list_f
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description _
    , "", "policy.policy_id = " & policy.policy_id & "first_run = " & first_run, "", True
    Resume outro
End Sub
Public Function get_target_questions_for_deal(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "central.get_target_questions_for_deal"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String, rs As ADODB.Recordset, i As Integer
    Dim question_list() As Variant, actual_binders()
    
    actual_binders = utilities.get_binder_list_for_deal(deal_id)
    If actual_binders(0) = -1 Then
        get_target_questions_for_deal = actual_binders
        GoTo outro
    End If
   
    'Select questions for these binders
    str_sql = "SELECT DISTINCT(template_question_id) FROM " & load.sources.binder_questions_view _
    & " WHERE is_retired_id__cm_binder_questions_t = 94 " _
    & " AND is_retired_id = 94" _
    & " AND (" _
        & "question_type_id = " & question_types.binder _
        & " OR question_type_id = " & question_types.auto_referral _
    & ")" _
    & " AND ("
    For i = 1 To UBound(actual_binders)
        If i = 1 Then
            str_sql = str_sql & " binder_id = " & actual_binders(i)
        Else
            str_sql = str_sql & " OR binder_id = " & actual_binders(i)
        End If
    Next i
    str_sql = str_sql & ") ORDER BY template_question_id ASC"
    
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        ReDim question_list(0 To CLng(rs.RecordCount))
        question_list(0) = rs.RecordCount
        For i = 1 To CLng(rs.RecordCount)
            question_list(i) = rs!template_question_id
            rs.MoveNext
        Next i
    rs.Close
    get_target_questions_for_deal = question_list
    
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
        .routine_name = "central.get_target_questions_for_deal"
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro

End Function
Public Function get_target_binders_for_policy(ByVal policy As cls_policy) As Collection
    Dim proc_name As String
    proc_name = "central.get_target_binders_for_policy"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim binder As cls_binder
    Dim binder_policy As cls_binder
    Dim boo_add As Boolean
    Dim boo_state_exists As Boolean
    Dim col_output As Collection
    Dim condition_is_ok As Boolean
    Dim counter As Long
    Dim deal_date As Date
    Dim i As Long
    Dim policy_id As Long
    Dim security_id As Long
    Dim states_found As Boolean
    Dim str_continent_check As String
    Dim str_sql As String
    Dim sub_counter As Long
    
    policy_id = policy.policy_id
    
    'loop through all binders and check them against conditions
    Set col_output = New Collection
    
    For Each binder In global_vars.col_all_binders
        boo_add = True
        
        'continent check
        If load.system_info.app_continent = load.system_info.continents.eurasia Then
            If binder.for_eur_id = menu_list.no Then
                boo_add = False
                GoTo end_of_binder_check
            End If
        End If
        If load.system_info.app_continent = load.system_info.continents.americas Then
            If binder.for_us_id = menu_list.no Then
                boo_add = False
                GoTo end_of_binder_check
            End If
        End If
        
        'Condition: check if binder is available for insured's registered country
        condition_is_ok = False
        For i = 1 To UBound(load.binder_jurisdictions_array)
            If load.binder_jurisdictions_array(i, 1) = binder.binder_id _
            And load.binder_jurisdictions_array(i, 2) = global_vars.deal.insured_registered_country_id Then
                condition_is_ok = True
            End If
        Next i
        If condition_is_ok = False Then
            GoTo end_of_binder_check
        End If
        
        'Condition: Are there states for the insured's registered country, and if yes, are binders available for the state/province?
        condition_is_ok = False
        boo_state_exists = False
        
        For counter = 1 To load.jurisdictions.Count
            If load.jurisdictions(counter)("parent_jurisdiction_id") = global_vars.deal.insured_registered_country_id Then
                If global_vars.deal.insured_main_region_id = placeholder_state Then
                    counter = load.jurisdictions.Count
                    condition_is_ok = True
                End If
            
                boo_state_exists = True
                For sub_counter = 1 To UBound(load.binder_jurisdictions_array)
                    If load.binder_jurisdictions_array(sub_counter, 1) = binder(binder_data.binder_id) _
                    And load.binder_jurisdictions_array(sub_counter, 2) = global_vars.deal.insured_main_region_id Then
                        condition_is_ok = True
                    End If
                Next sub_counter
            End If
        Next counter
        If boo_state_exists = False Then
            condition_is_ok = True
        End If
        
        If condition_is_ok = False Then
            GoTo end_of_binder_check
        End If

        'Condition: ev
        'TBD
        
        'Condition: start date and end date of binder
        If boo_add = True Then
            If global_vars.deal.inception_date = -1 Then
                deal_date = Date
            Else
                deal_date = global_vars.deal.inception_date
            End If
            If binder.start_date > deal_date _
            Or binder.end_date < deal_date Then
                boo_add = False
            End If
        End If
        
        'Condition: risk_type
        If boo_add = True Then
            condition_is_ok = False
            For counter = 1 To load.risk_types.Count
                If load.risk_types(counter)("binder_id") = binder.binder_id _
                And load.risk_types(counter)("risk_type_id") = global_vars.deal.risk_type_id Then
                    condition_is_ok = True
                End If
            Next counter
        End If
        If condition_is_ok = False Then
            boo_add = False
            GoTo end_of_binder_check
        End If
        
        'condition: check of underlying_limit
        For Each policy In global_vars.deal.col_policies
            If policy.policy_id = policy_id Then
                If policy.underlying_limit < binder.minimum_underlying_limit * global_vars.deal.deal_currency_to_eur Then
                    boo_add = False
                    Exit For
                End If
            End If
        Next policy
        
        If boo_add = True Then
            Set binder_policy = New cls_binder
            Set binder_policy = binder
            col_output.Add binder_policy
        End If
            
end_of_binder_check:
            
    Next binder
    
    If col_output.Count > 0 Then global_vars.deal.any_binders_on_policy = True
    Set get_target_binders_for_policy = col_output
    
outro:
    Exit Function
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "central.get_target_binders_for_policy"
        .milestone = ""
        .params = "policy_id = " & policy_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
    
End Function

Public Sub limit_distribution(ByRef policy As cls_policy)
    Const proc_name As String = "central.limit_distribution"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    'Purpose: Allocate limit per binder on policy
    
    Dim binder As cls_binder
    Dim default_quota As Double
    Dim limit_for_default_allocation As Currency
    Dim max_limit_for_default_allocation_deal_currency As Currency
    Dim policy_id As Long
    Dim quota As Double
    Dim quota_for_default_alloation As Double
    Dim rounding_level As Integer
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    Dim total_allocated_limit As Currency
    Dim total_manual_limit As Currency
    Dim total_reference_limit_without_binders_with_manual_limit As Currency
    Dim use_reference_limit As Boolean
    
    policy_id = policy.policy_id
    
    If policy.col_binders.Count = 0 Then
        GoTo outro
    End If
    
    If policy.get_total_manual_limit = policy.policy_limit Then
        '14 May 2023, CK: If this is 0, all security has manual limit. Hence, the CM should make sure the quotas reflect that, but not make any limit adjustments.
        GoTo outro
    End If
    
    limit_for_default_allocation = policy.get_limit_for_default_allocation
    max_limit_for_default_allocation_deal_currency = policy.get_max_limit_for_default_allocation_deal_currency
    quota_for_default_alloation = policy.get_quota_for_default_alloation
    total_reference_limit_without_binders_with_manual_limit = policy.get_total_reference_limit_without_binders_with_manual_limit
    
    'if we are near max limit, allocate limit based on max available limit and do nothing else
    If limit_for_default_allocation > max_limit_for_default_allocation_deal_currency * 0.9 Then
        policy.binder_allocation_unrounded
        GoTo outro
    End If
    
    use_reference_limit = policy.use_reference_limit
    
    total_allocated_limit = 0
    
    For Each binder In policy.col_binders
        default_quota = 0
        If binder.on_policy_id = menu_list.yes Then
            If binder.manual_limit = 0 Then
                If use_reference_limit = False Then
                    quota = binder.max_limit_deal_currency / max_limit_for_default_allocation_deal_currency * quota_for_default_alloation
                Else
                    quota = binder.reference_limit / total_reference_limit_without_binders_with_manual_limit * quota_for_default_alloation
                End If
                
                'round allocation
                rounding_level = 3
                If policy.binder_count_without_manual_limit > 1 Then
                    quota = Round(quota, rounding_level)
                End If
            
                total_allocated_limit = total_allocated_limit + quota * policy.policy_limit
            Else
                quota = binder.manual_limit / policy.policy_limit
                
                total_allocated_limit = total_allocated_limit + quota * policy.policy_limit
            End If
            binder.default_quota = quota
            binder.binder_quota = quota
        End If
    Next binder
    
    'check if all limit is allocated. If not, allocate residual to binder with highest limit on deal that is not set manually.
    Dim allocated_quota As Double
    Dim binder_with_higest_limit As cls_binder
    Dim current_limit_of_binder_with_highest_limit As Currency
    Dim highest_binder_limit As Currency
    Dim id_of_binder_with_highest_limit As Long
    
    If total_allocated_limit <> policy.policy_limit Then
        Set binder_with_higest_limit = policy.get_binder_with_highest_limit
        
        'calculate allocated quota without rounding binder
        allocated_quota = 0
        For Each binder In policy.col_binders
            If binder.on_policy_id = menu_list.yes Then
                If binder.binder_id <> binder_with_higest_limit.binder_id Then
                    allocated_quota = allocated_quota + binder.binder_quota
                End If
            End If
        Next binder
                  
        quota = 1 - allocated_quota
        policy.update_binder_quota Replace(CStr(quota), ",", "."), binder_with_higest_limit.binder_id
        policy.update_default_quota Replace(CStr(quota), ",", "."), binder_with_higest_limit.binder_id
    End If
    
    policy.send_binder_data_to_mysql
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
    
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "", "policy.policy_id = " & policy.policy_id, "", True
    Resume outro
End Sub
Public Sub add_refresh_deal_questions(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "central.add_refresh_deal_questions"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    Dim str_sql As String, rs As ADODB.Recordset, i As Long
    Dim target_questions(), actual_questions()
    
    target_questions = central.get_target_questions_for_deal(deal_id)
    actual_questions = utilities.get_actual_questions_for_deal(deal_id)
    
    If actual_questions(0) > 0 Then
        If Join(target_questions, Chr(0)) = Join(actual_questions, Chr(0)) Then
            'target_questions and actual_questions are the same
            GoTo outro
        End If
    End If
    
    'remove binder questions that should not be on
    str_sql = ""
    For i = 1 To UBound(actual_questions)
        If utilities.is_in_array(CLng(actual_questions(i)), target_questions) = False Then
            If str_sql = "" Then
                str_sql = "template_question_id = " & actual_questions(i)
            Else
                str_sql = str_sql & " OR template_question_id = " & actual_questions(i)
            End If
        End If
    Next i
    If str_sql <> "" Then
        str_sql = "UPDATE " & sources.deal_questions_table & " SET is_deleted = 1 WHERE deal_id = " & deal_id & " AND (" & str_sql & ")"
        conn.Execute str_sql
        load.rs_counter = load.rs_counter + 1
    End If
    
    'add missing binder questions
    Dim default_answer As Long
    str_sql = ""
    For i = 1 To UBound(target_questions)
        If utilities.is_in_array(CLng(target_questions(i)), actual_questions) = False Then
            default_answer = utilities.get_default_answer_for_template_question(target_questions(i))
            If str_sql = "" Then
                str_sql = "(" & deal_id & ", " & target_questions(i) & ", " & default_answer & ")"
            Else
                str_sql = str_sql & ", (" & deal_id & ", " & target_questions(i) & ", " & default_answer & ")"
            End If
        End If
    Next i
    
    If str_sql <> "" Then
        str_sql = "INSERT INTO " & sources.deal_questions_table & " (deal_id, template_question_id, answer_id) VALUES" & str_sql
        conn.Execute str_sql
        load.rs_counter = load.rs_counter + 1
    End If
    
    'Update recordsource of relevant form
    fix_rs.deal_questions_f deal_id
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "central.add_refresh_deal_questions"
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
    
End Sub
Public Sub add_inar_questions_and_confirmations_to_deal(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "central.add_inar_questions_and_confirmations_to_deal"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    Dim str_sql As String, rs As ADODB.Recordset, i As Integer
    
    'remove old ARs
    '30 March 2025, CK: As the question_type is not saved in the table, the below code is required to find the relevant deal questions to remove.
    str_sql = "SELECT id, question_type_id FROM " & load.sources.deal_questions_view & " WHERE deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        str_sql = ""
        Do Until .EOF
            If !question_type_id = question_types.inar _
            Or !question_type_id = question_types.confirmation _
            Or !question_type_id = question_types.confirmation_uw_positions Then
                If str_sql = "" Then
                    str_sql = "id = " & !id
                Else
                    str_sql = str_sql & " OR " & !id
                End If
            End If
            .MoveNext
        Loop
        .Close
        If str_sql <> "" Then str_sql = "(" & str_sql & ")"
    End With

    If str_sql <> "" Then

        str_sql = "UPDATE " & load.sources.deal_questions_table & " SET is_deleted = 1 " _
        & " WHERE " & str_sql & " AND deal_id = " & deal_id

        conn.Execute str_sql
        load.rs_counter = load.rs_counter + 1
    End If

    'add InAR questions
    Dim str_continent_condition As String
    str_continent_condition = ""
    If load.system_info.app_continent = load.system_info.continents.eurasia Then str_continent_condition = " AND for_eur_id = " & menu_list.yes
    If load.system_info.app_continent = load.system_info.continents.americas Then str_continent_condition = " AND for_us_id = " & menu_list.yes
    
    str_sql = "SELECT id, default_answer_id FROM " & sources.template_qs_table _
    & " WHERE (" _
        & "question_type_id = " & question_types.inar _
        & " OR question_type_id__cm_question_types_t = " & question_types.confirmation _
        & " OR question_type_id__cm_question_types_t = " & question_types.confirmation_uw_positions _
    & ") AND is_deleted = 0" _
    & " AND is_retired_id = " & menu_list.no _
    & str_continent_condition
    
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        str_sql = ""
        Do Until rs.EOF
            If str_sql = "" Then
                str_sql = "(" & rs!id & ", " & deal_id & ", " & rs!default_answer_id & ")"
            Else
                str_sql = str_sql & ", (" & rs!id & ", " & deal_id & ", " & rs!default_answer_id & ")"
            End If
            rs.MoveNext
        Loop
    rs.Close
    str_sql = "INSERT INTO " & sources.deal_questions_table & " (template_question_id, deal_id, answer_id) VALUES " & str_sql
    conn.Execute str_sql
    load.rs_counter = load.rs_counter + 1
    
'    'add confirmations
'    str_sql = "SELECT id, default_answer_id FROM " & load.sources.template_qs_table _
'    & " WHERE (" _
'        & "question_type_id__cm_question_types_t = " & question_types.confirmation _
'        & " OR question_type_id__cm_question_types_t = " & question_types.confirmation_uw_positions _
'    & ")" _
'    & " AND is_deleted = 0" _
'    & " AND is_retired_id = " & menu_list.no
'
'    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
'        Do Until rs.EOF
'            conn.Execute "INSERT INTO " & sources.deal_questions_table & " (template_question_id, deal_id, answer_id) VALUES (" & rs!id & ", " & deal_id & ", " & rs!default_answer_id & ")"
'            load.rs_counter = load.rs_counter + 1
'            rs.MoveNext
'        Loop
'    rs.Close
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "central.add_inar_questions_and_confirmations_to_deal"
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub remove_redundant_referrals(ByVal deal_id As Long)
    'purpose: Check if we have referrals for binders that are not on the deal. Remove if so.
    Dim proc_name As String
    proc_name = "central.remove_redundant_referrals"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String
    Dim binder_counter As Long
    Dim binder_found As Boolean
    Dim binder_list_for_question()
    Dim binder_list_for_deal()
    Dim target_referrals()
    Dim i As Integer
    Dim y As Integer
    Dim target_referral_count As Integer
    Dim rs As ADODB.Recordset
    Dim rs_check As ADODB.Recordset
    Dim end_loop As Boolean
    Dim skip_to_next_item As Boolean
    ReDim target_referrals(1 To 1)
    
    'check for binder referrals
    str_sql = "SELECT * FROM " & sources.referrals_table & " WHERE deal_id = " & deal_id & " AND is_deleted = 0"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    binder_list_for_deal = utilities.get_binder_list_for_deal(deal_id)
    With rs
        .Open
        If .RecordCount = 0 Then
            GoTo outro
        End If
        
        'check if referral is for a binder that is on deal
        
        .MoveFirst
        Do Until .EOF
            binder_found = False
            For binder_counter = 1 To binder_list_for_deal(0)
                If binder_list_for_deal(binder_counter) = rs!binder_id Then
                    binder_found = True
                    binder_counter = binder_list_for_deal(0)
                End If
            Next binder_counter
            
            If IsNull(rs!binder_id) = False And rs!binder_id <> "" And rs!binder_id <> "-1" Then
                If binder_found = False Then
                    str_sql = "UPDATE " & load.sources.referrals_table & " SET is_deleted = 1 WHERE deal_id = " & deal_id & " AND binder_id = " & rs!binder_id
                    conn.Execute str_sql
                End If
            End If
            .MoveNext
        Loop
        
    End With
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Sub

Public Sub check_referrals(ByVal deal_id As Long)
    'purpose: check status on referrals. Update referral_status in security_t.
    'How: Load referrals. Check status for each binder. If all referrals for a binders are resolved, set binder_status in security_t to ok. If not, set to pending referral.
    Dim proc_name As String
    proc_name = "central.check_referrals"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String, rs As ADODB.Recordset
    'get list of binders with referrals
    Dim binders()
    str_sql = "SELECT DISTINCT(binder_id) FROM " & sources.referrals_view & " WHERE question_type = " & question_types.binder & " AND deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        If rs.BOF And rs.EOF Then
            GoTo outro
        End If
        ReDim binders(1 To CLng(rs.RecordCount))
        Dim i As Integer
        For i = 1 To CLng(rs.RecordCount)
            binders(i) = rs!binder_id
            rs.MoveNext
        Next i
    rs.Close
    
    'check referrals for each binder. Outcome is one of three:
        'All referrals for a binder is ok. If so, set status for the binder on all policies on the deal to 'referral ok'. (487)
        'One or more referrals for a binder is denied. If so, set status for the binder on all policies on the deal to 'referral denied' (488)
        'No referrals are denied, but one or more ferrals are not completed. Set status for the binder on all policies on the deal to 'referral needed' (491)
    Dim rs_security As ADODB.Recordset, policy_list(), y
    For i = 1 To UBound(binders)
        str_sql = "SELECT * FROM " & sources.referrals_view & " WHERE deal_id = " & deal_id & " AND binder_id = " & binders(i)
        Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
            rs.MoveFirst
            Do Until rs.EOF
                policy_list = utilities.get_policy_list_for_deal(deal_id)
                
                str_sql = "UPDATE " & sources.policy_binders_table & " SET binder_compliant = " & rs!referral_status _
                & " WHERE binder_id = " & binders(i) & " AND (policy_id = " & policy_list(1)
                
                If UBound(policy_list) > 1 Then
                    For y = 2 To UBound(policy_list)
                        str_sql = str_sql & " OR policy_id = " & policy_list(y)
                    Next y
                End If
                str_sql = str_sql & ")"
                conn.Execute str_sql
                load.rs_counter = load.rs_counter + 1
                rs.MoveNext
            Loop
        rs.Close
    Next i
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub add_refresh_binder_limits(ByRef binder As cls_binder)
    'Purpose: Find max limit for a binder and add to security_t.
    Dim proc_name As String
    proc_name = "central.add_refresh_binder_limits"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim binder_currency As String
    Dim binder_limit As Scripting.Dictionary
    Dim default_currency As String
    Dim fx_date As String
    Dim fx_rate
    Dim max_limit_id As Long
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    'find max limit for that binder
    'First, check if there is max limit for deal_currency
    binder_currency = ""
    For Each binder_limit In global_vars.col_binder_limits
        If binder_limit(binder_data.max_binder_limit_currency) = global_vars.deal.deal_currency _
        And binder_limit(binder_data.binder_id) = binder.binder_id _
        Then
            binder_currency = binder_limit(binder_data.limit_currency)
        End If
    Next binder_limit
    
    'if no binder limit for that specific currency, find the default currency for the binder
    If binder_currency = "" Then binder_currency = binder.default_binder_currency
    
    For Each binder_limit In global_vars.col_binder_limits
        
        If binder_limit(binder_data.binder_id) = binder.binder_id _
        And binder_limit(binder_data.max_binder_limit_currency) = binder_currency Then
        
            max_limit_id = binder_limit("binder_limit_id")
            binder.max_limit_id = max_limit_id
            binder.max_binder_limit_ex_extra = binder_limit(binder_data.max_binder_limit)
            binder.max_binder_limit_currency = binder_limit(binder_data.max_binder_limit_currency)
        End If
    Next binder_limit
    
    'add fx to EUR
    '26 May 2025, CK: Pick rates for currency date
    fx_rate = global_vars.col_currencies(1)(global_vars.deal.deal_currency) / global_vars.col_currencies(1)(binder_currency)
    binder.fx_deal_binder_rate = fx_rate
    
    '26 May 2026: CK YOU ARE HERE!!!!!
    
'    'add fx to EUR
'    str_sql = "SELECT " & risk_object.deal_currency & " deal_currency, " & binder_currency & " binder_currency" _
'    & " FROM " & sources.currencies_table _
'    & " WHERE currency_date <= current_date" _
'    & " ORDER BY currency_date DESC LIMIT 1"
'
'    Set rs = utilities.create_adodb_rs(conn, str_sql)
'    rs.Open
'        fx_rate = rs!deal_currency / rs!binder_currency
'        fx_rate = Replace(CStr(fx_rate), ",", ".")
'    rs.Close
'    Set rs = Nothing
'
'    str_sql = "UPDATE " & sources.policy_binders_table & " SET max_limit_id = " & max_limit_id _
'    & ", max_limit_currency_to_deal_currency_fx = " & fx_rate _
'    & " WHERE policy_id = " & binder(binder_data.policy_id) _
'    & " AND binder_id = " & binder(binder_data.binder_id)
'
'    conn.Execute str_sql
'
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "binder.security_id = " & binder.security_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Sub

Public Sub reset_deal(ByVal deal_id As Long)
    Const proc_name As String = "central.reset_deal"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    Dim policy_object As cls_policy_object
    
    'delete all binders on all policies
    If global_vars.deal Is Nothing Then
        Set global_vars.deal = New cls_deal
    End If
    If global_vars.deal.is_init = False Then
        global_vars.deal.init deal_id
    End If
    
    Dim policy As cls_policy
    For Each policy In global_vars.deal.col_policies
        conn.Execute "UPDATE " & sources.policy_binders_table & " SET is_deleted = 1 WHERE policy_id = " & policy.policy_id
        load.rs_counter = load.rs_counter + 1
    Next policy
    
    'delete all deal questions
    conn.Execute "UPDATE " & sources.deal_questions_table & " SET is_deleted = 1 WHERE is_deleted = 0 AND deal_id = " & deal_id
    load.rs_counter = load.rs_counter + 1
    
    'delete all referrals
    conn.Execute "UPDATE " & sources.referrals_table & " SET is_deleted = 1 WHERE is_deleted = 0 AND deal_id = " & deal_id
    load.rs_counter = load.rs_counter + 1
    
    'delete extra limits
    str_sql = "UPDATE " & sources.extra_deal_limit_table & " SET is_deleted = 1 WHERE deal_id = " & deal_id
    conn.Execute str_sql
    load.rs_counter = load.rs_counter + 1
    
    'delete historical data
    str_sql = "UPDATE " & load.sources.cm_key_data_table & " SET is_deleted = 1 WHERE deal_id = " & deal_id
    conn.Execute str_sql
    load.rs_counter = load.rs_counter + 1
    
    'dete country procedures
    str_sql = "UPDATE " & load.sources.cm_country_procedures_deals_table & " SET is_deleted = 1 WHERE deal_id__deals_t = " & deal_id
    conn.Execute str_sql
    load.rs_counter = load.rs_counter + 1
    
    'close forms with outdated data
    With CurrentProject
        If .AllForms("deal_questions_f").IsLoaded = True Then
            DoCmd.Close acForm, "deal_questions_f"
        End If
        If .AllForms("deal_referrals_f").IsLoaded = True Then
            DoCmd.Close acForm, "deal_referrals_f"
        End If
        If .AllForms("binder_list_f").IsLoaded = True Then
            DoCmd.Close acForm, "binder_list_f"
        End If
        If .AllForms("extra_deal_limit_f").IsLoaded = True Then
            DoCmd.Close acForm, "extra_deal_limit_f"
        End If
    End With
    
    Set global_vars.deal = Nothing
    
    central.load_cm_internal deal_id
    
outro:
    Exit Sub
    
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
End Sub
Public Sub check_status_area_on_control_panel_f(ByVal deal_id, Optional ByVal rs_deal_questions As ADODB.Recordset)
    'Purpose: check and update each field in the status area of control_panel_f
    Dim proc_name As String
    proc_name = "central.check_status_area_on_control_panel_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim i As Long
    Dim input_control As cls_field
    Dim rs_policy_binders As ADODB.Recordset
    Dim rs_referrals As ADODB.Recordset
    Dim str_sql As String
    Dim timer_start As Single
    
    str_sql = "SELECT id, binder_id, question_type_id, referral_status_id, due_time_for_referral_id, template_question_id" _
    & " FROM " & load.sources.referrals_view _
    & " WHERE deal_id = " & deal_id
    
    Set rs_referrals = utilities.create_adodb_rs(conn, str_sql)
    rs_referrals.Open
    
    str_sql = "SELECT * FROM " & load.sources.policy_binders_view & " WHERE deal_id = " & deal_id & " ORDER BY id"
    Set rs_policy_binders = utilities.create_adodb_rs(conn, str_sql)
    rs_policy_binders.Open
    If rs_policy_binders.RecordCount = 0 Then
        load.control_panel.init_status_overview
        load.control_panel.init_binder_overview
        control_panel_m.refresh_control_panel_f deal_id, rs_policy_binders
        For Each input_control In load.control_panel.col_binder_overview
            input_control.field_visible = False
        Next input_control
        global_vars.deal.any_binders_on_policy = False
        GoTo outro
    Else
        global_vars.deal.any_binders_on_policy = True
    End If
        
    If rs_deal_questions Is Nothing Then
        str_sql = "SELECT * FROM " & sources.deal_questions_view & " WHERE deal_id = " & deal_id & " ORDER BY sort_order, template_question_category"
        Set rs_deal_questions = utilities.create_adodb_rs(conn, str_sql)
        rs_deal_questions.Open
    End If
    
    control_panel_m.check_deal_questions deal_id, rs_deal_questions
    control_panel_m.check_inar_referrals deal_id, rs_referrals
    control_panel_m.check_binder_referrals deal_id, rs_referrals, rs_policy_binders
    control_panel_m.check_binder_overview deal_id, rs_referrals, rs_policy_binders
    control_panel_m.flag_overlimited_binders load.control_panel.col_binder_max_limit_labels, load.control_panel.col_binder_limit_on_risk_labels
    control_panel_m.check_if_limits_are_allocated_for_all_policies deal_id, rs_policy_binders
    control_panel_m.check_country_procedures
    control_panel_m.check_nbi_ready
    control_panel_m.check_signing_ready
    load.control_panel.binder_and_policy_details.field_value = central.create_policy_summary(deal_id, rs_policy_binders, False)
    
outro:
    With load.control_panel
        utilities.paint_control load.control_panel.form_name, .col_headers
        utilities.paint_control load.control_panel.form_name, .col_headers_sub
        utilities.paint_control load.control_panel.form_name, .col_policy_overview
        utilities.paint_control load.control_panel.form_name, .col_status_overview
        utilities.paint_control load.control_panel.form_name, .col_binder_overview
    End With
    
    If Not rs_referrals Is Nothing Then
        If rs_referrals.State = 1 Then rs_referrals.Close
        Set rs_referrals = Nothing
    End If
    If Not rs_policy_binders Is Nothing Then
        If rs_policy_binders.State = 1 Then rs_policy_binders.Close
        Set rs_policy_binders = Nothing
    End If
    
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "central.check_status_area_on_control_panel_f"
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Sub

Public Sub fix_white_bg()
    If CurrentProject.AllForms("blank_f").IsLoaded = False Then
        DoCmd.OpenForm "blank_f"
    Else
        Forms("blank_f").SetFocus
        DoCmd.Maximize
    End If
    If CurrentProject.AllForms("control_panel_f").IsLoaded = True Then
        Forms("control_panel_f").SetFocus
    End If
End Sub
Public Function are_there_binders_on_policy(ByVal policy_id) As Boolean
    Dim proc_name As String
    proc_name = "central.are_there_binders_on_policy"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    'Purpose: check if a policy has allocated binders to it.
    Dim str_sql As String, rs As ADODB.Recordset
    str_sql = "SELECT id FROM " & load.sources.policy_binders_view & " WHERE policy_id = " & policy_id
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        If rs.BOF And rs.EOF Then
            are_there_binders_on_policy = False
            binders_on_policy = False
        Else
            are_there_binders_on_policy = True
            binders_on_policy = True
        End If
    rs.Close
    
outro:
    Set rs = Nothing
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
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
    
End Function
Public Function get_list_of_key_data()
    Dim proc_name As String
    proc_name = "central.get_list_of_key_data"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim col_key_fields As New Collection
    Set col_key_fields = Nothing
    With col_key_fields
        .Add "deal_currency"
        '.Add "ev"
        .Add "inception_date"
        .Add "insured_registered_country_id"
        .Add "insured_main_region_id"
        .Add "policy_period_in_months"
        .Add "risk_type_id"
        '.Add "spa_law"
        .Add "target_super_sector_id"
        .Add "target_sub_sector_id"
        If load.system_info.app_continent = load.system_info.continents.americas Then
            .Add "max_limit_quoted"
        End If
    End With
    
    Set get_list_of_key_data = col_key_fields
    
outro:
    Exit Function
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "central.get_list_of_key_data"
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Function
Public Function has_key_deal_data_changed(ByVal risk_object As cls_deal _
, ByVal cm_data As cls_deal) As Boolean
    'Purpose: check if data which dictates binders have changed since binders were allocated

    Dim proc_name As String
    proc_name = "central.has_key_deal_data_changed"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim col_key_fields As New Collection
    Dim data_field As Variant
    Dim data_deals_t As Variant
    Dim data_cm_deals_t As Variant
    
    Set col_key_fields = Nothing
    Set col_key_fields = central.get_list_of_key_data
    
    For Each data_field In col_key_fields
        data_deals_t = Trim(CallByName(risk_object, data_field, VbGet))
        data_cm_deals_t = Trim(CallByName(cm_data, data_field, VbGet))
        
        If data_deals_t <> data_cm_deals_t Then
            If data_field = "deal_currency" Then global_vars.deal.has_currency_changed = True
            has_key_deal_data_changed = True
            GoTo outro
        End If
    Next data_field

outro:
    Exit Function
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "central.has_key_data_been_changed"
        .milestone = ""
        .params = "deal_id = " & risk_object.deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Function
Public Sub add_key_data(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "central.add_key_data"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    Dim str_sql As String, rs As ADODB.Recordset, counter As Integer
    
    Dim col_key_data As New Collection
    Set col_key_data = central.get_list_of_key_data
    
    str_sql = "SELECT deal_id FROM " & sources.cm_key_data_table & " WHERE deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        If rs.BOF And rs.EOF Then
            conn.Execute "INSERT INTO " & sources.cm_key_data_table & " (deal_id) VALUES(" & deal_id & ")"
            load.rs_counter = load.rs_counter + 1
        End If
    rs.Close
    
    If global_vars.deal.is_init = False Then global_vars.deal.init deal_id
    
    str_sql = ""
    Dim input_data
    For counter = 1 To col_key_data.Count
        input_data = CallByName(global_vars.deal, col_key_data(counter), VbGet)
        
        'Check if data is date, as they require conversion to sql standard
        If Len(CStr(input_data)) - Len(Replace(CStr(input_data), "/", "")) = 2 Then
            'if input_data has two instances of / it must be a date. I think :)
            input_data = "'" & utilities.generate_sql_date(input_data) & "'"
        ElseIf Len(CStr(input_data)) - Len(Replace(CStr(input_data), "-", "")) = 2 And Len(CStr(input_data)) = 10 And Right(Left(input_data, 3), 1) = "-" Then
            'if input_data is 10 charcters, has two - and the third character is -, I will assume it is a date with the dd-mm-yyyy format, which needs to be converted to sql date format.
            input_data = "'" & utilities.generate_sql_date(input_data) & "'"
        ElseIf input_data = "-1" Then
            input_data = "NULL"
        Else
            input_data = "'" & input_data & "'"
        End If
        
        str_sql = str_sql & ", " & col_key_data(counter) & " = " & input_data
    Next counter
    str_sql = "UPDATE " & sources.cm_key_data_table & " SET is_deleted = 0" & str_sql & " WHERE deal_id = " & deal_id
    conn.Execute str_sql
    load.rs_counter = load.rs_counter + 1
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .change_source = proc_name
        .data_set = load.sources.cm_key_data_table
        .field_name = ""
        .executed_sql = str_sql
        .new_value = ""
        .record_id = deal_id
    End With
    log_object.data_logger log_object
    Set log_object = Nothing
    
outro:
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "central.add_key_data"
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Function create_policy_summary(ByVal deal_id _
, ByVal rs_policy_binders As ADODB.Recordset _
, Optional html As Boolean) As String

    Dim proc_name As String
    proc_name = "central.generate_policy_summary"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    'Purpose: Generate a summary of all policies on a deal
    Dim line_break
    Dim output As String
    
    If html = True Then
        line_break = "<br>"
    Else
        line_break = vbNewLine
    End If
    Dim bd_details As String
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim i As Integer
    Dim policy_item As Variant
    Dim rs_check As ADODB.Recordset
        
    For Each policy_item In global_vars.deal.col_policies
        bd_details = bd_details & policy_item.policy_no & line_break
        rs_policy_binders.Filter = "policy_id = '" & policy_item.policy_id & "'"
        'loop through all binders on policy
        With rs_policy_binders
            If .RecordCount > 0 Then
                .MoveFirst
                Do Until .EOF
                    bd_details = bd_details & "   " & !binder_name & ": " & Format(!binder_limit, "###,##0") & line_break
                    .MoveNext
                Loop
            End If
            bd_details = bd_details & line_break
        End With
    Next policy_item

    output = bd_details
    
outro:
    create_policy_summary = output
    Exit Function
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "central.generate_policy_summary"
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    output = "-1"
    GoTo outro
End Function

Public Sub init_deal(ByVal risk_object As cls_deal, ByVal first_run As Boolean)
    Dim proc_name As String
    proc_name = "central.init_deal"
    load.call_stack = load.call_stack & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim i As Integer
    Dim binder As cls_binder
    Dim deal_id As Long
    Dim policy As cls_policy
    Dim rs As ADODB.Recordset
    Dim str_sql As String
        
    deal_id = global_vars.deal.deal_id
    
    'add country procedures
    central.country_procedures_add_update deal_id, first_run
    
    'add binders to each policy on deal
    For Each policy In global_vars.deal.col_policies
        central.add_refresh_binder_list policy, first_run
    Next policy
    If global_vars.deal.binder_count = 0 Then
        GoTo outro
    End If
    global_vars.deal.has_currency_changed = False
    
    central.add_refresh_deal_questions deal_id
    
    'add inars and confirmations if not added already
    str_sql = "SELECT id FROM " & sources.deal_questions_view _
    & " WHERE question_type_id = " & load.question_types.confirmation _
    & " AND deal_id = " & deal_id
    
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        If rs.BOF And rs.EOF Then
            central.add_inar_questions_and_confirmations_to_deal deal_id
        End If
    rs.Close
    
    auto_referrals.start deal_id
    
    central.remove_redundant_referrals deal_id
    central.add_referrals_based_on_deal_questions deal_id
    
    'close and re-open binder list to reset form
    If CurrentProject.AllForms("binder_list_f").IsLoaded = True Then
        DoCmd.Close acForm, "binder_list_f"
    End If
    
outro:
    central.add_key_data deal_id
    
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    
    Exit Sub
    
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
End Sub
Public Sub add_referrals_based_on_deal_questions(ByVal deal_id As Long)
    'purpose: Take current set of questions and check if -binder- referrals should be triggered based on current answers
    Dim proc_name As String
    proc_name = "central.add_referrals_based_on_deal_questions"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String
    Dim binder_list_for_question()
    Dim binder_list_for_deal()
    Dim i As Integer
    Dim y As Integer
    Dim rs As ADODB.Recordset
    Dim rs_check As ADODB.Recordset
    
    ReDim target_referrals(1 To 1)
    
    'check for binder referrals
    str_sql = "SELECT * FROM " & load.sources.deal_questions_view _
    & " WHERE question_type_id = " & question_types.binder _
    & " AND deal_id = " & deal_id _
    & " AND NOT good_answer_id = answer_id"
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        If CLng(rs.RecordCount) = 0 Then
            GoTo outro
        End If
        rs.MoveFirst
        
        Do Until rs.EOF
            'check which binders are relevant for that question and which of those binders are on the deal. Puh!
            binder_list_for_question = utilities.get_binder_list_for_question(rs!template_question_id)
            binder_list_for_deal = utilities.get_binder_list_for_deal(deal_id)
            For i = 1 To UBound(binder_list_for_question)
                For y = 1 To UBound(binder_list_for_deal)
                    If binder_list_for_deal(y) = binder_list_for_question(i) Then
                        'referral triggered. Check if referral is already added
                        
                        str_sql = "SELECT * FROM " & sources.referrals_table _
                        & " WHERE is_deleted = 0" _
                        & " AND deal_question_id = " & rs!id _
                        & " AND deal_id = " & deal_id _
                        & " AND binder_id = " & binder_list_for_question(i)
                        
                        Set rs_check = utilities.create_adodb_rs(conn, str_sql)
                        rs_check.Open
                            If rs_check.RecordCount = 0 Then
                                str_sql = "INSERT INTO " & sources.referrals_table & " (deal_question_id, referral_status_id, binder_id, deal_id)" _
                                & " VALUES(" & rs.Fields("id").value _
                                & ", " & load.binder_referral_not_started _
                                & " , " & binder_list_for_question(i) _
                                & ", " & deal_id & ")"
                            
                                conn.Execute str_sql
                            End If
                        rs_check.Close
                    End If
                Next y
            Next i
            rs.MoveNext
        Loop
    rs.Close
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub update_referrals_for_new_response_to_deal_question(ByVal question_id, ByVal deal_id, ByVal rs_deal_questions As ADODB.Recordset)
    Dim proc_name As String
    proc_name = "central.update_referrals_for_new_response_to_deal_question"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    Dim str_sql As String, binder_list_for_question(), binder_list_for_deal(), target_referrals(), i As Integer, y As Integer, target_referral_count As Integer
    Dim rs_check As ADODB.Recordset, end_loop As Boolean, skip_to_next_item As Boolean
    Dim rs As ADODB.Recordset
    
    'check for binder referrals
    With rs_deal_questions
        .Filter = "id = " & question_id
        If CLng(.RecordCount) = 0 Then
            GoTo outro
        End If
        If !good_answer_id = !answer_id Then
            conn.Execute "UPDATE " & sources.referrals_table & " SET is_deleted = 1 WHERE deal_question_id = " & question_id & " AND deal_id = " & deal_id
            load.rs_counter = load.rs_counter + 1
            GoTo outro
        End If
        If !question_type_id = question_types.inar Then
            str_sql = "INSERT INTO " & sources.referrals_table & " (deal_question_id, referral_status_id, deal_id) VALUES(" & question_id & ", " & load.inar_referral_not_started & ", " & deal_id & ")"
            conn.Execute str_sql
            load.rs_counter = load.rs_counter + 1
            GoTo outro
        End If
        
        'check which binders are relevant for that question and which of those binders are on the deal.
        binder_list_for_question = utilities.get_binder_list_for_question(!template_question_id)
        binder_list_for_deal = utilities.get_binder_list_for_deal(deal_id)
        For i = 1 To UBound(binder_list_for_question)
            For y = 1 To UBound(binder_list_for_deal)
                If binder_list_for_deal(y) = binder_list_for_question(i) Then
                    
                    'check if referral is already added
                    str_sql = "SELECT * FROM " & load.sources.referrals_table _
                    & " WHERE is_deleted = 0" _
                    & " AND deal_question_id = " & question_id _
                    & " AND binder_id = " & binder_list_for_deal(y)
                    
                    Set rs = utilities.create_adodb_rs(conn, str_sql)
                    load.rs_counter = load.rs_counter + 1
                    With rs
                        .Open
                        
                        'if recordcount is 0, then referral is not added already
                        If CLng(.RecordCount) = 0 Then
                        
                            str_sql = "INSERT INTO " & sources.referrals_table & " (deal_question_id, referral_status_id, binder_id, deal_id) " _
                            & " VALUES(" & question_id & ", " & load.binder_referral_not_started _
                            & ", " & binder_list_for_question(i) & ", " & deal_id & ")"
                    
                            conn.Execute str_sql
                            load.rs_counter = load.rs_counter + 1
                    
                        End If
                        .Close
                    End With
                End If
            Next y
        Next i
        .Filter = ""
    End With
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "deal_id = " & deal_id & ", question_id = " & question_id & ", rs_deal_questions = [can't really be reproduced in a string]"
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub refresh_cm(ByVal deal_id As String)
    'purpose: [...]
    Dim proc_name As String
    proc_name = "central.refresh_cm"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim list_of_forms() As String, i As Long
    ReDim list_of_forms(1 To 50)
    i = 1
    list_of_forms(i) = "binder_list_f": i = i + 1
    list_of_forms(i) = "deal_questions_f": i = i + 1
    list_of_forms(i) = "deal_referrals_f": i = i + 1
    list_of_forms(i) = "extra_deal_limit_f": i = i + 1
    list_of_forms(i) = "help_f": i = i + 1
    list_of_forms(i) = "policies_f": i = i + 1
    list_of_forms(i) = "security_text_f": i = i + 1
    ReDim Preserve list_of_forms(1 To i - 1)
    Dim str_form As Variant
    For Each str_form In list_of_forms
        If CurrentProject.AllForms(str_form).IsLoaded = True Then
            DoCmd.Close acForm, str_form
        End If
    Next str_form
    central.load_cm_internal deal_id
    
outro:
    Exit Sub
    
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
    
End Sub
Public Function create_binder_referrals_array(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "central.create_binder_referrals_array"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim str_sql As String
    str_sql = "SELECT id, binder_id, question_type_id, referral_status_id, referral_status, due_time_for_referral FROM " & sources.referrals_view _
        & " WHERE deal_id = " & deal_id & " ORDER BY id"
    
    Dim array_info(0 To 1, 0 To 3)
    array_info(0, 0) = "number of items in array"
    array_info(0, 1) = "number of dimensions in array"
    array_info(0, 2) = "array headings"
    array_info(0, 3) = "array desc"
    
    array_info(1, 1) = 6
    
    Dim array_headings(0 To 5) As String
    array_headings(0) = "id"
    array_headings(1) = "question_type_id"
    array_headings(2) = "referral_status_id"
    array_headings(3) = "referral_status"
    array_headings(4) = "binder_id"
    array_headings(5) = "due_time_for_referral"
    
    array_info(1, 2) = array_headings
    
    array_info(1, 3) = "Array lists referrals for deal."
    
    Dim rs As ADODB.Recordset, output_array(), i As Long
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        
        ReDim output_array(0 To CLng(rs.RecordCount), 0 To array_info(1, 1) - 1)
        array_info(1, 0) = CLng(rs.RecordCount)
        output_array(0, 0) = array_info
    
        i = 1
        Do Until rs.EOF = True
            output_array(i, 0) = rs!id
            output_array(i, 1) = rs!question_type_id
            output_array(i, 2) = rs!referral_status_id
            output_array(i, 3) = rs!referral_status
            output_array(i, 4) = rs!binder_id
            output_array(i, 5) = rs!due_time_for_referral
            i = i + 1
            rs.MoveNext
        Loop
    rs.Close
    Set rs = Nothing
    
    create_binder_referrals_array = output_array
End Function
