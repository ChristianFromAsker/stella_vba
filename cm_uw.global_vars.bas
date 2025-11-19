Attribute VB_Name = "global_vars"
Option Compare Database
Option Explicit

Public Enum binder_data
    
    'binder
    security_id = 1
    binder_id = 2
    binder_name = 20
    binder_limit = 14
    limit_currency = 43
    default_binder_limit = 15
    default_binder_currency = 30
    default_quota = 7
    end_date = 31
    extra_limit = 9
    for_eur_id = 32
    for_us_id = 33
    insurer_business_name = 29
    insurer_id = 34
    insurer_legal_name = 27
    is_active_id = 40
    manual_limit = 10
    max_binder_limit = 19
    max_binder_limit_currency = 28
    max_limit_currency = 16
    max_limit_deal_currency = 17
    max_limit_id = 4
    minimum_underlying_limit = 41
    on_policy_id = 44
    quota = 8
    reference_limit = 13
    reference_limit_deal_currency = 18
    referral_status = 23
    referral_status_id = 6
    start_date = 42
    unique_reference = 21
    
    'deal
    currency_rate_deal = 104
    currency_rate_local = 105
    deal_currency = 101
    deal_id = 100
    insured_id = 102
    layer_quota = 103
    
    'policy
    layer_no = 204
    policy_id = 200
    policy_limit = 201
    on_policy = 202
    
End Enum

Public Enum policy_data
    policy_id = 1
    binders = 500
    budget_home = 2
    budget_home_id = 3
    display_view = 4
    inception_date = 5
    issuing_entity = 6
    issuing_entity_id = 7
    layer_no = 8
    layer_no_text = 9
    policy_limit = 10
    policy_name = 11
    policy_no = 12
    policy_premium = 13
    quota = 14
    underlying_limit = 15
End Enum

Public Enum control_types
    control_button = 3
    lbl = 1
    txt_box = 2
End Enum

Public binder_questions As New Collection
Public deal_referrals As New cls_frm_deal_referrals
Public col_all_binders As New Collection
Public col_binder_limits As New Collection
Public col_currencies As New Collection
Public col_policies As Collection
Public col_policy_binders As New Collection

Private Type typ_referral_status
    cleared_internally As Long
    approved As Long
    no_referral As Long
    in_process As Long
    denied As Long
    not_started As Long
End Type
Public referral_statuses As typ_referral_status

Public binder_policy As Scripting.Dictionary
Public deal As cls_deal

Public Sub init()
    Const proc_name As String = "global_vars.init"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim deal_id As Long
    Dim policy As cls_policy
    
    global_vars.col_all_binders_init
    global_vars.col_binder_limits_init
    global_vars.col_currencies_init
    global_vars.deal_referrals.init
    global_vars.init_binder_questions
     
    With global_vars.referral_statuses
        .approved = 487
        .cleared_internally = 486
        .denied = 490
        .in_process = 489
        .no_referral = 488
        .not_started = 521
    End With
    
    If CurrentProject.AllForms(load.control_panel.form_name).IsLoaded = True Then
        deal_id = Forms(load.control_panel.form_name)!deal_id
        If global_vars.deal Is Nothing Then
            Set global_vars.deal = New cls_deal
        End If
        If global_vars.deal.is_init = False Then
            With global_vars.deal
                .init_deal_data deal_id
                .init_policy_data deal_id
                For Each policy In .col_policies
                    policy.init_binder_data
                Next policy
            End With
        End If
    End If
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub col_currencies_init()
    Dim proc_name As String
    proc_name = "global_vars.init_currencies"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim field As ADODB.field
    Dim input_value As Variant
    Dim input_name As Variant
    Dim fx_currency As Scripting.Dictionary
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    Set global_vars.col_currencies = New Collection
    str_sql = "SELECT * FROM " & load.sources.currencies_table & " ORDER BY currency_date DESC"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        Do Until .EOF
            Set fx_currency = New Scripting.Dictionary
            For Each field In rs.Fields
                If IsNull(field.Name) = False And IsNull(field.value) = False Then
                    If field.Name = "currency_date" Then
                        fx_currency.Add field.Name, utilities.generate_sql_date(field.value)
                    Else
                        fx_currency.Add field.Name, field.value
                    End If
                End If
            Next field
            global_vars.col_currencies.Add fx_currency
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
outro:
    Exit Sub
    
err_handler:
    GoTo outro
End Sub
Public Sub col_binder_limits_init()
    Dim proc_name As String
    proc_name = "load.col_binder_limits_init"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim binder_limit As Scripting.Dictionary
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "SELECT id binder_limit_id, binder_id, limit_currency, limit_amount, is_active " _
    & " FROM " & load.sources.binder_limits_table _
    & " ORDER BY binder_id"
        
    Set global_vars.col_binder_limits = Nothing
    Set global_vars.col_binder_limits = New Collection
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        Do Until rs.EOF
            Set binder_limit = New Scripting.Dictionary
            With binder_limit
                .Add "binder_limit_id", rs!binder_limit_id.value
                .Add binder_data.binder_id, rs!binder_id.value
                .Add binder_data.limit_currency, rs!limit_currency.value
                .Add binder_data.max_binder_limit, rs!limit_amount.value
                .Add binder_data.max_binder_limit_currency, rs!limit_currency.value
            End With
            global_vars.col_binder_limits.Add binder_limit
            rs.MoveNext
        Loop
    rs.Close
    
outro:
    If Not rs Is Nothing Then Set rs = Nothing
    Set rs = Nothing
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "load.init_template_questions"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    Resume outro
End Sub
Public Sub col_all_binders_init()
    Dim proc_name As String
    proc_name = "load.col_all_binders_init"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim binder As cls_binder
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "SELECT *" _
    & " FROM " & load.sources.binder_list_view _
    & " WHERE is_active_id = " & menu_list.yes _
    & " ORDER BY binder_name"
        
    Set global_vars.col_all_binders = New Collection
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        Do Until rs.EOF
            Set binder = New cls_binder
            With binder
                .binder_id = rs!binder_id.value
                .binder_name = rs!binder_name.value
                .default_binder_currency = rs!default_currency.value
                .end_date = rs!end_date.value
                .for_eur_id = rs!for_eur_id.value
                .for_us_id = rs!for_us_id.value
                .insurer_business_name = rs!insurer_business_name.value
                .insurer_id = rs!insurer_id.value
                .is_active_id = rs!is_active_id.value
                .manual_limit = 0
                .minimum_underlying_limit = rs!minimum_underlying_limit.value
                .reference_limit = rs!reference_limit.value
                .unique_reference = rs!unique_reference.value
            End With
            global_vars.col_all_binders.Add binder
            rs.MoveNext
        Loop
    rs.Close
    
outro:
    If Not rs Is Nothing Then Set rs = Nothing
    Set rs = Nothing
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub init_binder_questions()
    Const proc_name As String = "global_vars.init_binder_questions"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim question_info As Scripting.Dictionary
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    Set global_vars.binder_questions = Nothing
    
    str_sql = "SELECT id, question, question_category, binder_id, binder_name, template_question_id, question_type_id" _
    & " FROM " & load.sources.binder_questions_view
    
    Set rs = utilities.create_adodb_rs(load.conn, str_sql)
    With rs
        .Open
        .MoveFirst
        Do Until .EOF
            Set question_info = New Scripting.Dictionary
            With question_info
                .Add "binder_id", rs!binder_id.value
                .Add "binder_name", rs!binder_name.value
                .Add "template_question_id", rs!template_question_id.value
                .Add "question", rs!question.value
                .Add "question_type_id", rs!question_type_id.value
            End With
            global_vars.binder_questions.Add question_info
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub

