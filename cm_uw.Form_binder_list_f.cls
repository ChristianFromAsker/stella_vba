VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_binder_list_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private binder_can_be_added As Boolean
Private Sub cmd_add_Click()
    Const proc_name As String = "binder_list_f.cmd_add.Click"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    If IsNull(Me!header_policy_id) Or IsEmpty(Me!header_policy_id) Then
        MsgBox "You must choose policy before you can add security to it.", , "No policy is chosen"
        Exit Sub
    End If
    
    Dim binder As cls_binder
    Dim binder_id As Long
    Dim current_binder As cls_binder
    Dim deal_id As Long
    Dim default_binder_currency As String
    Dim new_id As Long
    Dim policy As New cls_policy
    Dim policy_id As Long
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    deal_id = Me!header_deal_id
    policy_id = Me!header_policy_id
    
    If binder_can_be_added = False Then
        GoTo change_apperance
    End If
    If IsNull(Me!header_binder_id) Or IsEmpty(Me!header_binder_id) Then
        GoTo outro
    End If
    
    binder_id = Me!header_binder_id
        
    'check if binder is already on policy
    str_sql = "SELECT id FROM " & sources.policy_binders_view & " WHERE policy_id = " & policy_id & " AND binder_id = " & Me!header_binder_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        If rs.BOF And rs.EOF Then
            str_sql = "INSERT INTO " & load.sources.policy_binders_table _
            & " (policy_id, binder_id, on_policy_id, manually_added, added_by)" _
            & " VALUES(" & policy_id & ", " & binder_id & ", 93, 'yes', '" & Environ("username") & "')"
            
            conn.Execute str_sql, adExecuteNoRecords
            load.rs_counter = load.rs_counter + 1
            
            rs.Close
            
            Set binder = Nothing
            Set current_binder = Nothing
            For Each binder In global_vars.col_all_binders
                If binder.binder_id = binder_id Then
                    Set current_binder = binder
                End If
            Next binder
            
        Else
            MsgBox "Binder is already on the policy and cannot be added again", , "Binder already on policy"
            GoTo outro
        End If
    If rs.State = 1 Then
        rs.Close
    End If
    Set rs = Nothing
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .change_source = proc_name
        .changer_id = Environ("username")
        .data_set = load.sources.policy_binders_table
        .deal_id = deal_id
        .executed_sql = str_sql
        .field_name = ""
        .record_id = policy_id
    End With
    log_object.data_logger log_object
    Set log_object = Nothing
    load.rs_counter = load.rs_counter + 1

    
            
    '8 June 2023, CK: Consider implementing code to check if target_questions = actual questions. If yes, central.add_refresh_deal_questions deal_id and  fix_rs.deal_questions_f deal_id are not required.
    
    central.add_refresh_deal_questions deal_id
    fix_rs.deal_questions_f deal_id
    
    central.add_refresh_binder_limits current_binder
    
    Set policy = global_vars.deal.get_policy(policy_id)
    policy.col_binders.Add current_binder
    policy.send_binder_data_to_mysql
    policy.init_binder_data
    central.limit_distribution policy
    fix_rs.binder_list_f
    auto_referrals.start deal_id
    central.check_status_area_on_control_panel_f deal_id
    
    'reset form
change_apperance:
    With Me
        If binder_can_be_added = True Then
            !cmd_add.Top = 150
            !header_binder_id = Empty
            !cmd_add.Caption = "I need to add security"
            !header_binder_id.Visible = False
            !lbl_add_binder.Visible = False
            binder_can_be_added = False
        Else
            With !header_binder_id
                .Visible = True
                Do While .ListCount > 0
                    .RemoveItem (0)
                Loop
                str_sql = "SELECT binder_id id, binder_name menu_item FROM " & load.sources.binder_list_table & " b WHERE b.is_deleted = 0 AND is_active = 93 ORDER BY b.binder_name"
                Set rs = utilities.create_adodb_rs(conn, str_sql)
                rs.Open
                    Do While rs.EOF = False
                        .AddItem rs!id & ";" & rs!menu_item
                        rs.MoveNext
                    Loop
                rs.Close
                Set rs = Nothing
            End With
            !lbl_add_binder.Visible = True
            !cmd_add.Caption = "add"
            !cmd_add.Top = 1000
            binder_can_be_added = True
        End If
    End With
    
outro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    load.call_stack = ""
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "deal_id = " & deal_id, "", "", True
    Resume outro
End Sub

Private Sub cmd_policy_text_Click()
    'intro
    On Error GoTo err_handler
    If IsNull(Me!header_policy_id) = True Or Me!header_policy_id = "" Then
        MsgBox "You must choose a policy before you can see the security allocation.", , "Please choose policy"
        GoTo outro
    End If
    If is_debugging = True Then On Error GoTo 0
    load.check_conn_and_variables
    open_forms.security_text_f Me!header_policy_id_admin

outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "binder_list_f.cmd_policy_text.Click"
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Private Sub cmd_change_Click()
    Const proc_name As String = "binder_list_f.cmd_change"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    With Me
        !lbl_binder_on_policy.Visible = True
        !lbl_limit_adjustment.Visible = True
        !header_on_policy = Me!on_policy_id
        !header_on_policy.Visible = True
        !header_manual_limit.Visible = True
        !header_manual_limit = Me!manual_limit
        !header_id = Me!id
        !cmd_update.Visible = True
        !header_policy_id.Locked = True
        !header_policy_id.BackColor = colors.light_grey
        !header_binder_id_admin = Me!binder_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "binder_list_f.cmd_change.Click"
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Private Sub cmd_security_export_Click()
 Dim proc_name As String
    proc_name = "binder_list_f.cmd_security_export.click"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    load.check_conn_and_variables
    
    open_forms.working_on_it_f "Preparing the security export"
    
    Dim deal_id As Long
    
    deal_id = Forms("binder_list_f").Controls("header_deal_id")
    
    load.check_secondary_access_app
    With load.secondary_access_app
        On Error Resume Next
        .CloseCurrentDatabase
        On Error GoTo err_handler
        If load.is_debugging = True Then On Error GoTo 0

        .OpenCurrentDatabase load.system_info.system_paths.common_path & "reports.accdb", False
        .Visible = False
        .Run "reports.excel_extract", deal_id
        .CloseCurrentDatabase
        .OpenCurrentDatabase load.system_info.system_paths.common_path & "placeholder.accdb", False
    End With
    
outro:
    load.call_stack = ""
    open_forms.working_on_it_f__close
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

Private Sub cmd_update_Click()
    Const proc_name As String = "binder_list_f.cmd_update"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    Dim binder As cls_binder
    Dim binder_id As Long
    Dim deal_id As Long
    Dim manual_limit_text As String
    Dim policy As cls_policy
    Dim policy_id As Long
    Dim quota_text As String
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    'checking mandatory fields
    If Me!header_id = "" Or IsNull(Me!header_id) Then
        MsgBox "Please click on 'change' on relevant binder to change it.", , "no binder selected"
        Exit Sub
    End If
    
    binder_id = Me!header_binder_id_admin
    deal_id = Me!header_deal_id
    policy_id = Me!header_policy_id
    
    If Me!header_on_policy <> menu_list.yes And Me!header_on_policy <> menu_list.no Then
        MsgBox "You must state whether the binder is on the policy.", , "missing info"
        GoTo outro
    End If
    
    If global_vars.deal Is Nothing Then
        Set global_vars.deal = New cls_deal
    End If
    With global_vars.deal
        If .is_init = False Then
            .init_deal_data deal_id
        End If
        If .col_policies.Count = 0 Then
            .init_policy_data deal_id
        End If
        Set policy = .get_policy(policy_id)
        If policy.col_binders.Count = 0 Then
            policy.init_binder_data
        End If
    End With
    
    Set binder = policy.get_binder(binder_id)
    
    'update record
    If Me!header_on_policy = menu_list.no Then
        quota_text = ", quota = 0"
        manual_limit_text = ", manual_limit = 0"
        
        binder.binder_quota = 0
        binder.manual_limit = 0
        binder.on_policy_id = menu_list.no
    Else
        quota_text = "quota = " & Me!header_manual_limit / policy.policy_limit
        'some jurisdictions use , as decimal delimiter. Need to cater for that.
        quota_text = Replace(quota_text, ",", ".")
        quota_text = ", " & quota_text
        manual_limit_text = ", manual_limit = " & Me!header_manual_limit
        
        binder.manual_limit = Me!header_manual_limit
        binder.binder_quota = Me!header_manual_limit / policy.policy_limit
        binder.on_policy_id = menu_list.yes
    End If
    
    str_sql = "UPDATE " & load.sources.policy_binders_table & " SET" _
    & " on_policy_id = " & Me!header_on_policy _
    & manual_limit_text _
    & quota_text _
    & " WHERE id = " & Me!header_id
    
    conn.Execute str_sql
    load.rs_counter = load.rs_counter + 1
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.policy_binders_table
        .field_name = ""
        .executed_sql = str_sql
        .record_id = Me!header_id
    End With
    log_object.data_logger log_object
    Set log_object = Nothing
    load.rs_counter = load.rs_counter + 1
            
    'reset input fields
    With Me
        !header_on_policy = ""
        !header_manual_limit = 0
        !header_id = Empty
        !lbl_binder_on_policy.Visible = False
        !lbl_limit_adjustment.Visible = False
        !header_on_policy.Visible = False
        !header_manual_limit.Visible = False
        !header_policy_id.SetFocus
        !cmd_update.Visible = False
        !header_policy_id.Locked = False
        !header_policy_id.BackColor = colors.white
    End With
    central.limit_distribution policy
    fix_rs.binder_list_f
    central.check_status_area_on_control_panel_f deal_id
    
    'add data to allocated_limit field
    Dim total_limit As Currency
    total_limit = 0
    For Each binder In policy.col_binders
        If binder.on_policy_id = menu_list.yes Then
            total_limit = total_limit + binder.get_binder_limit
        End If
    Next binder
    Me!header_allocated_limit = total_limit
    
outro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    load.call_stack = ""
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "binder_list_f.cmd_update.Click"
        .milestone = ""
        .params = "str_sql = " & ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Private Sub Form_Load()
    binder_can_be_added = False
End Sub
Private Sub header_policy_id_AfterUpdate()
    Const proc_name As String = "binder_list_f.header_policy_id.AfterUpdate"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    Dim policy_no As Long
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "SELECT id FROM " & load.sources.policy_binders_view & " WHERE deal_id = " & Me!header_deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        If rs.BOF And rs.EOF Then
            GoTo outro
        End If
    rs.Close
    
    'fix recordsource
    fix_rs.binder_list_f
    
    'add data to allocated_limit field
    str_sql = "SELECT SUM(binder_limit) allocated_limit FROM " & load.sources.policy_binders_view & " WHERE policy_id = " & Me.header_policy_id & " AND on_policy_id = 93"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        Me!header_allocated_limit = rs!allocated_limit
    rs.Close
    
    policy_no = Me!header_policy_id.Column(0)
    With Me
        !header_policy_id.BackColor = colors.white
        !header_policy_limit.Visible = True
    End With
    
outro:
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    load.call_stack = ""
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "str_sql = " & str_sql, "", "", True
    Resume outro
End Sub
