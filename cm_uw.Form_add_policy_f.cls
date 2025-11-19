VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_add_policy_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub attach_AfterUpdate()
    Me!attach.BackColor = colors.white
End Sub

Private Sub broker_commission_AfterUpdate()
    Me.broker_commission.BackColor = load.colors.white
End Sub

Private Sub cmd_save_close_Click()
    Const proc_name As String = "add_policy_f.cmd_save_close.Click"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.event_id = utilities.get_event_id
    load.check_conn_and_variables
    open_forms.working_on_it_f "Making updates based on changes to the policy. Hold on please."
    
    Dim arr_fields() As Variant
    Dim check_failed As Boolean
    Dim deal_id As Long
    Dim i As Integer
    Dim log_object As cls_log_object
    Dim policy_id As Long
    Dim rs As ADODB.Recordset
    Dim shall_close_when_done As Boolean
    Dim str_milestone As String
    Dim str_sql As String
    
    deal_id = Me!deal_id
    shall_close_when_done = True
    
    'load policy object
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
    End With
    
    'check if this is the only policy on deal, and the UW is trying to take RP off and this is not a new policy
    If global_vars.deal.col_policies.Count = 0 And Me!rp_on_layer <> 93 Then
        If IsNull(Me!id) = False Then
            MsgBox "There are no policies where RP is on. Please add a policy which RP is on before taking RP off this one.", , "No policies with RP"
            GoTo outro
        End If
    End If
    
    'Check if mandatory fields are populated
    str_milestone = "Before mandatory fields check"
    arr_fields = declare_fields
    
    check_failed = False
    For i = 1 To arr_fields(0, 0)
        If (IsNull(Me.Controls(arr_fields(i, 0))) = True Or IsEmpty(Me.Controls(arr_fields(i, 0))) Or Me.Controls(arr_fields(i, 0)) = "") _
            And arr_fields(i, 1) = True Then
            Me.Controls(arr_fields(i, 0)).BackColor = colors.yellow
            check_failed = True
        End If
    Next i
    
    'check if underlying limit is properly populated
    str_milestone = "Before underlying limit check"
    If Me!layer_no = 0 And Me!underlying_limit <> 0 Then
        MsgBox "Underlying limit must be 0 for primary policies.", , "primary policies don't have underlying limits"
        check_failed = True
    End If
    If Me!layer_no > 0 And Me!underlying_limit = 0 Then
        MsgBox "All xs policies must have an underlying limit. Please add.", , "underlying limit missing"
        check_failed = True
    End If
    
    'check if either both fundamental limit and fundamental premium is populated, or none.
    If (Me!fundamental_limit = 0 And Me!fundamental_premium <> 0) Or _
    (Me!fundamental_limit <> 0 And Me!fundamental_premium = 0) Then
        MsgBox "A policy cannot have Fundamental Premium without Fundamental Limit and vice versa.", , "Issue with Fundamental Premium/Limit"
        check_failed = True
    End If
    
    If check_failed = True Then
        shall_close_when_done = False
        GoTo outro
    End If
    
    Dim current_attachment As Currency
    Dim current_policy_id_with_lowest_attachment As Long
    Dim policy_new As cls_policy
    Dim policy As cls_policy
    Dim str_message As String
    Dim str_quota As String
    
    str_quota = CStr(Me!quota)
    str_quota = Replace(str_quota, ",", ".")
    
    'Insert data into table
    'first, get id. If new, create new record
    If (IsNull(Me!id) = True Or Me!id = "") Then
    
        'check if there is already a primary layer
        If Me!layer_no = 0 Then
            For Each policy In global_vars.deal.col_policies
                With policy
                    If policy.layer_no = 0 Then
                        str_message = "There is already a primary policy on this deal. Are you sure you want to add another one?"
                        If MsgBox(str_message, vbQuestion + vbYesNo, "Confirm additional primary policy") = vbNo Then
                            GoTo outro
                        End If
                    End If
                End With
            Next policy
        End If
        
        'insert new policy into policy table. Then get newly created id
        str_milestone = "Before adding adding new policy"
        str_sql = "INSERT INTO " & load.sources.policies_table & "(deal_id) VALUES (" & Me!deal_id & ")"
        conn.Execute str_sql, adExecuteNoRecords
        load.rs_counter = load.rs_counter + 1
        
        str_sql = "SELECT id policy_id FROM " & load.sources.policies_table & " WHERE deal_id = " & Me!deal_id & " ORDER BY id DESC"
        Set rs = utilities.create_adodb_rs(conn, str_sql)
        rs.Open
            policy_id = rs!policy_id
        rs.Close
        Set rs = Nothing
        
        Set policy_new = New cls_policy
        With policy_new
            .deal_id = deal_id
            .policy_id = policy_id
        End With
        global_vars.deal.col_policies.Add policy_new
        
        'log change
        Set log_object = New cls_log_object
        With log_object
            .changer_id = Environ("username")
            .data_set = load.sources.policies_table
            .field_name = ""
            .executed_sql = "a new policy with id " & policy_id & " was added to deal_id " & deal_id
            .record_id = policy_id
        End With
        log_object.data_logger log_object
        Set log_object = Nothing
    Else
        policy_id = Me!id
        Set policy_new = global_vars.deal.get_policy(policy_id)
        
        If policy_new.col_binders.Count = 0 Then
            policy_new.init_binder_data
        End If
    End If
    
    'Update data
    Dim policy_name As String
    Dim stella_policy_no As String
    
    policy_name = ""
    If IsNull(Me!policy_name) = False Then
        policy_name = ", policy_name = '" & Replace(Me!policy_name, "'", "''") & "'"
    End If
    
    stella_policy_no = CStr(policy_id)
    If Len(stella_policy_no) <= 7 Then
        Do Until Len(stella_policy_no) = 8
            stella_policy_no = "0" & stella_policy_no
        Loop
    End If
    stella_policy_no = "RPMA" & stella_policy_no
    
    str_milestone = "Before updating data on policy"
    str_sql = "UPDATE " & load.sources.policies_table & " SET " _
        & "layer_no = " & Me!layer_no _
        & ", underlying_limit = " & Me!underlying_limit _
        & ", quota = " & str_quota _
        & ", layer_limit = " & Replace(Me!layer_limit, ",", ".") _
        & ", fundamental_limit = " & Replace(Me!fundamental_limit, ",", ".") _
        & ", rp_on_layer = " & CLng(Me!rp_on_layer) _
        & ", policy_premium = " & Replace(Me!policy_premium, ",", ".") _
        & ", fundamental_premium = " & Replace(Me!fundamental_premium, ",", ".") _
        & ", broker_commission = " & Replace(Me!broker_commission, ",", ".") _
        & ", budget_home_id = " & Me!budget_home_id _
        & ", stella_policy_no = '" & stella_policy_no & "'" _
        & ", issuing_entity_id = " & Me!issuing_entity_id _
        & policy_name _
        & " WHERE id = " & policy_id
    conn.Execute str_sql
    load.rs_counter = load.rs_counter + 1
    
    'refresh local policy object
    With policy_new
        .broker_commission = Me!broker_commission
        .budget_home_id = Me!budget_home_id
        .fundamental_limit = Me!fundamental_limit
        .fundamental_premium = Me!fundamental_premium
        .issuing_entity_id = Me!issuing_entity_id
        .layer_limit = Me!layer_limit
        .layer_no = Me!layer_no
        .policy_name = Nz(Me!policy_name, "-1")
        .policy_premium = Me!policy_premium
        .quota = Me!quota
        .rp_on_layer_id = CLng(Me!rp_on_layer)
        .stella_policy_no = stella_policy_no
        .underlying_limit = Me!underlying_limit
        If .col_binders Is Nothing Then
            Set .col_binders = New Collection
        End If
    End With
    
    'log change
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.policies_table
        .deal_id = deal_id
        .executed_sql = str_sql
        .record_id = policy_id
    End With
    log_object.data_logger log_object
    
    If CurrentProject.AllForms("policies_f").IsLoaded = True Then
        fix_rs.policies_f deal_id
    End If
    
    str_milestone = "Before running all the central modules"
    central.add_refresh_binder_list policy_new, False
    central.add_refresh_deal_questions deal_id
    central.limit_premium_and_summary_central deal_id
    central.check_status_area_on_control_panel_f deal_id
    
    'check if new data relevant for iManage changed. If yes, update iManage (done via create_workspace call).
    'iManage must be updated if navins home or lowest underlying limit changse
'    new_navins_home = -1
'    current_attachment = -1
'    For Each policy In global_vars.deal.col_policies
'        With policy
'            If current_attachment = -1 Then
'                current_policy_id_with_lowest_attachment = policy.policy_id
'            ElseIf current_attachment > policy.underlying_limit Then
'                current_policy_id_with_lowest_attachment = policy.policy_id
'            End If
'        End With
'    Next policy
    
    str_milestone = "Before iManage"
    imanage.update_workspace_caller deal_id, load.current_uw.email
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    
    open_forms.working_on_it_f__close
    If shall_close_when_done = True Then
        DoCmd.Close acForm, Me.Name
    End If
    load.call_stack = ""
    
    Exit Sub
    
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "str_sql = " & str_sql, "deal_id = " & deal_id _
    & vbNewLine & ", policy_id = " & policy_id, "", True
    Resume outro
End Sub
Private Function declare_fields() As Variant
    Const int_number As Integer = 12
    Dim input_controls(0 To int_number, 0 To 2) As Variant
    Dim i As Integer: i = 1
    ' i,0 is name of controls
    ' i,1 is whether fields is mandatory upon saving.
    input_controls(0, 0) = int_number
    input_controls(i, 0) = "deal_name"
    input_controls(i, 1) = False: i = i + 1
    input_controls(i, 0) = "deal_id"
    input_controls(i, 1) = False: i = i + 1
    input_controls(i, 0) = "layer_no"
    input_controls(i, 1) = True: i = i + 1
    input_controls(i, 0) = "quota"
    input_controls(i, 1) = True: i = i + 1
    input_controls(i, 0) = "underlying_limit"
    input_controls(i, 1) = True: i = i + 1
    input_controls(i, 0) = "layer_limit"
    input_controls(i, 1) = True: i = i + 1
    input_controls(i, 0) = "fundamental_limit"
    input_controls(i, 1) = True: i = i + 1
    input_controls(i, 0) = "rp_on_layer"
    input_controls(i, 1) = True: i = i + 1
    input_controls(i, 0) = "policy_premium"
    input_controls(i, 1) = True: i = i + 1
    input_controls(i, 0) = "fundamental_premium"
    input_controls(i, 1) = True: i = i + 1
    input_controls(i, 0) = "broker_commission"
    input_controls(i, 1) = True: i = i + 1
    input_controls(i, 0) = "issuing_entity_id"
    input_controls(i, 1) = True: i = i + 1
    declare_fields = input_controls
End Function

Private Sub fundamental_limit_AfterUpdate()
    load.check_conn_and_variables
    load.add_policy.update_policy_breakdown
    Me.fundamental_limit.BackColor = load.colors.white
End Sub

Private Sub fundamental_premium_AfterUpdate()
    load.check_conn_and_variables
    load.add_policy.update_policy_breakdown
    Me.fundamental_premium.BackColor = load.colors.white
End Sub

Private Sub layer_limit_AfterUpdate()
    load.check_conn_and_variables
    load.add_policy.update_policy_breakdown
    Me.layer_limit.BackColor = load.colors.white
End Sub
Private Sub layer_no_AfterUpdate()
    load.check_conn_and_variables
    Me!layer_no.BackColor = colors.white
End Sub

Private Sub policy_name_info_Click()
    MsgBox "Policy name is optional, and used for when you need to distinguish policies from one another. For example, if you have different policies for different capacities.", , "Info about policy names"
End Sub

Private Sub policy_premium_AfterUpdate()
    load.check_conn_and_variables
    load.add_policy.update_policy_breakdown
    Me.policy_premium.BackColor = load.colors.white
End Sub

Private Sub quota_AfterUpdate()
    load.check_conn_and_variables
    load.check_conn_and_variables
    load.add_policy.update_policy_breakdown
    Me!quota.BackColor = colors.white
End Sub
Private Sub rp_on_layer_AfterUpdate()
    load.check_conn_and_variables
    With Me!rp_on_layer
        .BackColor = colors.white
        If .value = 94 Then
            Me!lbl_message.Caption = "No need to add policies RP is not on :)"
            Me!lbl_message.BackColor = load.colors.yellow
        Else
            Me!lbl_message.Caption = ""
            Me!lbl_message.BackColor = load.colors.white
        End If
    End With
End Sub

Private Sub underlying_limit_AfterUpdate()
    Me.underlying_limit.BackColor = load.colors.white
End Sub
