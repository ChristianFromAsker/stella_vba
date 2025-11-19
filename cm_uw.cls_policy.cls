VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_policy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'data points
Public broker_commission As Variant
Public budget_home As String
Public budget_home_id As Long

Public col_binders As Collection
Public create_date

Public deal_id As Long
Public display_view As String

Public fundamental_limit As Currency
Public fundamental_premium As Currency

Public inception_date As Variant 'because it might be "-1" if not provided
Public issuing_entity As String
Public issuing_entity_id As Long
Public layer_limit As Currency
Public layer_no As Long
Public layer_no_text As String

Public policy_id As Long
Public policy_name As String
Public policy_no As String
Public policy_premium As Currency
Public quota As Variant
Public rp_on_layer_id As Long

Public stella_policy_no As String

Public underlying_limit As Currency

Public Sub update_binder_quota(ByVal input_quota, ByVal input_binder_id As Long)
    Dim proc_name As String
    proc_name = "cls_policy.update_binder_quota"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim binder As cls_binder
    For Each binder In Me.col_binders
        If binder.binder_id = input_binder_id Then
            binder.binder_quota = input_quota
        End If
    Next binder
End Sub
Public Sub update_default_quota(ByVal input_quota, ByVal input_binder_id As Long)
    Dim proc_name As String
    proc_name = "cls_policy.update_default_quota"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim binder As cls_binder
    For Each binder In Me.col_binders
        If binder.binder_id = input_binder_id Then
            binder.default_quota = input_quota
        End If
    Next binder
End Sub
Public Sub send_binder_data_to_mysql()
    Dim proc_name As String
    proc_name = "cls_policy.send_binder_data_to_mysql"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim binder As cls_binder
    Dim fx_rate
    Dim log_object As utilities.typ_log_object
    Dim str_quota As String
    Dim str_sql As String
    Dim tx_started As Boolean
    
    str_sql = ""
    For Each binder In Me.col_binders
        
        fx_rate = Replace(CStr(binder.fx_deal_binder_rate), ",", ".")
        
        str_quota = ""
        If IsEmpty(binder.binder_quota) = False Then
            str_quota = ", quota = " & Replace(binder.binder_quota, ",", ".") _
            & ", default_quota = " & Replace(binder.default_quota, ",", ".")
        End If
        
        str_sql = "UPDATE " & load.sources.policy_binders_table & " SET" _
            & " manual_limit = " & binder.manual_limit _
            & str_quota _
            & ", max_limit_id = " & binder.max_limit_id _
            & ", max_limit_currency_to_deal_currency_fx = " & fx_rate _
        & " WHERE" _
            & " binder_id = " & binder.binder_id _
            & " AND policy_id = " & Me.policy_id
        
        conn.Execute str_sql
        load.rs_counter = load.rs_counter + 1
        
        'log change
        With log_object
            .change_source = proc_name
            .changer_id = Environ("username")
            .data_set = load.sources.policy_binders_table
            .deal_id = global_vars.deal.deal_id
            .event_id = load.event_id
            .executed_sql = str_sql
            .operation_type = "update"
            .policy_id = Me.policy_id
            .record_id = binder.security_id
            .security_id = binder.security_id
        End With
        
        load.conn.BeginTrans
        tx_started = True
            utilities.log_change__field_change log_object, "manual_limit", binder.manual_limit, load.sources.policy_binders_table
            If IsEmpty(binder.binder_quota) = False Then
                utilities.log_change__field_change log_object, "default_quota", Replace(binder.default_quota, ",", "."), load.sources.policy_binders_table
                utilities.log_change__field_change log_object, "quota", Replace(binder.binder_quota, ",", "."), load.sources.policy_binders_table
            End If
            utilities.log_change__field_change log_object, "max_limit_id", binder.max_limit_id, load.sources.policy_binders_table
            utilities.log_change__field_change log_object, "max_limit_currency_to_deal_currency_fx", fx_rate, load.sources.policy_binders_table
        load.conn.CommitTrans
        tx_started = False

    Next binder
    
outro:
    Exit Sub
    
err_handler:
    If tx_started = True And Not load.conn Is Nothing Then
        If load.conn.State = adStateOpen Then
            load.conn.RollbackTrans
        End If
    End If
    
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
Public Function use_reference_limit() As Boolean
    Dim proc_name As String
    proc_name = "cls_policy.use_reference_limit"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim binder As cls_binder
    Dim binder_quota As Double
    Dim output As Boolean
    output = True
    
    If Me.get_total_reference_limit_without_binders_with_manual_limit < 0 Then
        output = False
    Else
        For Each binder In Me.col_binders
            If binder.on_policy_id = menu_list.yes Then
                binder_quota = binder.reference_limit / Me.get_total_reference_limit_without_binders_with_manual_limit
                If binder_quota * Me.policy_limit > binder.max_limit_deal_currency Then
                    output = False
                End If
            End If
        Next binder
    End If
    
    use_reference_limit = output
outro:
    Exit Function
err_handler:
    GoTo outro
End Function
Public Sub binder_allocation_unrounded()
    Dim proc_name As String
    proc_name = "cls_policy.binder_allocation_unrounded"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim binder As cls_binder
    Dim limit_for_default_allocation As Currency
    Dim quota
    Dim str_sql As String
    
    limit_for_default_allocation = Me.get_limit_for_default_allocation
    
    For Each binder In Me.col_binders
    If binder.on_policy_id = menu_list.yes Then
        If binder.manual_limit = 0 Then
            quota = binder.max_limit_deal_currency / Me.get_max_limit_for_default_allocation_deal_currency * Me.get_quota_for_default_alloation
            If quota * Me.policy_limit > binder.max_limit_deal_currency Then
                quota = binder.max_limit_deal_currency / Me.policy_limit
            End If
        Else
            quota = binder.manual_limit / Me.policy_limit
        End If
        quota = Replace(CStr(quota), ",", ".")
        
        str_sql = "UPDATE " & sources.policy_binders_table & " SET" _
        & " default_quota = " & quota _
        & ", quota = " & quota _
        & " WHERE id = " & binder.security_id
        
        conn.Execute str_sql
        load.rs_counter = load.rs_counter + 1
        
        binder.binder_quota = quota
        binder.default_quota = quota
    End If
    Next binder

outro:
    Exit Sub
    
err_handler:
    GoTo outro
    
End Sub
Public Function get_binder_with_highest_limit() As cls_binder
    Dim proc_name As String
    proc_name = "cls_policy.binder_count_without_manual_limit"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim binder As cls_binder
    Dim highest_limit As Currency
    Dim output As cls_binder
    
    highest_limit = 0
    For Each binder In Me.col_binders
        If binder.on_policy_id = menu_list.yes Then
            If binder.max_limit_deal_currency > highest_limit And binder.manual_limit = 0 Then
                highest_limit = binder.max_limit_deal_currency
                Set output = binder
            End If
        End If
    Next binder
    
    Set get_binder_with_highest_limit = output
    
outro:
    Exit Function
err_handler:
    GoTo outro
End Function
Public Function binder_count_without_manual_limit() As Long
    Dim proc_name As String
    proc_name = "cls_policy.binder_count_without_manual_limit"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim binder As cls_binder
    Dim binder_count As Long
    
    binder_count = 0
    For Each binder In Me.col_binders
        If binder.manual_limit = 0 Then
            binder_count = binder_count + 1
        End If
    Next binder
    
    binder_count_without_manual_limit = binder_count
outro:
    Exit Function
    
err_handler:
    GoTo outro
End Function
Public Function is_binder_on_policy(ByVal binder_id As Long) As Boolean
    Dim proc_name As String
    proc_name = "cls_policy.is_binder_on_policy"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim binder As cls_binder
    Dim output As Boolean
    
    output = False
    
    For Each binder In Me.col_binders
        If binder.binder_id = binder_id Then
            output = True
        End If
    Next binder
    
    is_binder_on_policy = output
    
outro:
    Exit Function
    
err_handler:
    GoTo outro
End Function
Public Sub init_binder_data(Optional ByVal rs_input As ADODB.Recordset)
    Dim proc_name As String
    proc_name = "cls_policy.init_binder_data"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim binder As cls_binder
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    Set Me.col_binders = New Collection
    If rs_input Is Nothing Then
        str_sql = "SELECT * FROM " & load.sources.policy_binders_view & " WHERE policy_id = " & Me.policy_id
        Set rs = utilities.create_adodb_rs(conn, str_sql)
        With rs
            .Open
            Do Until .EOF
                Set binder = New cls_binder
                binder.init rs
                Me.col_binders.Add binder
                binder.policy_id = Me.policy_id
                .MoveNext
            Loop
            .Close
        End With
    Else
        With rs_input
            .Filter = "policy_id = " & Me.policy_id
            Do Until .EOF
                Set binder = New cls_binder
                binder.init rs_input
                Me.col_binders.Add binder
                binder.policy_id = Me.policy_id
                .MoveNext
            Loop
            .Filter = ""
        End With
    End If
    
outro:
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = "str_sql = " & str_sql
        .params = "me.policy_id = " & Me.policy_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Public Function get_quota_for_default_alloation()
    Dim binder As cls_binder
    Dim quota_for_default_alloation As Variant
    quota_for_default_alloation = CDec(quota_for_default_alloation)
    
    quota_for_default_alloation = 0
    quota_for_default_alloation = (Me.policy_limit - Me.get_total_manual_limit) / Me.policy_limit
    
    get_quota_for_default_alloation = quota_for_default_alloation
End Function

Public Function get_total_reference_limit()
    Dim binder As cls_binder
    Dim total_reference_limit As Currency
    
    total_reference_limit = 0
    For Each binder In Me.col_binders
        If binder.on_policy_id = menu_list.yes Then
            total_reference_limit = total_reference_limit + binder.reference_limit
        End If
    Next binder
    
    get_total_reference_limit = total_reference_limit
End Function

Public Function get_max_limit_deal_currency()
    'Purpose: Get the total max limit in deal currency for all binders on policy
    Dim binder As cls_binder
    Dim max_limit_deal_currency As Currency
    
    max_limit_deal_currency = 0
    For Each binder In Me.col_binders
        If binder.on_policy_id = menu_list.yes Then
            max_limit_deal_currency = max_limit_deal_currency + binder.max_limit_deal_currency
        End If
    Next binder
    
    get_max_limit_deal_currency = max_limit_deal_currency
End Function

Public Function get_max_limit_for_default_allocation_deal_currency()
    Dim binder As cls_binder
    Dim max_limit_for_default_allocation_deal_currency As Currency
    
    max_limit_for_default_allocation_deal_currency = 0
    For Each binder In Me.col_binders
        If binder.manual_limit = 0 And binder.on_policy_id = menu_list.yes Then
            max_limit_for_default_allocation_deal_currency = max_limit_for_default_allocation_deal_currency + binder.max_limit_deal_currency
        End If
    Next binder
    
    get_max_limit_for_default_allocation_deal_currency = max_limit_for_default_allocation_deal_currency
End Function

Public Function get_total_reference_limit_without_binders_with_manual_limit()
    Const proc_name As String = "cls_policy.get_total_reference_limit_without_binders_with_manual_limit"
    utilities.call_stack_add_item proc_name
    
    Dim binder As cls_binder
    Dim total_reference_limit_without_binders_with_manual_limit As Currency
    
    total_reference_limit_without_binders_with_manual_limit = 0
    For Each binder In Me.col_binders
        If binder.on_policy_id = menu_list.yes And binder.manual_limit = 0 Then
            total_reference_limit_without_binders_with_manual_limit = total_reference_limit_without_binders_with_manual_limit + binder.reference_limit
        End If
    Next binder
    
    get_total_reference_limit_without_binders_with_manual_limit = total_reference_limit_without_binders_with_manual_limit
outro:
    utilities.call_stack_remove_last_item
    Exit Function
End Function
Public Function get_limit_for_default_allocation()
    get_limit_for_default_allocation = Me.policy_limit - Me.get_total_manual_limit
End Function

Public Function get_total_manual_limit()
    Dim binder As cls_binder
    Dim total_manual_limit As Currency
    
    total_manual_limit = 0
    For Each binder In Me.col_binders
        If binder.on_policy_id = menu_list.yes Then
            total_manual_limit = total_manual_limit + binder.manual_limit
        End If
    Next binder
    
    get_total_manual_limit = total_manual_limit
End Function

Public Function get_binder(ByVal input_binder_id As Long) As cls_binder
    Dim binder As cls_binder
    Dim output As cls_binder
    
    For Each binder In Me.col_binders
        If binder.binder_id = input_binder_id Then
            Set output = binder
            GoTo outro
        End If
    Next binder
    
outro:
    Set get_binder = output
    Exit Function
err_handler:

    GoTo outro
End Function
Public Function get_binder_count_without_manual_limit()
    Dim binder As cls_binder
    Dim binder_count As Long
    
    binder_count = 0
    For Each binder In Me.col_binders
        If binder.manual_limit = 0 _
        And binder.on_policy_id = menu_list.yes Then
            binder_count = binder_count + 1
        End If
    Next binder
    
    get_binder_count_without_manual_limit = binder_count
End Function

Public Sub init(ByVal rs As ADODB.Recordset)
    Dim proc_name As String
    proc_name = "cls_policy.init"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    With rs
        If .RecordCount = 0 Then
            GoTo outro
        End If
        
        Me.budget_home = Nz(!budget_home, "-1")
        Me.budget_home_id = Nz(!budget_home_id, "-1")
        
        Me.create_date = !create_date
        
        Me.display_view = !display_view
        Me.inception_date = Nz(!inception_date, "-1")
        Me.issuing_entity = Nz(!issuing_entity, "-1")
        Me.issuing_entity_id = Nz(!issuing_entity_id, "-1")
        Me.layer_limit = !layer_limit
        Me.layer_no = Nz(!layer_no, "-1")
        Me.layer_no_text = Nz(!layer_no_text, "-1")
        
        Me.policy_id = !policy_id
        Me.policy_name = Nz(!policy_name, "-1")
        Me.policy_no = !policy_no
        Me.policy_premium = !policy_premium
        Me.quota = !quota
        Me.underlying_limit = !underlying_limit
        
        If Me.col_binders Is Nothing Then
            Set Me.col_binders = New Collection
        End If
    End With

outro:
    Exit Sub

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
End Sub
Public Sub make_unrounded_limit_allocation(ByVal policy_object As cls_policy_object, ByVal rs_binder_list As ADODB.Recordset)
    '30 April 2025, CK: I don't think this is used anywhere, but saving ass I assume the logic shall be implemented as a method to cls_policy
    Dim proc_name As String
    proc_name = "central.make_unrounded_limit_allocation"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    With rs_binder_list
        Dim quota As Variant
        .MoveFirst
        Do Until .EOF
            If !manual_limit = 0 Then
                'Allocate limit for limit which is not manually set
                quota = !reference_limit_deal_currency / policy_object.total_reference_limit_without_binders_with_manual_limit * policy_object.quota_for_default_alloation
                
                ' Check if max limit is exceeded
                If quota * policy_object.policy_limit > !max_limit_deal_currency Then quota = !max_limit_deal_currency / policy_object.policy_limit
                
                'publish final quota for binder on policy
                conn.Execute "UPDATE " & sources.policy_binders_table & " SET quota = " & Replace(CStr(quota), ",", ".") & " WHERE id = " & !id
                load.rs_counter = load.rs_counter + 1
            Else
                quota = !manual_limit / policy_object.policy_limit
                conn.Execute "UPDATE " & sources.policy_binders_table & " SET quota = " & Replace(CStr(quota), ",", ".") & " WHERE id = " & !id
                load.rs_counter = load.rs_counter + 1
            End If
            .MoveNext
        Loop
    End With
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "central.make_unrounded_limit_allocation"
        .milestone = ""
        .params = "policy_object.policy_id = " & policy_object.policy_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro

End Sub

Public Sub find_default_limit_allocation(ByVal policy_object As cls_policy_object, ByVal rs_binder_list As ADODB.Recordset)
    '30 April 2025, CK: I don't think this is used anywhere, but saving as I assume the logic shall be implemented as a method to cls_policy
    Dim proc_name As String
    proc_name = "central.find_default_limit_allocation"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    With rs_binder_list
        Dim quota As Variant
        .MoveFirst
        Do Until .EOF
            ' Find default allocation
            quota = !reference_limit_deal_currency / policy_object.total_reference_limit
            
            ' Check if max limit is exceeded
            If quota * policy_object.policy_limit > !max_limit_deal_currency Then quota = !max_limit_deal_currency / policy_object.policy_limit
            
            'publish final quota for binder on policy
            conn.Execute "UPDATE " & sources.policy_binders_table & " SET default_quota = " & Replace(CStr(quota), ",", ".") & " WHERE id = " & !id
            load.rs_counter = load.rs_counter + 1
            
            'if manual limit is set, then override default quota
            If !manual_limit > 0 Then quota = !manual_limit / policy_object.policy_limit
            
            conn.Execute "UPDATE " & sources.policy_binders_table & " SET quota = " & Replace(CStr(quota), ",", ".") & " WHERE id = " & !id
            load.rs_counter = load.rs_counter + 1
            
            'track allocated limited and avalable max limit.
            'If allocated limit is less than policy limit and remaining_max_limit > remaining_policy_limit then do loop with remanining_max_limit as reference_amount
            .MoveNext
        Loop
    End With
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "central.find_default_limit_allocation"
        .milestone = ""
        .params = "policy_object.policy_id = " & policy_object.policy_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Sub
Public Function is_security_amount_equal_to_limit_for_policies_on_deal() As Boolean
    '30 April 2025, CK: I don't think this is used anywhere, but saving as I assume the logic shall be implemented as a method to cls_policy
    Dim proc_name As String
    proc_name = "central.is_security_amount_equal_to_limit_for_policies_on_deal"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim i As Integer
    Dim output As Boolean
    Dim policy As Scripting.Dictionary
    Dim policy_issue As Boolean
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    'select relevant policies and loop through to check that sum of quota is 1
    For Each policy In global_vars.deal.col_policies
        
        str_sql = "SELECT SUM(quota) quota_on_policy" _
        & " FROM " & sources.policy_binders_table _
        & " WHERE is_deleted = 0 AND policy_id = " & policy(policy_data.policy_id)
        
        Set rs = utilities.create_adodb_rs(conn, str_sql)
        rs.Open
            load.rs_counter = load.rs_counter + 1
            If Round(rs!quota_on_policy, 10) <> 1 Then
                policy_issue = True
            End If
        rs.Close
        
    Next policy
    
    output = True
    If policy_issue = True Then output = False
    is_security_amount_equal_to_limit_for_policies_on_deal = output
    
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
        .routine_name = "central.is_security_amount_equal_to_limit_for_policies_on_deal"
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
    
End Function

Public Function policy_limit() As Currency
    policy_limit = Me.layer_limit * Me.quota
End Function
