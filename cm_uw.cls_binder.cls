VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_binder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public added_by As String
Public binder_id As Long
Public binder_currency_date As Date
Public binder_name As String
Public binder_quota As Double

Public default_binder_currency As String
Public default_quota As Double
Public end_date As Date
Public extra_limit As Currency
Public for_eur_id As Long
Public for_us_id As Long

Public fx_deal_binder_rate As Variant

Public insurer_business_name As String
Public insurer_id As Long
Public insurer_legal_name As String
Public is_active_id As Long
Public limit_currency As String

Public manual_limit As Currency
Public manually_added As String
Public max_binder_limit_currency As String
Public max_binder_limit_ex_extra As Currency
Public max_limit_id As Long
Public minimum_underlying_limit As Currency

Public on_policy_id As Long

Public policy_id As Long

Public reference_limit As Currency
Public referral_status As String
Public referral_status_id As Long

Public security_id As Long
Public start_date As Date
Public unique_reference As String

Public Function get_binder_limit() As Currency
    Const proc_name As String = "cls_binder.get_binder_limit"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    Dim current_policy As cls_policy
    Dim output As Currency
    
    Set current_policy = global_vars.deal.get_policy(Me.policy_id)
    output = Me.manual_limit
    If Me.manual_limit = 0 Then
        output = CCur(Me.binder_quota * current_policy.policy_limit)
    End If
    get_binder_limit = output
outro:
    utilities.call_stack_remove_last_item
    Exit Function
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description _
    , "current_policy.policy_id = " & current_policy.policy_id _
    , "global_vars.deal.deal_id = " & global_vars.deal.deal_id _
    , "", True
    Resume outro
End Function
Public Function max_binder_limit() As Currency
    Dim output As Currency
    max_binder_limit = Me.max_binder_limit_ex_extra + Me.extra_limit
End Function
Public Function max_limit_deal_currency() As Currency
    Dim output As Currency
    
    If Me.fx_deal_binder_rate = 1 Then
        output = Me.max_binder_limit_ex_extra + Me.extra_limit
    Else
        output = (1 - global_vars.deal.deal_fx_buffer) * Me.max_binder_limit_ex_extra * Me.fx_deal_binder_rate + Me.extra_limit
        output = Round(output / 1000, 0) * 1000
    End If
    
    max_limit_deal_currency = output
    
End Function

Public Function reference_limit_deal_currency() As Currency
    Dim output As Currency
    
    If Me.fx_deal_binder_rate = 1 Then
        output = Me.reference_limit + Nz(Me.extra_limit, 0)
    Else
        '16 May 2025, CK: 0.96 is to give some cushion for currency fluctuations
        output = Round(0.96 * Me.reference_limit * Me.fx_deal_binder_rate, -3) + Nz(Me.extra_limit, 0)
    End If
    
    reference_limit_deal_currency = output
    
End Function

Public Function default_binder_limit() As Variant
    Dim current_policy As cls_policy
    Dim output As Variant
    Dim policy_limit As Currency
    
    Set current_policy = global_vars.deal.get_policy(Me.policy_id)
    output = Me.default_quota * current_policy.policy_limit
    
    default_binder_limit = output
    
    Set current_policy = Nothing
End Function
Public Sub refresh_binder_data()
    Dim proc_name As String
    proc_name = "cls_binder.refresh_data"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    If Me.security_id = 0 Then
        str_sql = "SELECT * FROM " & load.sources.policy_binders_view _
        & " WHERE binder_id = " & Me.binder_id & " AND policy_id = " & Me.policy_id
    Else
        str_sql = "SELECT * FROM " & load.sources.policy_binders_view _
        & " WHERE id = " & Me.security_id
    End If
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        Me.added_by = !added_by
        
        Me.binder_quota = !quota
        
        Me.default_binder_currency = !default_binder_currency
        Me.default_quota = !default_quota
        Me.extra_limit = !extra_limit
        Me.fx_deal_binder_rate = CDec(!max_limit_currency_to_deal_currency_fx)
        Me.manual_limit = !manual_limit
        Me.manually_added = !manually_added
        Me.max_binder_limit_currency = !max_binder_limit_currency
        Me.max_binder_limit_ex_extra = !max_binder_limit_ex_extra
        Me.max_limit_id = !max_limit_id
        Me.on_policy_id = !on_policy_id
        
        Me.reference_limit = !reference_limit
        Me.referral_status = !referral_status
        Me.referral_status_id = !referral_status_id
        .Close
    End With
outro:
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "me.binder_id = " & Me.binder_id _
        & ", global_vars.deal.deal_id = " & global_vars.deal.deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub init(ByVal rs As ADODB.Recordset)
    Dim proc_name As String
    proc_name = "cls_binder.init"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    With rs
        If .RecordCount = 0 Then
            GoTo outro
        End If
        Me.added_by = Nz(!added_by, "stella.birkelund")
        
        Me.binder_id = !binder_id
        Me.binder_name = !binder_name
        Me.binder_quota = !quota
        
        Me.default_binder_currency = !default_binder_currency
        Me.default_quota = !default_quota
        Me.extra_limit = !extra_limit
        Me.fx_deal_binder_rate = CDec(Nz(!max_limit_currency_to_deal_currency_fx, 0))
        Me.insurer_business_name = !insurer_business_name
        Me.insurer_legal_name = !insurer_legal_name
        Me.manual_limit = !manual_limit
        Me.manually_added = Nz(!manually_added, "no")
        Me.max_binder_limit_currency = Nz(!max_binder_limit_currency, "")
        Me.max_binder_limit_ex_extra = Nz(!max_binder_limit_ex_extra, 0)
        Me.max_limit_id = Nz(!max_limit_id, 0)
        Me.on_policy_id = !on_policy_id
        
        Me.reference_limit = !reference_limit
        Me.referral_status = !referral_status
        Me.referral_status_id = !referral_status_id
        
        Me.security_id = !security_id
        Me.unique_reference = !unique_reference
    End With

outro:
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "rs!policy_id = " & rs!policy_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
