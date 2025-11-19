VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_deal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public broker_firm As String
Public broker_firm_id As Long

Public closing_date As Variant
Public comments As String
Public create_date As Date
Public create_date_pretty As String
Public currency_date As Date
Public currency_date_manual As Long

Public deal_currency As String
Public deal_currency_to_eur As Variant
Public deal_fx_buffer As Variant
Public deal_id As Long
Public deal_name As String
Public deal_status As String
Public deal_status_id As Long

Public ev As Currency

Public fx_deal_eur_rate As Variant
Public inception_date As Variant
Public insured_main_country As String
Public insured_main_country_id As String
Public insured_main_region As String
Public insured_main_region_id As Long
Public insured_registered_country As String
Public insured_registered_country_id As Long
Public insured_registered_region As String
Public insured_registered_region_id As String

Public internal_approver As String
Public internal_approver_id As Long
Public internal_approver_at_signing As String
Public internal_approver_at_signing_id As Long

Public lead_insurer As String

Public max_limit_quoted As Currency
Public message_board As Variant

Public policy_period_in_months As Integer
Public primary_uw As String
Public primary_uw_id As Long
Public primary_or_xs As String
Public primary_or_xs_id As Long
Public program_limit As Currency
Public program_summary As String

Public retention As Currency
Public retention_drop_after As Currency
Public retention_drop_period As Integer
Public risk_type As String
Public risk_type_id As Long

Public seller_legal_name As String
Public SellerDomicile As String

Public spa_law As String
Public target_business_name As String
Public target_legal_name As String
Public target_naic_code As String
Public target_sic_code As String
Public target_sub_sector_id As Long
Public target_super_sector_id As Long

' Misc
Public any_binders_on_policy As Boolean
Public cm_shall_re_run As Long
Public has_currency_changed As Boolean
Public is_deal_found As Boolean
Public is_init As Boolean
Public is_locked As Boolean

'collections
Public col_policies As Collection

Public Sub remove_policy(ByVal input_policy_id As Long)
    Dim proc_name As String
    proc_name = "cls_deal.remove_policy"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim i As Long
    Dim policy As cls_policy
    
    i = 0
    For Each policy In Me.col_policies
        i = i + 1
        If policy.policy_id = input_policy_id Then
            Me.col_policies.Remove i
            GoTo outro
        End If
    Next policy
    
outro:
    Exit Sub

err_handler:
    GoTo outro
End Sub
Public Function get_policy(ByVal input_policy_id As Long) As cls_policy
    Dim proc_name As String
    proc_name = "cls_deal.get_policy"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim output As cls_policy
    Dim policy As cls_policy
    
    For Each policy In Me.col_policies
        If policy.policy_id = input_policy_id Then
            Set output = policy
        End If
    Next policy
    
    If Not output Is Nothing Then
        Set get_policy = output
    End If
End Function

Public Function get_policy_with_lowest_attachment() As cls_policy
    Dim proc_name As String
    proc_name = "cls_deal.get_policy_with_lowest_attachment"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim output As cls_policy
    Dim policy As cls_policy
    
    For Each policy In Me.col_policies
        With policy
            If output Is Nothing Then
                Set output = policy
            End If
            If output.underlying_limit > policy.underlying_limit Then
                Set output = policy
            End If
        End With
    Next policy
    
    If Not output Is Nothing Then
        Set get_policy_with_lowest_attachment = output
    End If
    
End Function
Public Sub print_deal()
    Dim proc_name As String
    proc_name = "cls_deal.print_deal"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    If Me.is_init = False Then
        Debug.Print "global_vars.deal.is_init = false"
        GoTo outro
    End If
    
    Dim binder As cls_binder
    Dim policy As cls_policy
    
    With Me
        Debug.Print "broker_firm: " & vbTab & .broker_firm
        Debug.Print "broker_firm_id: " & vbTab & .broker_firm_id
        Debug.Print "closing_date: " & vbTab & .closing_date
        Debug.Print "deal_currency: " & vbTab & .deal_currency
        Debug.Print "deal_id: " & vbTab & .deal_id
        Debug.Print "deal_name: " & vbTab & .deal_name
        For Each policy In .col_policies
            Debug.Print vbTab & "policy_id: " & vbTab & policy.policy_id
            Debug.Print vbTab & "policy_limit: " & vbTab & Format(policy.policy_limit, "#,##0")
            Debug.Print vbTab & "policy_no: " & vbTab & policy.policy_no
            If Not policy.col_binders Is Nothing Then
            For Each binder In policy.col_binders
                Debug.Print vbTab & vbTab & "binder_id: " & vbTab & binder.binder_id
                Debug.Print vbTab & vbTab & "binder_name: " & vbTab & binder.binder_name
                Debug.Print vbTab & vbTab & "max_limit_id:" & vbTab & binder.max_limit_id
                Debug.Print vbTab & vbTab & "fx_deal_binder_rate:" & vbTab & binder.fx_deal_binder_rate
                Debug.Print vbTab & vbTab & "default_binder_currency:" & vbTab & binder.default_binder_currency
                Debug.Print vbTab & vbTab & "max_binder_limit_currency:" & vbTab & binder.max_binder_limit_currency
                Debug.Print vbTab & vbTab & "max_binder_limit:" & vbTab & Format(binder.max_binder_limit, "#,##0")
                Debug.Print vbTab & vbTab & "max_limit_deal_currency:" & vbTab & Format(binder.max_limit_deal_currency, "#,##0")
                Debug.Print vbTab & vbTab & "binder_quota:" & vbTab & binder.binder_quota * 100 & "%"
                Debug.Print vbTab & vbTab & "binder_limit:" & vbTab & Format(binder.get_binder_limit, "#,##0")
                Debug.Print vbTab & vbTab & "on_policy_id:" & vbTab & binder.on_policy_id
                Debug.Print " "
            Next binder
            End If
        Next policy
    End With
    
outro:
    Exit Sub
    
End Sub
Public Sub init(ByVal deal_id_input As Long)
    Dim proc_name As String
    proc_name = "cls_deal.init"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Me.init_deal_data deal_id_input
    
    Me.init_policy_data deal_id_input
    
    Dim policy As cls_policy
    For Each policy In Me.col_policies
        policy.init_binder_data
    Next policy
    
    Me.is_init = True
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "deal_id_input = " & deal_id_input
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Sub

Public Sub init_deal_data(ByVal deal_id_input As Long)
    Dim proc_name As String
    proc_name = "cls_deal.init_deal_data"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim err_object As cls_err_object
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "SELECT * FROM " & load.sources.risk_details_f_view & " WHERE deal_id = " & deal_id_input
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        
        With Me
            .broker_firm_id = Nz(rs!broker_firm_id, -1)
            .broker_firm = Nz(rs!broker_firm, "-1")
            .create_date = rs!create_date
            .currency_date = rs!currency_date
            .currency_date_manual = rs!currency_date_manual
            
            .deal_currency = rs!deal_currency
            
            '19 December 2024, CK: Sometimes, the currency_rate_deal is not properly populated. Implemented special error handling for this.
            If IsNull(rs!currency_rate_deal) = False Then
                .deal_currency_to_eur = rs!currency_rate_deal
            Else
                deal_currency_to_eur = -1
                Set err_object = New cls_err_object
                With err_object
                    .routine_name = "cls_risk_object.init"
                    .milestone = "str_sql = " & str_sql
                    .params = "deal_id_input = " & deal_id_input
                    .stella_error_text = "rs!currency_rate_deal IS NULL"
                    .show_error_msg = False
                    .send_error err_object
                End With
                MsgBox "The CM has an issue with the currency exchange rate on this deal. Please let Tom and Christian know.", , "Fx rate issue on " & deal_id_input
            End If
            
            .deal_id = deal_id_input
            .deal_fx_buffer = Nz(rs!fx_buffer, 0)
            .deal_name = Nz(rs!deal_name, "-1")
            .deal_status = rs!deal_status
            .deal_status_id = rs!deal_status_id
            .ev = Nz(rs!ev, -1)
            .inception_date = Nz(rs!inception_date, -1)
            
            If IsNull(rs!insured_registered_country_id) = True Then
                MsgBox "The registered country of the insured could not be found. The CM will not work properly until this has been provided.", , "No registered country for the insured"
                GoTo outro
            End If
            .insured_registered_country_id = rs!insured_registered_country_id
            
            .insured_registered_country = rs!insured_registered_country
            .insured_main_region_id = Nz(rs!insured_main_region_id, -1)
            
            If load.system_info.app_continent = load.system_info.continents.americas Then
                .internal_approver_id = Nz(rs!internal_approver_quote_id, -1)
                .internal_approver_at_signing_id = Nz(rs!internal_approver_binding_id, -1)
            End If
            
            .is_init = True
            .is_locked = False
            If rs!is_locked = 1 Then .is_locked = True
            
            .policy_period_in_months = Nz(rs!policy_period_in_months, -1)
            
            .retention = rs!retention
            .retention_drop_after = Nz(rs!drop_end, -1)
            .retention_drop_period = Nz(rs!drop_period, -1)
            .risk_type_id = rs!risk_type_id
            .risk_type = rs!risk_type
            
            .spa_law = Nz(rs!spa_law, -1)
            .target_business_name = Nz(rs!target_business_name, "-1")
            .target_super_sector_id = Nz(rs!target_super_sector_id, -1)
            .target_sub_sector_id = Nz(rs!target_sub_sector_id, -1)
            
            .max_limit_quoted = Nz(rs!max_limit_quoted, -1)
            
            If .col_policies Is Nothing Then
                Set .col_policies = New Collection
            End If
        End With
    rs.Close
    
outro:
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Exit Sub
    
err_handler:
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = "str_sql = " & str_sql
        .params = "deal_id_input = " & deal_id_input
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub init_policy_data(ByVal deal_id_input As Long)
    Dim proc_name As String
    proc_name = "cls_deal.init_policy_data"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim policy As cls_policy
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "SELECT * FROM " & load.sources.policies_view & " WHERE deal_id = " & deal_id_input
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    Set Me.col_policies = New Collection
    With rs
        .Open
        Do Until .EOF
            Set policy = New cls_policy
            policy.init rs
            Me.col_policies.Add policy
            .MoveNext
        Loop
        
        .Close
    End With
    
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
        .params = "deal_id_input = " & deal_id_input
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Property Get target_path() As String
    target_path = Left(target_business_name, 15)
End Property
Property Get create_year()
    create_year = Year(create_date)
End Property
Property Get inception_year()
    If IsNull(inception_date) = True Then
        inception_year = -1
    Else
        inception_year = Year(inception_date)
    End If
End Property
Property Get inception_date_pretty()
    If IsNull(inception_date) = True Then
        inception_date_pretty = -1
    Else
        inception_date_pretty = Format(inception_date, "Long Date")
    End If
End Property
Public Sub populate_key_data(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "cls_deal.populate_key_data"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim col_key_data As Collection
    Dim key_data As Variant
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    Me.deal_id = deal_id
    
    Set col_key_data = central.get_list_of_key_data
    
    str_sql = "SELECT * FROM " & load.sources.cm_key_data_table _
    & " WHERE deal_id = " & deal_id _
    & " AND is_deleted = 0"
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        If rs.RecordCount = 0 Then
            Me.is_deal_found = False
        Else
            Me.is_deal_found = True
            For Each key_data In col_key_data
                CallByName Me, key_data, VbLet, Nz(rs.Fields(key_data), "-1")
            Next key_data
            Me.cm_shall_re_run = rs!shall_re_run
        End If
    rs.Close
    Set rs = Nothing
        
End Sub

Public Function binder_count()
    Dim binder As cls_binder
    Dim binder_counter As Long
    Dim policy As cls_policy

    binder_counter = 0
    For Each policy In Me.col_policies
        If Not policy.col_binders Is Nothing Then
            binder_counter = binder_counter + policy.col_binders.Count
        End If
    Next policy
    
    binder_count = binder_counter
End Function
