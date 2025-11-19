Attribute VB_Name = "fix_rs"
Option Compare Database
Option Explicit

Public Function uw_positions_f(ByVal deal_id As Long _
, Optional search_condition As String _
, Optional order_by As String)

    Dim proc_name As String
    proc_name = "fix_rs.uw_positions"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "SELECT * FROM " & load.sources.uw_positions_view & " WHERE deal_id = " & deal_id
    If search_condition <> "" Then
        str_sql = str_sql & " AND " & search_condition
    End If
    
    str_sql = str_sql & " " & order_by
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        Set Forms("uw_positions_f").Recordset = rs
        uw_positions_f = CLng(.RecordCount)
        .Close
    End With
    Set rs = Nothing
        
outro:
    Exit Function

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
End Function
Public Function deal_procedures_f(ByVal deal_id As Variant) As Long
    Dim str_form As String, str_sql As String
    str_form = "deal_procedures_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then GoTo outro
    
    str_sql = "SELECT * FROM " & load.sources.cm_country_procedures_deals_view & " WHERE deal_id__deals_t = " & deal_id
    Dim rs As ADODB.Recordset
    With Forms(str_form)
        Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
            Set .Recordset = rs
            deal_procedures_f = CLng(rs.RecordCount)
        rs.Close: Set rs = Nothing
    End With
    
outro:
    Exit Function
    
End Function
Public Function policies_f(ByVal deal_id As Variant) As Long
    Dim str_sql As String
    str_sql = "SELECT * FROM " & load.sources.policies_view & " WHERE deal_id = " & deal_id & " ORDER BY layer_no ASC"
    Dim rs As ADODB.Recordset
    With Forms("policies_f")
        Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
            Set .Recordset = rs
            policies_f = CLng(rs.RecordCount)
        rs.Close: Set rs = Nothing
    End With
End Function

Public Sub binder_list_f()
    Const proc_name As String = "fix_rs.binder_list_f"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_height
    Dim policy_id As Long
    Dim rs As ADODB.Recordset
    Dim str_form As String
    Dim str_sql As String
    
    str_form = "binder_list_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    
    'if no policy is selected in the binders_f, then set recordset to blank id
    If Forms(str_form)!header_policy_id = "" Or IsNull(Forms(str_form)!header_policy_id) Then
        policy_id = 666
    Else
        policy_id = Forms(str_form)!header_policy_id.Column(0)
    End If
    
    str_sql = "SELECT * FROM " & load.sources.policy_binders_view & " WHERE policy_id = " & policy_id & " ORDER BY binder_name"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        With Forms(str_form)
            Set .Recordset = rs
            .SetFocus
            
            form_height = 3300 + rs.RecordCount * utilities.twips_converter(0.4, "inch")
            .InsideHeight = form_height
        End With
    rs.Close
    Set rs = Nothing

outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "policy_id = " & policy_id, "", "", True
    Resume outro
End Sub
Public Sub deal_questions_f(ByVal deal_id, Optional ByVal rs_deal_questions As ADODB.Recordset)
    Dim str_form As String
    str_form = "deal_questions_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        Exit Sub
    End If
    With Forms(str_form)
        If rs_deal_questions Is Nothing Then
            Dim str_sql As String
            
            str_sql = "SELECT * FROM " & sources.deal_questions_view _
            & " WHERE deal_id = " & deal_id _
            & " AND (" _
                & "question_type_id = " & question_types.binder _
                & " OR question_type_id = " & question_types.inar _
                & " OR question_type_id = " & question_types.confirmation _
            & ")" _
            & " ORDER BY sort_order, template_question_category"
            
            Dim rs As ADODB.Recordset
            Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
                Set .Recordset = rs
                !question_count.Caption = "There are " & rs.RecordCount & " questions."
            rs.Close: Set rs = Nothing
        Else
            Set .Recordset = rs_deal_questions
        End If
    End With
End Sub
Public Sub internal_questions_rs(ByVal deal_id)
    Dim str_form As String
    str_form = "internal_questions_f"
    If CurrentProject.AllForms(str_form).IsLoaded = True Then
        Forms(str_form).RecordSource = "SELECT * FROM cm_deal_questions_v WHERE question_type = 451 AND deal_id = " & deal_id
    End If
End Sub
Public Sub procedures_rs(ByVal deal_id)
    Dim str_form As String
    str_form = "procedures_f"
    If CurrentProject.AllForms(str_form).IsLoaded = True Then
        Forms(str_form).RecordSource = "SELECT * FROM cm_deal_questions_v WHERE (question_type = 461 OR question_type = 462) AND deal_id = " & deal_id
    End If
End Sub
Public Sub deal_referrals_f(ByVal deal_id As Variant)
    Dim str_form As String
    str_form = "deal_referrals_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        Exit Sub
    End If
    Dim str_sql As String
    str_sql = "SELECT * FROM " & sources.referrals_view & " WHERE deal_id = " & deal_id
    Dim rs As ADODB.Recordset
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        Set Forms(str_form).Recordset = rs
    rs.Close: Set rs = Nothing
End Sub
Public Sub extra_deal_limit_f(ByVal deal_id)
    Dim str_form As String
    str_form = "extra_deal_limit_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        Exit Sub
    End If
    Dim str_sql As String
    str_sql = "SELECT * FROM " & sources.extra_deal_limit_view & " WHERE deal_id = " & deal_id
    Dim rs As ADODB.Recordset
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        Set Forms(str_form).Recordset = rs
    rs.Close: Set rs = Nothing
End Sub
