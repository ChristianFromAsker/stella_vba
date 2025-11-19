Attribute VB_Name = "open_forms"
Option Compare Database
Option Explicit

Public Const second_view_left As Integer = 14000

Public Sub uw_positions_f(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "open_forms.uw_positions_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    load.check_conn_and_variables
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_name As String
    Dim question_count As Long
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    form_name = "uw_positions_f"
    If CurrentProject.AllForms(form_name).IsLoaded = False Then
        DoCmd.OpenForm form_name, acNormal
        DoCmd.MoveSize 500, 500, 15000, 12000
    End If
    With Forms(form_name)
        .Caption = "uw positions"
        !header_deal_id = deal_id
        str_sql = "SELECT deal_id, deal_name FROM " & load.sources.deals_table & " WHERE deal_id = " & deal_id
        Set rs = utilities.create_adodb_rs(conn, str_sql)
        rs.Open
            .Caption = "uw positions for " & rs!deal_name
            
        rs.Close
        Set rs = Nothing
        !question.SetFocus
        
        question_count = fix_rs.uw_positions_f(deal_id)
        
        !footer_lbl_question_count.Caption = "question count: " & question_count
        
        'Color code deal names
        
        'add color code questions based on status
        Dim objFormatConds As FormatCondition
        With .Controls("answer")
            .FormatConditions.Delete
            Set objFormatConds = .FormatConditions.Add(acExpression, , "[answer] <> [good_answer]")
            .FormatConditions(0).BackColor = colors.light_red
        End With
        
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
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Public Sub deal_procedures_f(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "open_forms.deal_procedures_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    'open and relocate form
    Dim str_form As String
    str_form = "deal_procedures_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    End If
    
    fix_rs.deal_procedures_f deal_id
    
    open_forms.resize_deal_procedures_f
    
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
Public Sub add_policy_f(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "open_forms.add_policy_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String, rs As ADODB.Recordset, str_form As String
    str_form = "add_policy_f"
    
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    End If
    
    With Forms(str_form)
        .SetFocus
        DoCmd.MoveSize 100, 100, 10000, 6000
        str_sql = "SELECT deal_id, deal_name FROM " & load.sources.deals_view & " WHERE deal_id = " & deal_id
        Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
            !deal_id = deal_id
            !deal_name = rs!deal_name
            .Caption = "add policy to " & rs!deal_name
        rs.Close
        'add values to combo boxes.
        'add the rest of the combo boxes
        Dim i As Long
        i = 1
        ReDim arr_controls(0 To 50, 0 To 1)
        
        arr_controls(i, 0) = "rp_on_layer"
        arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & load.sources.menu_list_table & " WHERE item_type =  'YesNo'"
        i = i + 1
        arr_controls(i, 0) = "layer_no"
        arr_controls(i, 1) = "SELECT layer_no id, layer_text menu_item FROM " & load.sources.policy_layer_texts_table & " WHERE is_deleted = 0 ORDER BY layer_no"
        i = i + 1
        arr_controls(i, 0) = "budget_home_id"
        arr_controls(i, 1) = "SELECT id, entity_business_name menu_item FROM " & load.sources.rp_entity_info_view & " WHERE entity_type = 475 ORDER BY entity_business_name"
        i = i + 1
        arr_controls(i, 0) = "issuing_entity_id"
        arr_controls(i, 1) = "SELECT id, entity_business_name menu_item FROM " & load.sources.rp_entity_info_view & " WHERE entity_type = 475 ORDER BY entity_business_name"
        
        arr_controls(0, 0) = i
        
        'remove existing items
        For i = 1 To arr_controls(0, 0)
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        
        'add new values
        Dim string_of_items As String
        For i = 1 To arr_controls(0, 0)
            string_of_items = ""
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
            Dim check_helper As Variant
            check_helper = 666
            Do While rs.EOF = False
                If check_helper = 666 Then
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                Else
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item & ";" & rs!menu_info
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next i
        
        'set default values
        !rp_on_layer = menu_list.yes
        !issuing_entity_id = current_uw.budget_home_id
        !budget_home_id = current_uw.budget_home_id
    End With
    
outro:
    If Not rs Is Nothing Then
       If rs.State <> 0 Then rs.Close
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
Public Sub working_on_it_f(ByVal input_text As String)
    Dim str_form As String
    str_form = "working_on_it_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    With Forms(str_form)
        !input_text.value = input_text
        !placeholder.SetFocus
    End With
    DoCmd.MoveSize Right:=200, Down:=200, Width:=10000, Height:=5000
    
End Sub
Public Sub working_on_it_f__close()
    On Error Resume Next
    If CurrentProject.AllForms(load.form_names.working_on_it).IsLoaded = True Then DoCmd.Close acForm, load.form_names.working_on_it
End Sub
Public Sub deal_referrals_help_f(ByVal template_question_id As Long, ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "open_forms.deal_referrals_help_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    'open and relocate form
    Dim str_form As String
    str_form = "deal_referrals_help_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    End If
    
    With Forms(str_form)
        
        Dim str_sql As String
        Dim rs As ADODB.Recordset
        
        str_sql = "SELECT id, question, referral_help, due_time_for_referral, referral_level" _
        & " FROM " & load.sources.template_qs_view _
        & " WHERE id = " & template_question_id
        
        Set rs = utilities.create_adodb_rs(conn, str_sql)
        rs.Open
        .Controls("template_question") = rs!question
        .Controls("referral_help") = rs!referral_help
        .Controls("referral_due_time") = rs!due_time_for_referral
        .Controls("referral_level") = rs!referral_level
        .Controls("template_question_id") = rs!id
        .Controls("deal_id") = deal_id
        
        Dim form_height As Long
        Dim line_count As Integer
        
        line_count = Len(!referral_help) / 100
        'adding a line for every occurance of <br> and <li>
        line_count = line_count + (Len(!referral_help) - Len(Replace(!referral_help, "<br>", ""))) / 4
        line_count = line_count + (Len(!referral_help) - Len(Replace(!referral_help, "<li>", ""))) / 4
        line_count = line_count + (Len(!referral_help) - Len(Replace(!referral_help, "<p>", ""))) / 3
        
        .Controls("referral_help").Height = utilities.twips_converter((line_count + 1) * 0.35, "cm") + 1000
        form_height = 4000 + .Controls("referral_help").Height
        
        rs.Close
        Set rs = Nothing
        
        .SetFocus
        .Controls("deal_id").SetFocus
        DoCmd.MoveSize Right:=50, Down:=200, Width:=12700, Height:=form_height
    End With
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "deal_id = " & deal_id _
        & vbNewLine & "template_question_id = " & template_question_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub deal_referrals_f(ByVal deal_id As Variant)
    Dim proc_name As String
    proc_name = "open_forms.deal_referrals_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    If load.is_debugging = True Then On Error GoTo 0
    load.check_conn_and_variables
    
    'open and relocate form
    Dim str_form As String
    str_form = "deal_referrals_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
        global_vars.deal_referrals.init
    End If
    
    fix_rs.deal_referrals_f deal_id
    open_forms.deal_referrals_f_resize
    Forms(str_form).SetFocus
    
    With Forms(str_form)
        .FormHeader.Height = 500
        
        'add color code questions based on status
        Dim objFormatConds As FormatCondition
        With .Controls("referral_status")
            .FormatConditions.Delete
            Set objFormatConds = .FormatConditions.Add(acExpression, , "[referral_status_id] <> " & load.binder_referral_ok & " and [referral_status_id] <> " & load.inar_referral_ok & " and [referral_status_id] <> " & load.no_binder_referral)
            .FormatConditions(0).BackColor = colors.yellow
        End With
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
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub deal_referrals_f_resize()
    Dim proc_name As String
    proc_name = "open_forms.deal_referrals_f_resize"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim detail_height As Integer
    Dim form_height
    Dim highest_char_count As Long
    Dim line_count As Integer
    Dim str_sql As String, rs As ADODB.Recordset
    Dim str_form As String
        
    str_form = "deal_referrals_f"
    str_sql = Forms(str_form).RecordSource
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        form_height = 2800
        If .RecordCount > 0 Then
            .MoveFirst
            'find appropriate form.detail.height
            highest_char_count = 0
            Do Until .EOF
                If Len(!referral_trigger) > highest_char_count Then
                    highest_char_count = Len(!referral_trigger)
                End If
                .MoveNext
            Loop
            line_count = highest_char_count / 50
            detail_height = utilities.twips_converter(line_count * 0.5 + 0.2, "cm")
            If detail_height < utilities.twips_converter(1.5, "cm") Then detail_height = utilities.twips_converter(2, "cm")
            With global_vars.deal_referrals
                .txt_template_question.field_height = utilities.twips_converter((line_count + 1) * 0.47, "cm")
            End With
            form_height = 3000 + rs.RecordCount * detail_height
        End If
        .Close
    End With
    Set rs = Nothing
    
    If form_height > 12000 Then form_height = 12000
    With Forms(str_form)
        .SetFocus
        DoCmd.MoveSize Right:=50, Down:=200, Width:=21000, Height:=form_height
        .Detail.Height = detail_height
    End With
    
    Dim form_control As cls_field
    For Each form_control In global_vars.deal_referrals.col_header_controls
        form_control.field_top = 50
        form_control.field_height = 50
        form_control.field_visible = False
    Next form_control
    
    global_vars.deal_referrals.paint_form global_vars.deal_referrals.col_all_controls
    
outro:
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
Public Sub binder_list_f(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "open_forms.binder_list_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String
    str_form = "binder_list_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    End If
    
    With Forms(str_form)
        .SetFocus
        DoCmd.MoveSize Right:=second_view_left, Down:=5000, Width:=load.binder_view_f_width, Height:=3900
        !header_deal_currency = global_vars.deal.deal_currency
        !lbl_extra_limit.Caption = "manual limit (" & global_vars.deal.deal_currency & ")"
        !lbl_limit_on_policy.Caption = "limit on policy (" & global_vars.deal.deal_currency & ")"
        !lbl_default_limit_on_policy.Caption = "default limit on policy (" & global_vars.deal.deal_currency & ")"
        !header_policy_limit.Visible = False
        !header_deal_id = deal_id
    End With
    
    'add values to comboboxes
    With Forms(str_form)
        'remove any existing lists
        Dim i As Integer
        i = 0
        Dim arr_controls() As Variant
        ReDim arr_controls(0 To 1, 0 To 1)
        arr_controls(i, 0) = "header_policy_id"
        arr_controls(i, 1) = "SELECT id, CONCAT(stella_policy_no, ' | ', layer_no_text, ' | ', IFNULL(policy_name, '')) menu_item FROM " & load.sources.policies_view _
        & " WHERE rp_on_layer = 93 AND deal_id = " & deal_id & " ORDER BY layer_no"
        
        i = i + 1
        arr_controls(i, 0) = "header_on_policy"
        arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & load.sources.menu_list_table _
            & " WHERE item_type = 'yesNo'": i = i + 1
        For i = 0 To UBound(arr_controls)
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        
        'add new values
        Dim check_helper As Variant, first_policy As Long
        Dim str_sql As String
        Dim rs As ADODB.Recordset
        For i = 0 To UBound(arr_controls)
            check_helper = 666
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
            Do While rs.EOF = False
                'several items don't have a menu info, and vba cannot check if fields exists (nor can it try-catch)
                On Error Resume Next
                    check_helper = rs!menu_info
                On Error GoTo err_handler
                If check_helper = 666 Then
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                Else
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item & ";'" & rs!menu_info & "'"
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
        
        fix_rs.binder_list_f
        
    End With
outro:
    If Not rs Is Nothing Then
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
Public Sub deal_questions_f(ByVal deal_id)
    Dim proc_name As String
    proc_name = "open_forms.deal_questions_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0

    Dim form_height As Long
    Dim objFormatConds As FormatCondition
    Dim record_count As Long
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    Dim str_form As String
    
    str_form = "deal_questions_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    End If
    Forms(str_form).SetFocus
    
    'add menu options to form
    With Forms(str_form)
        'add deal_id to form. Needed for other logic
        !header_deal_id = deal_id
        
        'set recordset
        fix_rs.deal_questions_f deal_id
        
        'add color code questions based on status
        Set objFormatConds = .Controls("question").FormatConditions.Add(acExpression, , "[answer] <> [good_answer]")
        Set objFormatConds = .Controls("answer_hr").FormatConditions.Add(acExpression, , "[answer] <> [good_answer]")
        .Controls("question").FormatConditions(0).BackColor = colors.light_red
        .Controls("answer_hr").FormatConditions(0).BackColor = colors.light_red
        
        Set rs = utilities.create_adodb_rs(conn, .RecordSource)
        rs.Open
            record_count = CLng(rs.RecordCount)
            form_height = 2500 + CLng(rs.RecordCount) * utilities.twips_converter(0.6, "inch")
            If form_height > 12000 Then form_height = 11000
        rs.Close
        .SetFocus
        DoCmd.MoveSize Right:=15000, Down:=200, Width:=18000, Height:=form_height
        
        If record_count > 0 Then
            .Controls("template_question_id").SetFocus
        End If
    End With
    
outro:
    If Not rs Is Nothing Then
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
Public Sub control_panel_f(ByVal deal_id)
    Dim proc_name As String
    proc_name = "open_forms.control_panel_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    'load form
    Dim str_form As String
    str_form = "control_panel_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    End If
    Forms(str_form).SetFocus
    DoCmd.MoveSize Right:=50, Down:=200, Width:=17500, Height:=11000
    With Forms(str_form)
        !deal_id = global_vars.deal.deal_id
        !deal_name = global_vars.deal.deal_name
        !insured_registered_country = global_vars.deal.insured_registered_country
        !deal_currency = global_vars.deal.deal_currency
        !risk_type = global_vars.deal.risk_type
        !status = global_vars.deal.deal_status
        !ev = global_vars.deal.ev
        
    End With
outro:
    Application.Echo True
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

Public Sub security_text_f(ByVal policy_id)
    Dim proc_name As String
    proc_name = "open_forms.security_text_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    'open and move form
    Dim str_form As String
    str_form = "security_text_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    End If
    Forms(str_form).SetFocus
    DoCmd.MoveSize Right:=200, Down:=200, Width:=9000, Height:=12000
    
    'prepare text
    Dim str_sql As String
    
    str_sql = "SELECT insurer_legal_name, unique_reference, quota FROM " & load.sources.policy_binders_view _
    & " WHERE on_policy_id = 93 AND policy_id = " & policy_id & " ORDER BY insurer_legal_name"
    
    Dim rs As ADODB.Recordset
    Dim security_text As String
    Dim unique_ref As String
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        If rs.BOF And rs.EOF Then
            security_text = "No binders on policy no " & rs!policy_no & "."
        Else
            Do Until rs.EOF
                unique_ref = ""
                If rs!unique_reference <> "" Then
                    unique_ref = rs!unique_reference & vbNewLine
                End If
                security_text = security_text & vbNewLine & rs!insurer_legal_name & vbNewLine _
                & unique_ref _
                & FormatPercent(rs!quota, 4) & vbNewLine
                rs.MoveNext
            Loop
        End If
    rs.Close
    Set rs = Nothing
    Forms(str_form)!security_text = Replace(security_text, ",", ".")
    
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
Public Sub policies_f(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "open_forms.policies_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String, rs As ADODB.Recordset, i As Integer
    Dim str_form As String
    str_form = "policies_f"
    str_sql = "SELECT deal_id, deal_name FROM " & sources.deals_view & " WHERE deal_id = " & deal_id
    Dim deal_name As String
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        deal_name = rs!deal_name
    rs.Close
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    End If
    fix_rs.policies_f deal_id
    
    'Move focus to top of list
    With Forms(str_form)
        .Caption = "Policies for " & deal_name & " (" & deal_id & ")"
        .SelTop = 1
        Dim form_height As Long
        str_sql = .RecordSource
        Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
            form_height = 2200 + utilities.twips_converter(3.6, "cm") * CLng(rs.RecordCount)
        rs.Close
        !cmd_add_layer.SetFocus
        .SetFocus
        DoCmd.MoveSize 500, 500, 15000, form_height
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
        .routine_name = "open_forms.policies_f"
        .milestone = ""
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Public Sub extra_deal_limit_f(ByVal deal_id)
    Dim proc_name As String
    proc_name = "open_forms.extra_deal_limit_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim rs As ADODB.Recordset
    'open and move form
    Dim str_form As String
    str_form = "extra_deal_limit_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    End If
    Dim str_sql As String
    fix_rs.extra_deal_limit_f deal_id

    'add values to comboboxes
    With Forms(str_form)
        Dim i As Integer
        i = 0
        Dim arr_controls() As Variant
        ReDim arr_controls(0 To 0, 0 To 1)
        arr_controls(i, 0) = "header_binder_id"
        arr_controls(i, 1) = "SELECT binder_id id, binder_name menu_item FROM " & load.sources.binder_list_view & " WHERE is_active = 93 ORDER BY binder_name": i = i + 1
        For i = 0 To UBound(arr_controls)
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        'add new values
        For i = 0 To UBound(arr_controls)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
    End With
    Forms(str_form).SetFocus
    Forms(str_form)!header_extra_deal_limit.SetFocus
    Dim form_height As Long
    Set rs = utilities.create_adodb_rs(conn, Forms(str_form).RecordSource): rs.Open
        form_height = 3000 + CLng(rs.RecordCount) * utilities.twips_converter(0.4, "inch")
    rs.Close
    DoCmd.MoveSize Right:=500, Down:=500, Width:=8500, Height:=form_height
    'load admin section
    str_sql = "SELECT deal_id, deal_currency, deal_name FROM deals_t WHERE deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
         With Forms(str_form)
            !header_deal_currency = rs!deal_currency
            !header_deal_id = rs!deal_id
            !header_deal_name = rs!deal_name
        End With
    rs.Close
outro:
    Application.Echo True
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: open_forms.extra_deal_limit_f" & vbNewLine _
        & "Parameters: deal_id = " & deal_id & vbNewLine _
        & "App: " & load.system_info.app_name, , load.system_info.error_msg_heading
    GoTo outro
End Sub

Public Sub resize_deal_procedures_f()
    Dim proc_name As String
    proc_name = "open_forms.resize_deal_procedures_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim str_sql As String, rs As ADODB.Recordset
    Dim str_form As String
    str_form = "deal_procedures_f"
    str_sql = Forms(str_form).RecordSource
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        Dim form_height
        form_height = 4000 + rs.RecordCount * utilities.twips_converter(1, "inch")
        If form_height > 8000 Then form_height = 8000
        Forms(str_form).SetFocus
        DoCmd.MoveSize Right:=50, Down:=200, Width:=12000, Height:=form_height
    rs.Close
End Sub
