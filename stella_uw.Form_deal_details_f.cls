VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_deal_details_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub broker_firm_id_AfterUpdate()
    load.check_conn_and_variables
    With load.deal_details.txt_broker_firm_id
        utilities.update_data .field_name, Me.Controls(.field_name_in_recordset), Me!deal_id
        .field_bg_color = load.colors.white
    End With

    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "broker_firm_id"
        .new_value = Me!broker_firm_id
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
err_handler:
    GoTo outro
    
End Sub
Private Sub budget_home_id_AfterUpdate()
    Dim proc_name As String
    proc_name = "deal_details_f.budget_home_id.AfterUpdate"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    load.check_conn_and_variables
    
    Central.budget_home_change Me!deal_id, Me!budget_home_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub closing_premium_due_date_AfterUpdate()
    With Me!closing_premium_due_date
        utilities.update_data .Name, utilities.generate_sql_date(.value), Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    
End Sub

Private Sub closing_premium_received_date_AfterUpdate()
    With Me!closing_premium_received_date
        utilities.update_data .Name, utilities.generate_sql_date(.value), Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    
End Sub

Private Sub closing_set_received_id_AfterUpdate()
    utilities.update_data "closing_set_received_id", Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
End Sub

Private Sub cmd_expand_admin_section_Click()
    load.check_conn_and_variables
    With load.deal_details
        If .is_init = False Then .init deal_id
        .change_section_visible_status .col_labels_for_admin_section
        .change_section_visible_status .col_txt_fields_for_admin_section
        If .header_admin.is_collapsed = True Then
            With .header_admin
                .is_collapsed = False
                .user_has_expanded_section = True
            End With
            .cmd_expand_admin_section.field_caption = "collapse"
            .put_data_into_form .col_txt_fields_for_admin_section, Me!deal_id
        Else
            With .header_admin
                .is_collapsed = True
                .user_has_expanded_section = False
            End With
            .cmd_expand_admin_section.field_caption = "expand"
        End If
        .refresh_first_column
        .paint_deal_details
    End With
    
outro:
    load.call_stack = ""
    
End Sub

Private Sub cmd_expand_closing_info_section_Click()
    load.check_conn_and_variables
    With load.deal_details
        If .is_init = False Then .init deal_id
        .change_section_visible_status .col_labels_for_closing_info
        .change_section_visible_status .col_txt_fields_for_closing_info
        If .header_closing_info.is_collapsed = True Then
            With .header_closing_info
                .is_collapsed = False
                .user_has_expanded_section = True
            End With
            .cmd_expand_closing_info_section.field_caption = "collapse"
            .put_data_into_form .col_txt_fields_for_closing_info, Me!deal_id
        Else
            With .header_closing_info
                .is_collapsed = True
                .user_has_expanded_section = False
            End With
            .cmd_expand_closing_info_section.field_caption = "expand"
        End If
        .paint_deal_details
    End With
    
outro:
    load.call_stack = ""
    
End Sub

Private Sub cmd_expand_submission_notes_section_Click()
    load.check_conn_and_variables
    With load.deal_details
        If .is_init = False Then .init deal_id
        .change_section_visible_status .col_labels_for_submission_notes_section
        .change_section_visible_status .col_txt_fields_for_submission_notes_section
        If .header_submission_notes.is_collapsed = True Then
            With .header_submission_notes
                .is_collapsed = False
                .user_has_expanded_section = True
            End With
            .cmd_expand_submission_notes_section.field_caption = "collapse"
            .put_data_into_form .col_txt_fields_for_submission_notes_section, Me!deal_id
        Else
            With .header_submission_notes
                .is_collapsed = True
                .user_has_expanded_section = False
            End With
            .cmd_expand_submission_notes_section.field_caption = "expand"
        End If
        .refresh_third_column
        .paint_deal_details
    End With
    
outro:
    load.call_stack = ""
    
End Sub

Private Sub cmd_general_tags_Click()
    open_forms.deal_tags_f Me!deal_id
    
outro:
    load.call_stack = ""
    
End Sub

Private Sub cmd_general_uw_positions_Click()
    Dim proc_name As String
    proc_name = "deal_details_f.cmd_general_uw_positions"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    load.check_conn_and_variables
    
    Dim deny_open As Boolean
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    deny_open = False
    
    str_sql = "SELECT deal_question_id " _
    & " FROM " & load.sources.deal_questions_view _
    & " WHERE deal_id = " & Me!deal_id _
    & " AND question_type_id = " & 400
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        If CLng(.RecordCount) = 0 Then
            deny_open = True
        End If
        .Close
    End With
    Set rs = Nothing
    
    If deny_open = False Then
        open_forms.uw_positions_f Me!deal_id
        
        str_sql = "INSERT INTO " & load.sources.log_activity_table & " (activity, activity_source, user_name, app_name, app_continent) VALUES (" _
        & "'opened uw_positions_f'" _
        & ", '" & proc_name & "'" _
        & ", '" & Environ("username") & "'" _
        & ", '" & load.system_info.app_name & "'" _
        & ", '" & load.system_info.app_continent & "'" _
        & ")"
        conn.Execute str_sql
    Else
        MsgBox "You need to open the CM once before you can open the UW positions.", , "Open CM once first"
    End If
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub cmd_generate_rr_Click()
    load.call_stack = "deal_details_f.cmd_generate_rr"
    load.deal_details.print_rr Me!deal_id
    
outro:
    load.call_stack = ""
    
End Sub

Private Sub cmd_open_workspace_Click()
    load.call_stack = "deal_details_f.cmd_open_workspace"
    load.check_conn_and_variables
    
    Dim col_missing_fields As Collection, form_field As cls_field
    Set col_missing_fields = mandatory_data.check_data(deal_id, load.mandatory_tests.imanage_check)
    If col_missing_fields.Count > 0 Then
        For Each form_field In col_missing_fields
            With form_field
                .field_bg_color = load.colors.yellow
            End With
        Next form_field
        GoTo outro
    End If
    
    imanage.open_workspace Me!deal_id, load.current_uw.user_name & "@rpgroup.com"
    
outro:
    load.call_stack = ""
    open_forms.working_on_it_f__close
    Exit Sub
    
End Sub

Private Sub cmd_refresh_nbi_templates_Click()
    load.call_stack = "deal_details_f.cmd_refresh_nbi_templates"
    load.check_conn_and_variables
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    'check if mandatory fields are populated
    If IsNull(Me!spa_law) Or IsEmpty(Me!spa_law) Or Me!spa_law = "" Then
        MsgBox "Please provide SPA law to refresh templates.", , "SPA law missing"
        GoTo outro
    End If
    
    'find deal folder
    With load.system_info.texts
        open_forms.working_on_it_f .refreshing_nbi_templates_heading, .refreshing_nbi_templates_details, 4000
    End With
    
    Dim obj_deal As cls_deal, deal_path As String, folder_found As Boolean
    Set obj_deal = New cls_deal
    obj_deal.init_deal_object Me!deal_id
    
    'establish where deal folder should be
    deal_path = Paths.folder_path_from_scripts(Me!deal_id, "normative")
    If deal_path = "-1" Then
        MsgBox "I've looked in the basement, and attic, and all the storage rooms, but I simply cannot find the deal folder." & vbNewLine & vbNewLine _
        & "Maybe the bats ate it, or someone stole it. Or misplaced it.", , "No deal folder!"
        GoTo outro
    End If
    
    'create path for new template folder inside nbi folder
    Dim new_template_path As String, counter As String, i As Integer
    new_template_path = deal_path & "\NBI etc\" & Year(Date) & Month(Date) & Day(Date) & " NBI templates"
    
    'Create new folder for refreshed templates ot avoid overwriting existing templates
    counter = ""
    i = 1
    Do Until Dir(new_template_path & counter, vbDirectory) = ""
        counter = "_" & i
        i = i + 1
    Loop
    new_template_path = new_template_path & counter
    
    'create new template folder and copy templates
    MkDir new_template_path
    Dim target_rating_tool_path As String
    target_rating_tool_path = new_template_path & "\" & deal_name & " - Rating Tool" & ".xlsm"
    FileCopy load.system_info.system_paths.rating_tool_template, target_rating_tool_path
    FileCopy Central.get_nbi_template_path(, Me!spa_law), new_template_path & "\" & deal_name & " - RiskPoint NBI" & ".docx"
    FileCopy load.system_info.system_paths.template_path & "Emails\NBI_approval_request.msg", new_template_path & "\" & deal_name & " - NBI_approval_request.msg"
    FileCopy load.system_info.system_paths.template_path & "Empty document templates\Key issues.docx", new_template_path & "\" & deal_name & " - key issues.docx"
    
    'put deal_id into rating tool
    Dim appExcel As Excel.Application
    Set appExcel = New Excel.Application
    appExcel.Visible = False
    Dim wb As workbook
    Set wb = appExcel.Workbooks.Open(target_rating_tool_path, ReadOnly:=False)
        Dim ws_input As Worksheet
        Set ws_input = wb.Worksheets("input")
            ws_input.Range("c3") = deal_id
            wb.Save
        Set ws_input = Nothing
        wb.Close
    Set wb = Nothing
    Set appExcel = Nothing
    
    Application.FollowHyperlink new_template_path
    
outro:
    open_forms.working_on_it_f__close
    Set obj_deal = Nothing
    load.call_stack = ""
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "deal_details_f.cmd_refresh_nbi_templates"
        .milestone = ""
        .params = "me!deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
End Sub

Private Sub cmd_web_search_Click()
    Dim proc_name As String
    proc_name = "deal_deatils_f.cmd_web_search.click"
    load.call_stack = proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
        
    Dim col_missing_fields As Collection
    Dim form_field As cls_field
    
    load.check_conn_and_variables
    Set col_missing_fields = mandatory_data.check_data(deal_id, load.mandatory_tests.web_searches)
    If col_missing_fields.Count > 0 Then
        For Each form_field In col_missing_fields
            With form_field
                .field_bg_color = load.colors.yellow
            End With
        Next form_field
        load.deal_details.paint_deal_details
        GoTo outro
    End If
    
    open_forms.web_search_f Me!deal_id
    
    'to dos. CHRISTIAN YOU GOT HERE
        'Log the search and the current sector
        'add link to guide document
    
outro:
    load.call_stack = ""
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro

End Sub

Private Sub counsel_fee_AfterUpdate()
    load.check_conn_and_variables
    Me.SetFocus
    conn.Execute "UPDATE " & sources.deals_table & " SET counsel_fee_amount = " & Screen.ActiveControl & "  WHERE deal_id = " & Me!deal_id
    load.deal_details.txt_counsel_fee.field_bg_color = load.colors.white
    Dim col_fields As New Collection
    col_fields.Add load.deal_details.txt_counsel_fee
    utilities.paint_control load.deal_details.form_name, col_fields
    load.deal_details.paint_financial_info_section Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub counsel_fee_due_date_AfterUpdate()
    With Me!counsel_fee_due_date
        utilities.update_data .Name, utilities.generate_sql_date(.value), Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub counsel_fee_received_date_AfterUpdate()
    With Me!counsel_fee_received_date
        utilities.update_data .Name, utilities.generate_sql_date(.value), Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub deal_info_AfterUpdate()
    utilities.update_data "deal_info", Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub deal_status_id_GotFocus()
    Dim proc_name As String
    proc_name = "deal_details_f.deal_status_id.GotFocus"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    load.check_conn_and_variables

    With load.deal_details
        If .is_init = False Then .init deal_id
        With .help_text
            .field_value = load.system_info.texts.deal_status_got_focus
            .field_top = load.deal_details.txt_deal_status_id.field_top
            .field_left = load.deal_details.txt_deal_status_id.field_left + load.deal_details.txt_deal_status_id.field_width
            .field_height = 8000
            .field_width = 10000
            .field_visible = True
            .field_bg_color = load.colors.very_light_green
            .field_type = load.field_type.text_field
        End With
    End With
    
    Dim col_fields As New Collection
    Set col_fields = Nothing
    col_fields.Add load.deal_details.help_text
    utilities.paint_control load.deal_details.form_name, col_fields
    Set col_fields = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub deal_status_id_AfterUpdate()
    Dim proc_name As String
    proc_name = "deal_details_f.deal_status_id.AfterUpdate"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    Dim str_milestone As String
    str_milestone = "1"
    
    Dim col_fields_update As New Collection
    Dim col_missing_fields As New Collection
    Dim form_field As cls_field
    Dim old_deal_status_id As Long
    Dim test_type
        
    'close help text
    Set col_fields_update = Nothing
    With load.deal_details
        If .is_init = False Then .init deal_id
        With .help_text
            .field_visible = False
        End With
        col_fields_update.Add .help_text
    End With
    utilities.paint_control load.deal_details.form_name, col_fields_update
    Set col_fields_update = Nothing
    
    str_milestone = "2"
    old_deal_status_id = Central.get_current_risk_status(Me!deal_id)
    
    str_milestone = "3"
    Select Case CStr(Me!deal_status_id)
    
        'working statuses
        Case CStr(deal_statuses.NDA)
            test_type = load.mandatory_tests.nda_check
        Case CStr(deal_statuses.submission)
            test_type = load.mandatory_tests.nda_check
        Case CStr(deal_statuses.nbi)
            test_type = load.mandatory_tests.quote_check
        Case CStr(deal_statuses.preferred)
            test_type = load.mandatory_tests.quote_check
        Case CStr(deal_statuses.expensed)
            test_type = load.mandatory_tests.quote_check
        Case CStr(deal_statuses.uw)
            test_type = load.mandatory_tests.quote_check
            
        'non bound statuses
        Case CStr(deal_statuses.cancelled)
            test_type = load.mandatory_tests.non_bound_check
        Case CStr(deal_statuses.declined)
            test_type = load.mandatory_tests.non_bound_check
        Case CStr(deal_statuses.lost)
            test_type = load.mandatory_tests.non_bound_check
        
        'bound statuses
        Case CStr(deal_statuses.signed)
            test_type = load.mandatory_tests.signing_check
        Case CStr(deal_statuses.closed)
            test_type = load.mandatory_tests.closing_check
        Case Else
    
    End Select
    
    str_milestone = "4"
    Set col_missing_fields = Nothing
    Set col_missing_fields = mandatory_data.check_data(Me!deal_id, test_type)
    
    For Each form_field In col_missing_fields
        form_field.field_bg_color = load.colors.yellow
        col_fields_update.Add form_field
    Next form_field
    
    If col_missing_fields.Count > 0 Then
        Me!deal_status_id = old_deal_status_id
        GoTo outro
    End If
    
    str_milestone = "5"
    Central.deal_status_was_changed Me!deal_id, Me!deal_status_id, False, old_deal_status_id
    
    'turn field white
    load.deal_details.txt_deal_status_id.field_bg_color = load.colors.white
    col_fields_update.Add load.deal_details.txt_deal_status_id
    
outro:
    utilities.paint_control load.deal_details.form_name, col_fields_update
    open_forms.working_on_it_f__close
    load.call_stack = ""
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .milestone = str_milestone
        .params = "me!deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .routine_name = "deal_details_f.deal_status_id_AfterUpdate"
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
    
End Sub

Private Sub deal_status_id_LostFocus()
    Dim proc_name As String
    proc_name = "deal_details_f.deal_status_id.LostFocus"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    load.check_conn_and_variables
    
    'close help text
    With load.deal_details
        If .is_init = False Then .init deal_id
        With .help_text
            .field_visible = False
        End With
    End With
    
    Dim col_fields As New Collection
    Set col_fields = Nothing
    col_fields.Add load.deal_details.help_text
    utilities.paint_control load.deal_details.form_name, col_fields
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub drop_period_AfterUpdate()
    load.check_conn_and_variables
    
    With Me!drop_period
        utilities.update_data .Name, .value, Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub emails_AfterUpdate()
    Dim proc_name As String
    proc_name = "deal_details_f.emails_AfterUpdate"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    load.check_conn_and_variables
    
    Dim form_field As cls_field
    Dim obj_deal As New cls_deal
    Dim rs As ADODB.Recordset
    Dim prepare_email As Boolean
    Dim str_sql As String
    Dim col_fields_with_missing_data As New Collection
    
    Set col_fields_with_missing_data = Nothing
    If Me!emails = 0 Then
        MsgBox "You must choose an email template from the drop-down.", , "No email selected"
        GoTo outro
    End If
    
    If Me!emails = 5 Then
        If load.system_info.app_continent <> load.system_info.continents.americas Then
            MsgBox "This is only available in America", , "Sorry"
            GoTo outro
        End If
                
        'check if mandatory fields are present on deal level
        Dim fields_with_issues As New Collection, data_field As Scripting.Dictionary, field_name As String
        Set fields_with_issues = data_verification.capacity_email_us(deal_id)
        If fields_with_issues.Count > 0 Then
            For Each data_field In fields_with_issues
                field_name = data_field(load.global_field_properties.field_name_in_deal_details_f)
                If field_name <> "-1" Then
                    With Forms(Me.Name).Controls(field_name)
                        .BackColor = load.colors.yellow
                    End With
                End If
            Next data_field
            GoTo outro
        End If
        
        With load.system_info.texts
            open_forms.working_on_it_f .preparing_email_heading, .preparing_email_details, 4000
        End With
        
        email_module.capacity_email_us deal_id
    End If
    
    obj_deal.init_deal_object Me!deal_id
    
    If Me!emails = 4 Then
        prepare_email = True
        
        'check if data is mssing on deal level
        Set col_fields_with_missing_data = mandatory_data.check_data(Me!deal_id, load.mandatory_tests.signing_check)
        If col_fields_with_missing_data.Count > 0 Then
            For Each form_field In col_fields_with_missing_data
                form_field.field_bg_color = load.colors.yellow
                prepare_email = False
            Next form_field
        End If
        
        'check if policy premium is different from the default premium
        str_sql = "SELECT id, deal_id, policy_premium FROM " & load.sources.policies_table & " WHERE is_deleted = 0 AND deal_id = " & deal_id
        Set rs = utilities.create_adodb_rs(conn, str_sql)
        rs.Open
        Do Until rs.EOF
            If rs!policy_premium = 0 Then
                prepare_email = False
                MsgBox "No premium is added to the policy. Please open the Compliance Module and provide data for the policy", , "Missing policy data"
            End If
            rs.MoveNext
        Loop
        rs.Close
        
        If Me!deal_status_id <> 6 And Me!deal_status_id <> 436 Then
            load.deal_details.txt_deal_status_id.field_bg_color = load.colors.yellow
            prepare_email = False
            MsgBox "Status must be signed or closed to request booking and/or an invoice.", , "Fix status"
        End If
        
        If prepare_email = True Then
            email_module.prepare_signing_invoice obj_deal
        Else
            load.deal_details.paint_deal_details
        End If
        
        GoTo outro
    End If
    
    Dim deal_path As String
    deal_path = Paths.folder_path_from_scripts(Me!deal_id, "find")
    If Dir(deal_path, vbDirectory) = vbNullString Then
        MsgBox "No deal folder found for this deal. Please create a deal folder before continuing.", , "No deal folder"
        GoTo outro
    End If
    
    If Me!emails = 1 Then
        FileCopy load.system_info.system_paths.base_path & "templates\emails\NBI_approval_request.msg", deal_path & "\Miscellaneous\NBI_approval_request.msg"
        Application.FollowHyperlink deal_path & "\Miscellaneous\NBI_approval_request.msg"
    End If
    
    If Me!emails = 2 Then
        FileCopy load.system_info.system_paths.base_path & "templates\emails\signing_pr_request.msg", deal_path & "\Miscellaneous\signing_pr_request.msg"
        Application.FollowHyperlink deal_path & "\Miscellaneous\signing_pr_request.msg"
    End If
    
    If Me!emails = 3 Then
        FileCopy load.system_info.system_paths.base_path & "templates\emails\closing_pr_request.msg", deal_path & "\Miscellaneous\closing_pr_request.msg"
        Application.FollowHyperlink deal_path & "\Miscellaneous\closing_pr_request.msg"
    End If
    
outro:
    open_forms.working_on_it_f__close
    load.call_stack = ""
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = "n/a"
        .params = "me!deal_id = " & Me!deal_id & ", me!emails = " & Me!emails
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Sub

Private Sub esg_company_AfterUpdate()
    With Me!esg_company
        utilities.update_data .Name, .value, Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub esg_country_AfterUpdate()
    With Me!esg_country
        utilities.update_data .Name, .value, Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim str_form As String
    str_form = "deal_details_f"
    Forms(str_form).Visible = False
    If load.is_debugging = False Then Cancel = True
    If load.stella_is_closing = True Or load.deal_details.shall_be_refreshed = True Then
        Cancel = False
    End If
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub insured_legal_name_AfterUpdate()
    load.check_conn_and_variables
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "insured_legal_name"
        .new_value = Me!insured_legal_name
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
    imanage.update_workspace_caller Me!deal_id, load.current_uw.email
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub buyer_business_name_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub buyer_law_firm_1_id_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub buyer_law_firm_2_id_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub insured_main_region_id_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
            
End Sub

Private Sub insured_registered_country_id_AfterUpdate()
    load.check_conn_and_variables
    conn.Execute "UPDATE " & sources.deals_table & " SET insured_registered_country_id = " & Me!insured_registered_country_id & " WHERE deal_id = " & Me!deal_id
    load.deal_details.txt_insured_registered_country_id.field_bg_color = load.colors.white
    Me!insured_registered_country_id.BackColor = load.colors.white
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "insured_registered_country_id"
        .new_value = Me!insured_registered_country_id
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
    load.deal_details.add_regions_to_insured_main_region_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub BuyerFinancialCounsel_AfterUpdate()
    utilities.update_data "buyer_financial_firm_id", Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub c5_emails_filed_AfterUpdate()
    load.check_conn_and_variables
    conn.Execute "UPDATE " & sources.deals_table & " SET are_emails_filed_id = " & Me!c5_emails_filed & " WHERE deal_id = " & Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub closing_date_AfterUpdate()
    load.check_conn_and_variables
    Dim output_value As String
    
    If IsNull(Me!closing_date) = True Or Me!closing_date = "" Then
        output_value = "NULL"
    Else
        output_value = utilities.generate_sql_date(Me!closing_date)
    End If
    
    utilities.update_data "closing_date", output_value, Me!deal_id

    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "closing_date"
        .new_value = output_value
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub ClosingBooked_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub cmdDelete_Click()
    Const proc_name As String = "deal_details_f.cmdDelete.Click"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    load.event_id = utilities.get_event_id
    
    Dim col_binders As Collection
    Dim col_policies As Collection
    Dim deal_id As Long
    Dim log_object As utilities.typ_log_object
    Dim policy_id
    Dim rs As ADODB.Recordset
    Dim security_id
    Dim str_sql As String
    Dim str_sql_binders As String
    Dim str_sql_deal As String
    Dim str_sql_policy As String
    Dim tx_started As Boolean
    
    If MsgBox("Are you sure you want to delete the deal?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm delete") = vbNo Then
        Exit Sub
    End If
    
    deal_id = Me!deal_id
    
    'deal
    str_sql_deal = "UPDATE deals_t SET is_deleted = 1 WHERE deal_id = " & deal_id
    
    load.conn.BeginTrans
    tx_started = True
    
        conn.Execute str_sql_deal
        
        With log_object
            .change_source = proc_name
            .changer_id = Environ("username")
            .comment = ""
            .data_set = load.sources.deals_table
            .deal_id = deal_id
            .event_id = load.event_id
            .executed_sql = str_sql
            .field_name = "is_deleted"
            .new_value = 1
            .operation_type = "update"
            .policy_id = 0
            .record_id = deal_id
            .security_id = 0
        End With
        
        utilities.log_change log_object
    
    load.conn.CommitTrans
    tx_started = False
    
    'policies
    str_sql_policy = "UPDATE " & load.sources.policies_table & " SET is_deleted = 1 WHERE deal_id = " & deal_id
    
    str_sql = "SELECT id policy_id FROM " & load.sources.policies_table & " WHERE is_deleted = 0 AND deal_id = " & deal_id
    
    Set col_policies = New Collection
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        Do Until .EOF
            col_policies.Add !policy_id.value
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    load.conn.BeginTrans
    tx_started = True
            
        conn.Execute str_sql_policy
                
        For Each policy_id In col_policies
            With log_object
                .data_set = load.sources.policies_table
                .policy_id = policy_id
                .record_id = policy_id
            End With
            utilities.log_change log_object
        Next policy_id
            
    load.conn.CommitTrans
    tx_started = False
    
    'binders
    str_sql = ""
    For Each policy_id In col_policies
        str_sql_binders = "UPDATE " & load.sources.policy_binders_table & " SET is_deleted = 1 WHERE is_deleted = 0 AND policy_id = " & policy_id
                
        str_sql = "SELECT id security_id FROM " & load.sources.policy_binders_table & " WHERE is_deleted = 0 AND policy_id = " & policy_id
        
        Set col_binders = New Collection
        Set rs = utilities.create_adodb_rs(load.conn, str_sql)
        With rs
            .Open
            Do Until .EOF
                col_binders.Add rs!security_id.value
                .MoveNext
            Loop
            .Close
        End With
        Set rs = Nothing
        
        load.conn.BeginTrans
        tx_started = True
            
            load.conn.Execute str_sql_binders
            
            For Each security_id In col_binders
                With log_object
                    .data_set = load.sources.policy_binders_table
                    .policy_id = policy_id
                    .record_id = security_id
                    .security_id = security_id
                End With
                utilities.log_change log_object
            Next security_id
        
        load.conn.CommitTrans
        tx_started = False
    
    Next policy_id
    
    Forms(Me.Name).Visible = False
    
outro:
    Set col_policies = Nothing
    load.call_stack = ""
    load.event_id = ""
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "me!deal_id = " & Me!deal_id, "", True
    Resume Next
End Sub

Private Sub cmdOpenFolders_Click()
    load.call_stack = "deal_details_f.cmdOpenFolders"
    load.check_conn_and_variables
    
    With load.system_info.texts
        open_forms.working_on_it_f .opening_deal_folder_heading, .opening_deal_folder_details, 4000
    End With
    
    Paths.open_deal_folder Me!deal_id
    
outro:
    load.call_stack = ""
    open_forms.working_on_it_f__close
    Exit Sub
    
End Sub
Private Sub cmdSearchNewDeal_Click()
    Dim proc_name As String
    proc_name = "deal_details_f.cmdSearchNewDeal"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    load.check_conn_and_variables
    
    Forms("MenuMainF").SetFocus
    
    Dim str_sql As String
    str_sql = "INSERT INTO " & load.sources.log_activity_table _
    & " (activity, activity_source, user_name, app_name, app_continent, user_ip) VALUES " _
    & "('Clicked search othe deal'" _
    & ", '" & proc_name & "'" _
    & ", '" & Environ("username") & "'" _
    & ", ' " & load.system_info.app_name & "'" _
    & ", '" & load.system_info.app_continent & "'" _
    & ", '" & load.system_info.ip_current & "')"
    conn.Execute str_sql
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub cmdCreateFolder_Click()
    Const proc_name As String = "deal_details_f.cmdCreateFolder"
    load.call_stack = proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    Dim str_milestone As String
    
    'check if mandatory fields are populated
    Dim err_object As cls_err_object
    Dim form_field As cls_field
    Dim missing_fields As New Collection
    Dim obj_deal As New cls_deal
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    Set missing_fields = mandatory_data.check_data(Me!deal_id, load.mandatory_tests.folder_check)
    If missing_fields.Count > 0 Then
        For Each form_field In missing_fields
            With form_field
                Me.Controls(.field_name).BackColor = load.colors.yellow
            End With
        Next form_field
        Me.Controls(missing_fields(1).field_name).SetFocus
        GoTo outro
    End If
    
    'check if status is valid
    If load.deal_details.is_init = False Then
        load.deal_details.init Me!deal_id
    End If
    If Me!deal_status_id = "6" Or Me!deal_status_id = "436" Or Me!deal_status_id = "7" Or Me!deal_status_id = "8" Or Me!deal_status_id = "9" Then
        Me.Controls(load.deal_details.txt_deal_status_id.field_name).BackColor = load.colors.yellow
        MsgBox "Status cannot be 'Signed', 'Closed', 'Lost', 'Cancelled' or 'Declined' when creating folders.", , "Unacceptable deal status"
        GoTo outro
    End If
    
    With load.system_info.texts
        open_forms.working_on_it_f .creating_folder_heading, .creating_folder_details, 6000
    End With
    
    str_sql = "SELECT deal_id, deal_name, spa_law, MessageBoard, is_folder_created FROM " & load.sources.deals_table & " WHERE deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        With obj_deal
            .deal_id = rs!deal_id
            .deal_name = rs!deal_name.value
            .is_folder_created = rs!is_folder_created.value
            .message_board = rs!messageboard.value
            .spa_law = rs!spa_law.value
        End With
    rs.Close
    Set rs = Nothing
    
    str_milestone = "Before getting folder_creation_feedback"
    Dim folder_creation_feedback As Integer
    If load.system_info.app_continent = load.system_info.continents.americas Then
        folder_creation_feedback = load.deal_details.folders_create_folder_us(obj_deal)
    Else
        folder_creation_feedback = load.deal_details.folders_create_folder_eurasia(obj_deal)
    End If
    
    str_milestone = "Before logging folder creation failure"
    If folder_creation_feedback = -1 Then
        Set err_object = New cls_err_object
        With err_object
            .routine_name = "deal_details.cmdCreateFolder"
            .milestone = str_milestone
            .stella_error_text = "Stella failed to create folder via load.deal_details.create_folder(_us)"
            .params = "deal_id = " & Me!deal_id
            .system_error_code = Err.Number
            .system_error_text = Err.Description
            .send_error err_object
        End With
    End If
    
    imanage.create_workspace_caller Me!deal_id, current_uw.email
    
outro:
    open_forms.working_on_it_f__close
    load.call_stack = ""
    Exit Sub
    
err_handler:
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = str_milestone
        .params = "deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
    
End Sub
Private Sub cmd_open_cm_Click()
    Dim proc_name As String
    proc_name = "deal_details_f.cmd_open_cm"
    load.call_stack = proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    Dim feedback
    Dim str_milestone As String
    
    'check that mandatory fields are populated
    str_milestone = "checking mandatory fields"
    
    feedback = load.deal_details.are_required_field_to_start_cm_populated(Me!deal_id)
    If feedback = False Then
        load.deal_details.paint_deal_details
        GoTo outro
    ElseIf feedback = "-1" Then
        MsgBox load.system_info.error_instruction, , "CM could not start"
        str_milestone = "Tried to launch CM, but got -1 from load.deal_details.are_required_field_to_start_cm_populated(Me!deal_id)"
        GoTo err_handler
        GoTo outro
    End If
    
    load.check_secondary_access_app
    
    With load.system_info.texts
        open_forms.working_on_it_f .opening_cm_heading, .opening_cm_details, 4000
    End With
    
    Central.open_external_resource_app "cm_uw.accdb", True, load.current_uw.cm_address
    
    With load.secondary_access_app
        .Visible = True
        str_milestone = "Trying to launch the CM_3"
        .Run "cm_uw.load_cm", Me!deal_id.value, load.system_info.app_continent
    End With
    Set load.secondary_access_app = Nothing

outro:
    open_forms.working_on_it_f__close
    load.call_stack = ""
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "deal_details_f.cmd_open_cm.Click"
        .milestone = str_milestone
        .params = "me!deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Sub
Private Sub cmd_policies_Click()
    load.check_conn_and_variables
    open_forms.policies_f Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub broker_person_AfterUpdate()
    utilities.update_data "broker_person", Me!broker_person, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub Comments_AfterUpdate()
    utilities.update_data "comments", Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub CounselInvoiceReceived_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub deal_currency_AfterUpdate()
    load.call_stack = "deal.details_f.deal_currency_AfterUpdate()"
    load.check_conn_and_variables
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Me.SetFocus
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "deal_currency"
        .new_value = Me!deal_currency
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
    Central.update_currencies Me!deal_id
    
outro:
    load.call_stack = ""
    open_forms.working_on_it_f__close
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .call_stack = load.call_stack
        .routine_name = "deal_details_f.deal_currency.AfterUpdate"
        .milestone = "n/a"
        .params = "me!deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
    
End Sub
Private Sub deal_name_afterupdate()
    load.call_stack = "deal_details_f.deal_name_afterupdate"
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    load.check_conn_and_variables
    Dim str_milestone As String
    
    str_milestone = "before opening 'working on it'"
    
    With load.system_info.texts
        open_forms.working_on_it_f .updating_deal_name_heading, .updating_deal_name_details, 4000
    End With
    
    Dim deal_name As String, checked_deal_name As String
    deal_name = Me!deal_name
    
    'Make sure illegal character is not in project name
    str_milestone = "Before checking illegal chars"
    checked_deal_name = utilities.remove_illegal_characters(deal_name)
    If checked_deal_name <> deal_name Then
        MsgBox "Hi! Very sorry, but you have used illegal characters in the Risk Name. These have been removed.", , "illegal characters"
    End If
    Me!deal_name = checked_deal_name
    
    'save deal name
    str_milestone = "Before saving new name"
    utilities.update_data "deal_name", Me!deal_name, Me!deal_id
    
    'log change
    str_milestone = "before logging change"
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "deal_name"
        .new_value = Me!deal_name
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
    'check if deal with identical deal name has been created today (to avoid duplicates)
    str_milestone = "Before duplicate check"
    
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "SELECT deal_id, creating_uw FROM " & load.sources.duplicate_check_view _
    & " WHERE deal_name = '" & checked_deal_name & "' AND create_date >= '" & utilities.generate_sql_date(Date) & "'"
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        If CLng(.RecordCount) > 1 Then
            MsgBox "A deal with this exact name was created earlier today by " & !creating_uw & ". It might be a coincidence, or a duplicate. Please check." & vbNewLine & vbNewLine _
            & "Delete this (or the other risk) if there is a duplicate. The deal folder must be manually deleted.", , "Duplicate warning"
        End If
        .Close
    End With
    Set rs = Nothing
    
    '27 March 2025, CK: Ideally, we want iManage to be created and/or updated when changing the deal name. However, it takes too long, so need to find a different approach.
'    str_milestone = "Before iManage thing"
'    imanage.update_workspace_caller deal_id, load.current_uw.user_name & "@rpgroup.com"

outro:
    load.call_stack = ""
    open_forms.working_on_it_f__close
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "deal_details_f.deal_name.AfterUpdate"
        .milestone = str_milestone
        .params = "me!deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Private Sub internal_approver_binding_id_AfterUpdate()
    utilities.update_data "internal_approver_at_signing", Screen.ActiveControl, Me!deal_id
    
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "SELECT uw_id, uw_name FROM " & sources.underwriters_view & " WHERE uw_id = " & Me!internal_approver_binding_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        Me!internal_approver_binding_full_name = rs!uw_name
    rs.Close
    Set rs = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub internal_approver_quote_id_AfterUpdate()
    utilities.update_data "internal_approver", Screen.ActiveControl, Me!deal_id
    Dim str_sql As String, rs As ADODB.Recordset
    str_sql = "SELECT uw_id, uw_name FROM " & sources.underwriters_view & " WHERE uw_id = " & Me!internal_approver_quote_id
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        Me!internal_approver_quote_full_name = rs!uw_name
    rs.Close: Set rs = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub is_test_deal_id_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "is_test_deal_id"
        Me.SetFocus
        .new_value = Screen.ActiveControl
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub deal_tag_1_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_2_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_3_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_4_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_5_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_6_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_7_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_8_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_9_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_10_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_11_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_12_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_13_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_14_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_15_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_16_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_17_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_18_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_19_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_20_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_21_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_22_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_23_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_24_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_25_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_26_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_27_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_28_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_29_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_30_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_31_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_32_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_33_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_34_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_35_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_36_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_37_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_38_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_39_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_40_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_41_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_42_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_43_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_44_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_45_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_46_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_47_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_48_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_49_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub
Private Sub deal_tag_50_Click()
    open_forms.deal_tags_f Me!deal_id
End Sub

Private Sub max_limit_quoted_AfterUpdate()
    With load.deal_details.txt_max_limit_quoted
        utilities.update_data .field_name_in_recordset, Me.Controls(.field_name), Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub stage_id_AfterUpdate()
    With load.deal_details.txt_stage_id
        utilities.update_data .field_name_in_recordset, Me.Controls(.field_name), Me!deal_id
    End With
End Sub

Private Sub us_producer_no_AfterUpdate()
    With load.deal_details.txt_us_producer_no
        utilities.update_data .field_name_in_recordset, Me.Controls(.field_name), Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub uw_financial_advisor_id_AfterUpdate()
    load.check_conn_and_variables
    With load.deal_details
        If .is_init = False Then .init Me!deal_id
        With .txt_uw_financial_advisor_id
            utilities.update_data .field_name_in_recordset, Forms(load.deal_details.form_name).Controls(.field_name), Me!deal_id, False
        End With
    End With
    
outro:
    load.call_stack = ""
    Exit Sub

End Sub

Private Sub uw_law_firm_id_AfterUpdate()
    load.check_conn_and_variables
    utilities.update_data "uw_law_firm_id", Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub nbi_prepper_id_AfterUpdate()
    load.check_conn_and_variables
    
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    Me.SetFocus
    
    str_sql = "UPDATE " & sources.deals_table _
    & " SET " & load.deal_details.txt_nbi_prepper_id.field_name_in_table & " = " & Screen.ActiveControl _
    & " WHERE deal_id = " & Me!deal_id
    
    conn.Execute str_sql
    
    str_sql = "SELECT uw_id, uw_name FROM " & sources.underwriters_view & " WHERE uw_id = " & Me!nbi_prepper_id
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        Me!nbi_prepper_full_name = rs!uw_name
    rs.Close: Set rs = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub retention_dropped_AfterUpdate()
    With Me!retention_dropped
        utilities.update_data "drop_end", .value, Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub reviewer_notes_AfterUpdate()
    utilities.update_data "submission_notes", Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub risk_feel_id_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub ev_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id, False
        
    Central.update_currencies Me!deal_id
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "ev"
        Me.SetFocus
        .new_value = Me!ev.value
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub analyst_id_AfterUpdate()
    utilities.update_data load.deal_details.txt_analyst_id.field_name, Screen.ActiveControl, Me!deal_id
    
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    
    str_sql = "SELECT uw_id, uw_name FROM " & sources.underwriters_view & " WHERE uw_id = " & Me!analyst_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        Me!analyst_full_name = rs!uw_name
    rs.Close
    Set rs = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub inception_date_AfterUpdate()
    Dim proc_name As String
    proc_name = Me.Name & "." & Screen.ActiveControl.Name
    load.call_stack = proc_name

    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim new_data
    If Me!inception_date = "" Or IsNull(Me!inception_date) = True Then
        new_data = "NULL"
    Else
        new_data = utilities.generate_sql_date(Me!inception_date)
    End If
    
    utilities.update_data "inception_date", new_data, Me!deal_id
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "inception_date"
        .new_value = new_data
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
    Central.update_currencies Me!deal_id

outro:
    open_forms.working_on_it_f__close
    load.call_stack = ""
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "deal_details_f.inception_date.AfterUpdate"
        .milestone = "n/a"
        .params = "me!deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
    
End Sub
Private Sub is_underwritten_AfterUpdate()
    load.check_conn_and_variables
    Me.SetFocus
    conn.Execute "UPDATE " & sources.deals_table & " SET " & Screen.ActiveControl.Name & " = " & Screen.ActiveControl & "  WHERE deal_id = " & Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub policy_period_in_months_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "policy_period_in_months"
        .new_value = Me!policy_period_in_months
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub primary_insurer_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub primary_or_xs_id_AfterUpdate()
    load.check_conn_and_variables
    Dim new_value As Long
    new_value = Screen.ActiveControl
    Me.SetFocus
    conn.Execute "UPDATE " & sources.deals_table & " SET primary_or_xs_id = " & new_value & " WHERE deal_id = " & Me!deal_id
    Me.Controls(Screen.ActiveControl.Name).BackColor = colors.white
    
    'delete when unification is done
    If load.system_info.app_continent = load.system_info.continents.americas Then
        conn.Execute "UPDATE " & sources.deals_table & " SET primary_or_xs = " & new_value & " WHERE deal_id = " & Me!deal_id
    End If
    'dend delete
    
    '25 April 2025, CK: To avoid uws forgetting to set UW fees to 0 when deal is xs only
    With load.deal_details.txt_uw_fee
        If Me.Controls(.field_name) > 0 And new_value = menu_list.xs_only Then
            With Me.Controls(.field_name)
                .BackColor = colors.yellow
                .SetFocus
            End With
        End If
    End With
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "primary_or_xs_id"
        .new_value = Me!primary_or_xs_id
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub primary_uw_AfterUpdate()
    If IsNull(Me!primary_uw) = True Then Exit Sub
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    Dim str_sql As String, rs As ADODB.Recordset
    str_sql = "SELECT uw_id, uw_name FROM " & sources.underwriters_view & " WHERE uw_id = " & Me!primary_uw
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        Me!primary_uw_full_name = rs!uw_name
    rs.Close: Set rs = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub program_limit_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "program_limit"
        .new_value = Me!program_limit
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub quote_due_date_AfterUpdate()
    load.check_conn_and_variables
    Dim str_sql As String
    
    str_sql = "UPDATE " & load.sources.deals_table & " SET quote_due_date = '" & utilities.generate_sql_date(Screen.ActiveControl) & "' WHERE deal_id = " & Me!deal_id
    conn.Execute str_sql
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub risk_type_id_AfterUpdate()
    load.check_conn_and_variables
    conn.Execute "UPDATE " & sources.deals_table & " SET risk_type_id = " & Me!risk_type_id & " WHERE deal_id = " & Me!deal_id
    Me.Controls(Screen.ActiveControl.Name).BackColor = colors.white
    
    'delete when unfication is done
    If load.system_info.app_continent = load.system_info.continents.americas Then
        conn.Execute "UPDATE deals_t SET risk_type = " & Me!risk_type_id & " WHERE deal_id = " & Me!deal_id
    End If
    
    Me!risk_type_id.BackColor = colors.white
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "risk_type_id"
        .new_value = Me!risk_type_id
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub rr_done_AfterUpdate()
    load.check_conn_and_variables
    conn.Execute "UPDATE " & sources.deals_table & " SET rr_done = " & Me!rr_done & "  WHERE deal_id = " & Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub sanction_checks_AfterUpdate()
    With Me!sanction_checks
        utilities.update_data "sanction_checks_done", .value, Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub secondary_uw_AfterUpdate()
    utilities.update_data "secondary_uw", Screen.ActiveControl, Me!deal_id
    Dim str_sql As String, rs As ADODB.Recordset
    str_sql = "SELECT uw_id, uw_name FROM " & sources.underwriters_view & " WHERE uw_id = " & Me!secondary_uw
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        Me!secondary_uw_full_name = rs!uw_name
    rs.Close: Set rs = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub seller_business_name_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub SellerLegalFirm_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub signing_premium_amount_AfterUpdate()
    utilities.update_data "signing_invoice_amount", Screen.ActiveControl, Me!deal_id
    Me!closing_premium_amount = Me!total_premium - Me!signing_premium_amount
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "signing_premium_amount"
        .new_value = Me!signing_premium_amount
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub signing_premium_due_date_AfterUpdate()
    utilities.update_data "signing_premium_due_date", utilities.generate_sql_date(Screen.ActiveControl), Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub signing_premium_received_date_AfterUpdate()
    utilities.update_data "signing_premium_received_date", utilities.generate_sql_date(Screen.ActiveControl), Me!deal_id
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub spa_signing_date_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, utilities.generate_sql_date(Screen.ActiveControl), Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub spa_law_afterupdate()
    load.check_conn_and_variables
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    utilities.update_data "spa_law", Me!spa_law, Me!deal_id
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "spa_law"
        .new_value = Me!spa_law
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
    'update RPEntity (also referred to as budget_home)
    Central.spa_law_was_updated Me!deal_id
    Central.populate_deal_details_f Me!deal_id
    
    'update currencies
    Central.update_currencies Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "deal_details_f.spa_law.AfterUpdate"
        .milestone = "n/a"
        .params = "me!deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
End Sub

Private Sub submission_limits_AfterUpdate()
    utilities.update_data "submission_limits", Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub surplus_broker_firm_AfterUpdate()
    With Me!surplus_broker_firm
        utilities.update_data .Name, .value, Me!deal_id
    End With
    
    'log change
    load.check_conn_and_variables
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = Me!surplus_broker_firm.Name
        .new_value = Me!target_business_name.value
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub surplus_broker_license_no_AfterUpdate()
    load.check_conn_and_variables
    utilities.update_data "surplus_broker_license_no", Me!surplus_broker_license_no, Me!deal_id
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = Me!surplus_broker_license_no.Name
        .new_value = Me!surplus_broker_license_no.value
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub surplus_broker_person_AfterUpdate()
    With Me!surplus_broker_person
        utilities.update_data .Name, .value, Me!deal_id
    End With
    
    'log change
    load.check_conn_and_variables
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = sources.deals_table
        .field_name = Me!surplus_broker_person.Name
        .new_value = Me!surplus_broker_person.value
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub target_business_name_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
    'log change
    load.check_conn_and_variables
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = sources.deals_table
        .field_name = "target_business_name"
        .new_value = Me!target_business_name
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub target_country_AfterUpdate()
    Dim input_value As Long
    input_value = Me!target_country
    utilities.update_data "TargetDomicile", input_value, Me!deal_id
    
    'delete once unification is done
    If load.system_info.app_continent = load.system_info.continents.americas Then
        conn.Execute "UPDATE deals_t SET target_registered_jurisdiction_id = " & input_value & " WHERE deal_id =  " & Me!deal_id
    End If
    'delete end
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub retention_AfterUpdate()
    load.check_conn_and_variables
    utilities.update_data "retention", Me!retention, Me!deal_id
    Central.limit_premium_and_summary_central Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub target_desc_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub target_legal_name_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub target_main_jurisdiction_id_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
    'log change
    load.check_conn_and_variables
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = sources.deals_table
        .field_name = "target_main_jurisdiction_id"
        .new_value = Me!target_main_jurisdiction_id
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Public Sub target_naic_code_AfterUpdate()
    load.check_conn_and_variables
    With Me!target_naic_code
        If IsNumeric(.value) = False Or Len(.value) > 6 Or Len(.value) < 4 Then
            .BackColor = load.colors.yellow
            MsgBox "Target NAIC code must be four, five or six digits only.", , "Invalid NAIC code"
            GoTo outro
        End If
        If Len(.value) = 4 Then .value = .value * 10
        If Len(.value) = 5 Then .value = .value * 10
        utilities.update_data .Name, .value, Me!deal_id
    End With

outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub target_sic_code_AfterUpdate()
    load.check_conn_and_variables
    With Me!target_sic_code
        If IsNumeric(.value) = False Or Len(.value) <> 4 Then
            .BackColor = load.colors.yellow
            MsgBox "Target SIC code must be four digits only.", , "Invalid SIC code"
            GoTo outro
        End If
        utilities.update_data .Name, .value, Me!deal_id
    End With

outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub target_sub_sector_id_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub target_super_sector_id_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    load.deal_details.txt_target_super_sector_id.field_bg_color = load.colors.white
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = sources.deals_table
        .field_name = "target_super_sector_id"
        .new_value = Me!target_super_sector_id
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
    'populate target_sub_sector_id based on update of target_super_sector_id
    Dim str_sql As String, rs As ADODB.Recordset
    Dim target_super_setor_id As Long
    str_sql = "SELECT deal_id, target_super_sector_id FROM " & sources.deals_table & " WHERE deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        If IsNull(rs!target_super_sector_id) = True Then
            GoTo outro
        Else
            target_super_setor_id = rs!target_super_sector_id.value
        End If
    rs.Close
    
    'remove existing items and add new ones
    With Me!target_sub_sector_id
        .value = load.sub_sector_all
        Do While .ListCount > 0
            .RemoveItem (0)
        Loop
        str_sql = "SELECT sector_id id, sector_name menu_item FROM " & sources.sectors_table & " WHERE parent_sector_id = " & target_super_sector_id & " ORDER BY sector_name"
        Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
            Do While rs.EOF = False
                .AddItem rs!id & ";'" & rs!menu_item & "'"
                rs.MoveNext
            Loop
        rs.Close
    End With
    
    'reset target_sub_sector_id
    str_sql = "UPDATE " & load.sources.deals_table & " SET target_sub_sector_id = NULL WHERE deal_id = " & deal_id
    conn.Execute str_sql
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    load.call_stack = ""
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "deal_details_f.target_super_sector_id.AfterUpdate"
        .milestone = "n/a"
        .params = "str_sql = " & str_sql & ", deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Private Sub target_super_sector_id_GotFocus()
    load.check_conn_and_variables
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    Dim str_tool_tip As String
    
    str_tool_tip = "When selecting the super-sector and sub-sector, use your best judgement. If in doubt, consider what is most important to" _
        & " the target group (i.e. the primary commercial driver of the target business)." _
        & "<br><br>Often there is no perfect, or obvious, choice, and in those cases, you should choose what you think is right."

    load.check_conn_and_variables
    With load.deal_details
        If .is_init = False Then .init deal_id
        With .help_text
            .field_value = str_tool_tip
            .field_top = load.deal_details.txt_target_super_sector_id.field_top
            .field_left = load.deal_details.txt_target_super_sector_id.field_left + load.deal_details.txt_target_super_sector_id.field_width + utilities.twips_converter(4, "cm")
            .field_height = 2500
            .field_width = 5000
            .field_visible = True
            .field_bg_color = load.colors.very_light_green
        End With
    End With
    
    Dim col_controls As New Collection
    Set col_controls = Nothing
    col_controls.Add load.deal_details.help_text
    utilities.paint_control load.deal_details.form_name, col_controls
    
outro:
    load.call_stack = ""
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "deal_details_f.target_super_sector_id.GotFocus"
        .milestone = "n/a"
        .params = "deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Private Sub target_super_sector_id_LostFocus()
    load.check_conn_and_variables
    With load.deal_details
        If .is_init = False Then .init deal_id
        With .help_text
            .field_visible = False
        End With
    End With
    
    Dim col_controls As New Collection
    Set col_controls = Nothing
    col_controls.Add load.deal_details.help_text
    utilities.paint_control load.deal_details.form_name, col_controls
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub ultimate_buyer_AfterUpdate()
    utilities.update_data "UltimateBuyer", Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub UltimateSeller_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub uw_fee_AfterUpdate()
    load.check_conn_and_variables
    Me.uw_fee.BackColor = load.colors.white
    Me.SetFocus
    conn.Execute "UPDATE " & sources.deals_table & " SET uw_fee_amount = " & Screen.ActiveControl & "  WHERE deal_id = " & Me!deal_id
    load.deal_details.txt_uw_fee.field_bg_color = load.colors.white
    load.deal_details.paint_financial_info_section Me!deal_id
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "uw_fee"
        .new_value = Me!uw_fee
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub uw_fee_due_date_AfterUpdate()
    With Me!uw_fee_due_date
        utilities.update_data .Name, utilities.generate_sql_date(.value), Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub uw_fee_received_date_AfterUpdate()
    With Me!uw_fee_received_date
        utilities.update_data .Name, utilities.generate_sql_date(.value), Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub UwCounselPerson1_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub UwCounselPerson2_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub vat_home_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id, True
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub VDRPassword_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = sources.deals_table
        .field_name = "VDRPassword"
        .new_value = Me!VDRPassword
        .record_id = Me!deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub VDRReceived_AfterUpdate()
    utilities.update_data Screen.ActiveControl.Name, Screen.ActiveControl, Me!deal_id
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub was_quoted_id_AfterUpdate()
    Dim str_sql As String
    str_sql = "UPDATE " & load.sources.deals_table & " SET was_quoted_id = " & Me!was_quoted_id & " WHERE deal_id = " & Me!deal_id
    conn.Execute str_sql
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub was_quoted_id_GotFocus()
    load.check_conn_and_variables
    
    Dim col_update_fields As New Collection
    Dim str_tool_tip As String
    str_tool_tip = "Quoted is automatically changed by Stella. Only change if the data is wrong."
    
    With load.deal_details
        If .is_init = False Then .init deal_id
        With .help_text
            .field_value = str_tool_tip
            .field_top = load.deal_details.txt_was_quoted_id.field_top
            .field_left = load.deal_details.txt_was_quoted_id.field_left + load.deal_details.txt_was_quoted_id.field_width
            .field_height = 1500
            .field_width = 2000
            .field_visible = True
            .field_bg_color = load.colors.very_light_green
        End With
        Set col_update_fields = Nothing
        col_update_fields.Add .help_text
        utilities.paint_control load.deal_details.form_name, col_update_fields
    End With
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub was_quoted_id_LostFocus()
    Dim col_update_fields As New Collection
    
    With load.deal_details
        If .is_init = False Then .init deal_id
        With .help_text
            .field_visible = False
        End With
        Set col_update_fields = Nothing
        col_update_fields.Add .help_text
        utilities.paint_control load.deal_details.form_name, col_update_fields
    End With
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub


