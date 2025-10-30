VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_form_weekly_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public cmd_change_analyst As New cls_field
Public cmd_change_budget_home As New cls_field
Public cmd_change_is_repeat_buyer As New cls_field
Public cmd_change_nbi_prepper As New cls_field
Public cmd_change_primary_uw As New cls_field
Public cmd_change_requote_info As New cls_field
Public cmd_change_submission_notes As New cls_field

Public col_all_controls As New Collection
Public col_header_labels As New Collection
Public col_header_labels_all As New Collection
Public col_header_labels_us As New Collection
Public col_text_fields As New Collection
Public col_text_fields_all As New Collection
Public col_text_fields_us As New Collection
Public col_cmds_change As New Collection
Public col_cmds_change_us As New Collection

Public form_name As String

Public header_cmd_foreign_deals As New cls_field
Public header_cmd_foreign_deals_frame As New cls_field

Public header_lbl_broker As New cls_field
Public header_lbl_budget_home As New cls_field
Public header_lbl_buyer_business_name As New cls_field
Public header_lbl_buyer_law_firm_1 As New cls_field
Public header_lbl_change_roles As New cls_field
Public header_lbl_deadline As New cls_field
Public header_lbl_deal_id As New cls_field
Public header_lbl_deal_info As New cls_field
Public header_lbl_deal_name As New cls_field
Public header_lbl_deal_status As New cls_field
Public header_lbl_ev As New cls_field
Public header_lbl_is_repeat_buyer As New cls_field
Public header_lbl_max_limit_quoted As New cls_field
Public header_lbl_re_quote_info As New cls_field
Public header_lbl_risk_type_major_id As New cls_field
Public header_lbl_roles As New cls_field
Public header_lbl_submission_limits As New cls_field
Public header_lbl_submission_notes As New cls_field
Public header_lbl_target_business_name As New cls_field
Public header_lbl_target_description As New cls_field
Public header_selector_uws As New cls_field

Public is_init As Boolean

Public make_all_fields_visible_for_debugging As Boolean

Public shall_be_refreshed As Boolean

Public tog_foreign_deals As New cls_toggle_button

Public txt_broker As New cls_field
Public txt_budget_home As New cls_field
Public txt_budget_home_change As New cls_field
Public txt_buyer_business_name As New cls_field
Public txt_buyer_law_firm_1 As New cls_field
Public txt_deadline As New cls_field
Public txt_deal_id As New cls_field
Public txt_deal_info As New cls_field
Public txt_deal_name As New cls_field
Public txt_deal_status As New cls_field
Public txt_deal_status_change As New cls_field
Public txt_ev_deal As New cls_field
Public txt_is_repeat_buyer As New cls_field
Public txt_max_limit_quoted As New cls_field
Public txt_nbi_prepper_change As New cls_field
Public txt_re_quote_info As New cls_field
Public txt_risk_type_major_id As New cls_field
Public txt_roles As New cls_field
Public txt_row_bg As New cls_field
Public txt_submission_limits As New cls_field
Public txt_submission_notes As New cls_field
Public txt_target_business_name As New cls_field
Public txt_target_description As New cls_field
Public txt_week_no As New cls_field

Public Sub change_deal_status(ByVal deal_id As Long _
, ByVal current_row As Long _
, ByVal new_deal_status_id As Long)
    Dim proc_name As String
    proc_name = "weekly_view_f.cls_form_weekly_view.change_deal_status"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim deal_status_id As Long
    Dim str_sql As String
    
    If new_deal_status_id = deal_statuses.signed Or new_deal_status_id = deal_statuses.closed Then
        MsgBox "Deals can only be set to 'signed' or 'closed' from the Deal Details View."
        GoTo outro
    End If
    
    deal_status_id = new_deal_status_id
    str_sql = "UPDATE " & load.sources.deals_table & " SET aim_for_xs = 'no' WHERE deal_id = " & deal_id
    
    If new_deal_status_id = -2 Then
        deal_status_id = deal_statuses.declined
        str_sql = "UPDATE " & load.sources.deals_table & " SET aim_for_xs = 'yes' WHERE deal_id = " & deal_id
    End If
    
    conn.Execute str_sql
    
    Central.deal_status_was_changed deal_id, deal_status_id, False, Central.get_current_risk_status(deal_id)
    
    fix_rs.weekly_view_f 0, Forms(Me.form_name).RecordSource
    
    Forms(Me.form_name).SelTop = current_row
    
outro:
    Forms(Me.form_name).Controls(Me.txt_deal_status_change.field_name) = ""
    Forms(Me.form_name).Controls(Me.txt_deal_name.field_name).SetFocus
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
Public Sub change_nbi_prepper(ByVal deal_id As Long _
, ByVal current_row As Long _
, ByVal nbi_prepper_id As Long)
    Dim proc_name As String
    proc_name = "weekly_view_f.cls_form_weekly_view.change_nbi_prepper"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String
    If nbi_prepper_id = -1 Then
        str_sql = "UPDATE " & load.sources.deals_table & " SET " & Me.txt_nbi_prepper_change.field_name_in_table & " = NULL WHERE deal_id = " & deal_id
    Else
        str_sql = "UPDATE " & load.sources.deals_table & " SET " & Me.txt_nbi_prepper_change.field_name_in_table & " = " & nbi_prepper_id & " WHERE deal_id = " & deal_id
    End If
    conn.Execute str_sql
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = load.sources.deals_table
        .field_name = "nbi_prepper_id"
        .new_value = nbi_prepper_id
        .record_id = deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
    fix_rs.weekly_view_f 0, Forms(Me.form_name).RecordSource
    
    Forms(Me.form_name).SelTop = current_row
    
outro:
    Forms(Me.form_name).Controls(Me.txt_nbi_prepper_change.field_name) = ""
    Forms(Me.form_name).Controls(Me.txt_deal_name.field_name).SetFocus
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
Public Sub init_collections()
    Const proc_name As String = "cls_form_weekly_view.init_collections"
    utilities.call_stack_add_item proc_name
    
    Dim col_col As New Collection
    Dim form_field As cls_field
    
    'EurAsia
    Set Me.col_header_labels = Nothing
    With Me.col_header_labels
        .Add Me.header_lbl_broker
        .Add Me.header_lbl_submission_limits
        .Add Me.header_lbl_deadline
        .Add Me.header_lbl_target_business_name
        .Add Me.header_lbl_deal_status
        .Add Me.header_lbl_roles
        .Add Me.header_lbl_change_roles
        .Add Me.header_lbl_budget_home
        .Add Me.header_lbl_submission_notes
        .Add Me.header_lbl_deal_info
        .Add Me.header_lbl_buyer_business_name
        .Add Me.header_lbl_deal_id
        .Add Me.header_lbl_is_repeat_buyer
        .Add Me.header_lbl_re_quote_info
    End With

    Set Me.col_text_fields = Nothing
    With Me.col_text_fields
        .Add Me.txt_broker
        .Add Me.txt_submission_limits
        .Add Me.txt_deadline
        .Add Me.txt_target_business_name
        .Add Me.txt_deal_status
        .Add Me.txt_roles
        .Add Me.txt_budget_home
        .Add Me.txt_submission_notes
        .Add Me.txt_deal_info
        .Add Me.txt_buyer_business_name
        .Add Me.txt_deal_id
        .Add Me.txt_is_repeat_buyer
        .Add Me.txt_re_quote_info
    End With
    
    Set Me.col_cmds_change = Nothing
    With Me.col_cmds_change
        .Add Me.cmd_change_analyst
        .Add Me.cmd_change_budget_home
        .Add Me.cmd_change_is_repeat_buyer
        .Add Me.cmd_change_nbi_prepper
        .Add Me.cmd_change_primary_uw
        .Add Me.cmd_change_requote_info
        .Add Me.cmd_change_submission_notes
        .Add Me.txt_deal_status_change
        .Add Me.txt_nbi_prepper_change
    End With
    
    'Americas
    Set Me.col_header_labels_us = Nothing
    With Me.col_header_labels_us
        .Add Me.header_lbl_deadline
        .Add Me.header_lbl_broker
        .Add Me.header_lbl_target_description
        .Add Me.header_lbl_target_business_name
        .Add Me.header_lbl_buyer_business_name
        .Add Me.header_lbl_max_limit_quoted
        .Add Me.header_lbl_buyer_law_firm_1
        .Add Me.header_lbl_roles
        .Add Me.header_lbl_change_roles
        .Add Me.header_lbl_deal_status
        .Add Me.header_lbl_re_quote_info
        .Add Me.header_lbl_submission_notes
        .Add Me.header_lbl_deal_id
        .Add Me.header_lbl_risk_type_major_id
        .Add Me.header_lbl_is_repeat_buyer
    End With
    
    Set Me.col_text_fields_us = Nothing
    With Me.col_text_fields_us
        .Add Me.txt_deadline
        .Add Me.txt_broker
        .Add Me.txt_target_description
        .Add Me.txt_target_business_name
        .Add Me.txt_buyer_business_name
        .Add Me.txt_max_limit_quoted
        .Add Me.txt_buyer_law_firm_1
        .Add Me.txt_roles
        .Add Me.txt_deal_status
        .Add Me.txt_re_quote_info
        .Add Me.txt_submission_notes
        .Add Me.txt_deal_id
        .Add Me.txt_risk_type_major_id
        .Add Me.txt_is_repeat_buyer
    End With
    
    Set Me.col_cmds_change_us = Nothing
    With Me.col_cmds_change_us
        .Add Me.cmd_change_analyst
        .Add Me.cmd_change_is_repeat_buyer
        .Add Me.cmd_change_nbi_prepper
        .Add Me.cmd_change_primary_uw
        .Add Me.cmd_change_requote_info
        .Add Me.cmd_change_submission_notes
        .Add Me.txt_deal_status_change
        .Add Me.txt_nbi_prepper_change
    End With
    
    'all items
    Set Me.col_header_labels_all = Nothing
    With Me.col_header_labels_all
        For Each form_field In Me.col_header_labels
            .Add form_field
        Next form_field
        For Each form_field In Me.col_header_labels_us
            .Add form_field
        Next form_field
    End With
    
    Set Me.col_text_fields_all = Nothing
    With Me.col_text_fields_all
        For Each form_field In Me.tog_foreign_deals.col_controls
            .Add form_field
        Next form_field
        For Each form_field In Me.col_text_fields
            .Add form_field
        Next form_field
        For Each form_field In Me.col_text_fields_us
            .Add form_field
        Next form_field
    End With
End Sub

Public Sub init()
    Dim proc_name As String
    proc_name = "cls_form_weekly_view.init"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Me.form_name = "weekly_view_f"
    Me.init_collections
    Me.init_txt_fields
    Me.init_lbls
    Me.init_cmds
    Me.init_header_selectors
        
    is_init = True

outro:
    Exit Sub
    
End Sub
Public Sub init_header_selectors()
    With Me.header_selector_uws
        .field_name = "header_selector_uws"
        .field_name_in_table = global_vars.data_fields.nbi_prepper.field_name_in_table
    End With
End Sub
Public Sub init_cmds()
    Dim form_field As cls_field
    
    For Each form_field In Me.col_cmds_change
        With form_field
            .field_height = utilities.twips_converter(0.2, "inch")
            .field_visible = True
        End With
    Next form_field
    
    With Me.cmd_change_analyst
        .field_left = Me.header_lbl_change_roles.field_left
        .field_name = "cmd_change_analyst"
        .field_top = Me.cmd_change_nbi_prepper.field_bottom
        .field_width = Me.header_lbl_change_roles.field_width
    End With
    With Me.cmd_change_budget_home
        .field_left = Me.txt_budget_home.field_left
        .field_name = "cmd_change_budget_home"
        .field_top = utilities.twips_converter(0.45, "inch")
        .field_width = Me.txt_budget_home.field_width
        .margin_external_left = utilities.twips_converter(0.8, "inch") 'catering for control boxes to the left
    End With
    With Me.cmd_change_is_repeat_buyer
        .field_left = Me.txt_buyer_business_name.field_left
        .field_name = "cmd_change_is_repeat_buyer"
        .field_top = utilities.twips_converter(0.65, "inch")
        .field_width = Me.txt_buyer_business_name.field_width
    End With
    With Me.cmd_change_nbi_prepper
        .field_left = Me.header_lbl_change_roles.field_left
        .field_name = "cmd_change_nbi_prepper"
        .field_top = 0
        .field_width = Me.header_lbl_change_roles.field_width
    End With
    With Me.cmd_change_primary_uw
        .field_left = Me.header_lbl_change_roles.field_left
        .field_name = "cmd_change_primary_uw"
        .field_top = Me.cmd_change_analyst.field_bottom
        .field_width = Me.header_lbl_change_roles.field_width
    End With
    With Me.cmd_change_requote_info
        .field_width = utilities.twips_converter(0.65, "inch")
        .field_left = Me.txt_re_quote_info.field_right - .field_width
        .field_name = "cmd_change_requote_info"
    End With
    With Me.cmd_change_submission_notes
        .field_width = utilities.twips_converter(0.65, "inch")
        .field_left = Me.txt_submission_notes.field_right - .field_width
        .field_name = "cmd_change_submission_notes"
    End With
    
    With Me.txt_deal_status_change
        .field_left = Me.txt_deal_status.field_left
        .field_name = "deal_status_change"
        .field_width = Me.txt_deal_status.field_width
        If load.system_info.app_continent = load.system_info.continents.americas Then
            .margin_external_left = Me.header_lbl_change_roles.field_width 'catering for control boxes to the left
        End If
    End With
    With Me.txt_budget_home_change
        .field_name_in_table = global_vars.data_fields.budget_home_id.field_name_in_table
        .field_left = Me.txt_deal_status.field_left
        .field_name = "budget_home_change"
        .field_top = utilities.twips_converter(0.65, "inch")
        .field_width = Me.txt_deal_status.field_width
        If load.system_info.app_continent = load.system_info.continents.americas Then
            .margin_external_left = Me.header_lbl_change_roles.field_width 'catering for control boxes to the left
        End If
    End With
    With Me.txt_nbi_prepper_change
        .field_name_in_table = global_vars.data_fields.nbi_prepper.field_name_in_table
        .field_left = Me.txt_roles.field_left
        .field_name = "nbi_prepper_change"
        .field_width = utilities.twips_converter(0.8, "inch")
    End With
    
    Me.tog_foreign_deals.init 100, 9500, "foreign_deals", "Include foreign deals with local UWs"
    
End Sub
Public Sub init_lbls()
    
    Dim form_field As cls_field
    
    For Each form_field In Me.col_header_labels
        With form_field
            .field_left = -1
            .field_visible = False
        End With
    Next form_field
    
    For Each form_field In Me.col_header_labels_us
        With form_field
            .field_left = -1
            .field_visible = False
        End With
    Next form_field
    
    With Me.header_lbl_broker
        .field_name = "lbl_broker_info"
        .field_width = Me.txt_broker.field_width
    End With
    With Me.header_lbl_budget_home
        .field_name = "lbl_budget_home"
        .field_width = Me.txt_budget_home.field_width
    End With
    With Me.header_lbl_buyer_business_name
        .field_name = "lbl_buyer_business_name"
        .field_width = Me.txt_buyer_business_name.field_width
    End With
    With Me.header_lbl_buyer_law_firm_1
        .field_name = "lbl_buyer_law_firm_1"
        .field_width = Me.txt_buyer_law_firm_1.field_width
    End With
    With Me.header_lbl_change_roles
        .field_name = "lbl_change_roles"
        .field_width = utilities.twips_converter(0.8, "inch")
    End With
    With Me.header_lbl_deadline
        .field_name = "lbl_deadline"
        .field_width = Me.txt_deadline.field_width
    End With
    With Me.header_lbl_deal_id
        .field_name = "lbl_deal_id"
        .field_width = Me.txt_deal_id.field_width
    End With
    With Me.header_lbl_deal_info
        .field_name = "lbl_deal_info"
        .field_width = Me.txt_deal_info.field_width
    End With
    With Me.header_lbl_deal_status
        .field_name = "lbl_deal_status"
        .field_width = Me.txt_deal_status.field_width
    End With
    With Me.header_lbl_ev
        .field_name = "lbl_ev"
        .field_width = Me.txt_ev_deal.field_width
        .field_left = Me.txt_ev_deal.field_left
    End With
    With Me.header_lbl_deal_name
        .field_name = "lbl_deal_name"
        .field_width = Me.txt_deal_name.field_width
    End With
    With Me.header_lbl_is_repeat_buyer
        .field_name = "header_lbl_is_repeat_buyer"
        .field_width = Me.txt_is_repeat_buyer.field_width
    End With
    With Me.header_lbl_max_limit_quoted
        .field_name = "header_lbl_max_limit_quoted"
        .field_width = Me.txt_max_limit_quoted.field_width
    End With
    With Me.header_lbl_re_quote_info
        .field_name = "lbl_re_quote_info"
        .field_width = Me.txt_re_quote_info.field_width
    End With
    With Me.header_lbl_risk_type_major_id
        .field_name = "lbl_risk_type_major_id"
        .field_width = Me.txt_risk_type_major_id.field_width
    End With
    With Me.header_lbl_roles
        .field_name = "lbl_roles"
        .field_width = Me.txt_roles.field_width
    End With
    With Me.header_lbl_submission_limits
        .field_name = "lbl_submission_limits"
        .field_width = Me.txt_submission_limits.field_width
    End With
    With Me.header_lbl_submission_notes
        .field_name = "lbl_submission_notes"
        .field_width = Me.txt_submission_notes.field_width
    End With
    With Me.header_lbl_target_business_name
        .field_name = "lbl_target_business_name"
        .field_width = Me.txt_target_business_name.field_width
    End With
    With Me.header_lbl_target_description
        .field_name = "lbl_target_description"
        .field_width = Me.txt_target_description.field_width
    End With
    
End Sub
Public Sub init_txt_fields()
    
    Dim form_field As cls_field
    
    For Each form_field In Me.col_text_fields
        With form_field
            .field_left = -1
            .field_visible = False
        End With
    Next form_field
    
    For Each form_field In Me.col_text_fields_us
    With form_field
            .field_left = -1
            .field_visible = False
        End With
    Next form_field
    
    With Me.txt_broker
        .field_name = "broker_info"
        .field_name_in_recordset = "broker_info"
        .field_top = 0
        .field_width = utilities.twips_converter(1, "inch")
    End With
    With Me.txt_budget_home
        .field_name = "budget_home"
        .field_name_in_recordset = "budget_home_spa_law"
        .field_name_in_table = "budget_home_id"
        .field_top = 0
        .field_width = utilities.twips_converter(1.2, "inch")
        .margin_external_left = utilities.twips_converter(0.8, "inch") 'catering for control boxes to the left
    End With
    With Me.txt_buyer_business_name
        .field_name = "buyer_business_name"
        .field_name_in_recordset = "buyer_business_name"
        .field_name_in_table = "buyer_business_name"
        .field_top = 0
        .field_width = utilities.twips_converter(1.2, "inch")
    End With
    With Me.txt_buyer_law_firm_1
        .field_name = "buyer_law_firm_1"
        .field_name_in_recordset = "buyer_law_firm_1"
        .field_name_in_table = "buyer_law_firm_1_id"
        .field_top = 0
        .field_width = utilities.twips_converter(1.2, "inch")
    End With
    With Me.txt_deadline
        .field_name = "nbi_deadline"
        .field_name_in_recordset = "nbi_deadline"
        .field_top = 0
        .field_width = utilities.twips_converter(1.2, "inch")
    End With
    With Me.txt_deal_id
        .field_name = "deal_id"
        .field_name_in_recordset = "deal_id"
        .field_top = 0
        .field_width = utilities.twips_converter(0.5, "inch")
    End With
    With Me.txt_deal_info
        .field_name = "deal_info"
        .field_name_in_recordset = "deal_info"
        .field_top = 0
        .field_width = utilities.twips_converter(3, "inch")
    End With
    With Me.txt_deal_name
        .field_name = "deal_name"
        .field_name_in_recordset = "deal_name_risk_type"
        .field_top = 0
        .field_width = utilities.twips_converter(1.2, "inch")
    End With
    With Me.txt_deal_status
        .field_name = "deal_status"
        .field_name_in_recordset = "deal_status_xs"
        If load.system_info.app_continent = load.system_info.continents.americas Then
            .field_name_in_recordset = "deal_status_us_xs"
            .margin_external_left = utilities.twips_converter(0.8, "inch") 'catering for control boxes to the left
        End If
        .field_top = 0
        .field_width = utilities.twips_converter(0.9, "inch")
    End With
    With Me.txt_ev_deal
        .field_left = utilities.twips_converter(1.9, "inch")
        .field_name = "ev_deal"
        .field_name_in_recordset = "ev_deal"
        .field_top = 0
        .field_width = utilities.twips_converter(0.8, "inch")
    End With
    With Me.txt_is_repeat_buyer
        .field_name = "is_repeat_buyer"
        .field_name_in_recordset = "is_repeat_buyer"
        .field_top = 0
        .field_width = utilities.twips_converter(0.5, "inch")
    End With
    With Me.txt_max_limit_quoted
        .field_name = "max_limit_quoted"
        .field_name_in_recordset = "max_limit_quoted"
        .field_top = 0
        .field_width = utilities.twips_converter(1, "inch")
    End With
    With Me.txt_re_quote_info
        .field_name = "re_quote_info"
        .field_name_in_recordset = "re_quote_info"
        .field_top = 0
        .field_width = utilities.twips_converter(1.2, "inch")
    End With
    With Me.txt_risk_type_major_id
        .field_name = "risk_type_major_id"
        .field_name_in_recordset = "risk_type_major_id"
        .field_top = 0
        .field_width = utilities.twips_converter(0.5, "inch")
    End With
    With Me.txt_roles
        .field_name = "roles"
        .field_name_in_recordset = "-1"
        .field_top = 0
        .field_width = utilities.twips_converter(0.8, "inch")
    End With
    With Me.txt_row_bg
        .field_name = "row_bg"
        .field_name_in_recordset = "-1"
        .field_top = 0
        .field_width = utilities.twips_converter(20, "inch")
    End With
    With Me.txt_submission_limits
        .field_name = "submission_limits"
        .field_name_in_recordset = "submission_limits"
        .field_top = 0
        .field_width = utilities.twips_converter(1.2, "inch")
    End With
    With Me.txt_submission_notes
        .field_name = "submission_notes"
        .field_name_in_recordset = "submission_notes"
        .field_top = 0
        .field_width = utilities.twips_converter(2.6, "inch")
    End With
    With Me.txt_target_business_name
        .field_name = "target_business_name"
        .field_name_in_recordset = "target_business_name"
        .field_top = 0
        .field_width = utilities.twips_converter(1.2, "inch")
    End With
    With Me.txt_target_description
        .field_name = "target_description"
        .field_name_in_recordset = "target_desc"
        .field_top = 0
        .field_width = utilities.twips_converter(2, "inch")
    End With
    With Me.txt_week_no
        .field_name = "week_no"
        .field_name_in_recordset = ""
        .field_top = 0
        .field_width = utilities.twips_converter(0.7, "inch")
    End With
End Sub

Public Sub paint_visible_only(ByVal input_col As Collection)
    Dim proc_name As String
    proc_name = "cls_form_weekly_view.paint_visible_only"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Forms(Me.form_name).Controls("week_no").SetFocus
    For Each form_field In input_col
        With Forms(Me.form_name).Controls(form_field.field_name)
            If .Visible = True Then
                .Visible = form_field.field_visible
            End If
        End With
    Next form_field

outro:
    If Not form_field Is Nothing Then Set form_field = Nothing
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
Public Sub paint(ByVal input_col As Collection)
    Dim proc_name As String
    proc_name = "cls_form_weekly_view.paint"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    If CurrentProject.AllForms(Me.form_name).IsLoaded = False Then
        GoTo outro
    End If
    
    Dim form_field As cls_field
    Dim form_field_prev As cls_field
    Dim margin_external_left As Long
    
    Set form_field_prev = Me.txt_ev_deal
    For Each form_field In input_col
        With Forms(Me.form_name).Controls(form_field.field_name)
            
            If form_field.field_bg_color <> -1 Then
                .BackColor = form_field.field_bg_color
            End If
            
            If form_field.field_caption <> "-1" Then
                .Caption = UCase(Left(form_field.field_caption, 1)) & Right(form_field.field_caption, Len(form_field.field_caption) - 1)
            End If
            If form_field.field_name_in_recordset <> "-1" Then
                .ControlSource = form_field.field_name_in_recordset
            End If
            
            'width needs to be set first because it affects the right property of the next control
            .Width = form_field.field_width
            margin_external_left = 0
            
            If form_field.margin_external_left <> -1 Then margin_external_left = form_field.margin_external_left
            If form_field.field_left = -1 Then
                form_field.field_left = form_field_prev.field_right
            End If
            
            .Left = form_field.field_left + margin_external_left
            .Visible = form_field.field_visible
            If form_field.field_top <> -1 Then
                .Top = form_field.field_top
            End If
            Set form_field_prev = form_field
        End With
    Next form_field
    
outro:
    If Not form_field Is Nothing Then Set form_field = Nothing
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

Public Sub quick_change(ByVal data_field_in_recordset As String _
, ByVal display_name As String _
, ByVal menu_list_source As String _
, ByVal deal_id As Long)

    Dim proc_name As String
    proc_name = "cls_form_master_deal_list.quick_change"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    load.check_conn_and_variables
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim rs As ADODB.Recordset
    Dim str_form As String
    Dim str_sql As String
    
    str_form = "deal_list__quick_change_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    
    'add values to form
    With Forms(str_form)
        .Caption = display_name
        !deal_id = deal_id
        !field_name = data_field_in_recordset
        
        'remove any current values
        Do While !field_value.ListCount > 0
            !field_value.RemoveItem (0)
        Loop
        
        'add new values
        str_sql = menu_list_source
        Set rs = utilities.create_adodb_rs(conn, str_sql)
        rs.Open
            rs.MoveFirst
            Do While rs.EOF = False
                !field_value.AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
        rs.Close
        Set rs = Nothing
        
        'move form
        Dim lng_down As Long
        Dim lng_right As Long
        lng_down = windows_apis.MouseY * 12
        lng_right = windows_apis.MouseX * 6
        .SetFocus
        DoCmd.MoveSize Right:=lng_right, Down:=lng_down, Width:=4000, Height:=1500
    
    End With
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "cls_form_master_deal_list.quick_change"
        .milestone = ""
        .params = "data_field_in_recordset = " & data_field_in_recordset & ", display_name = " & display_name & ", menu_list_source = " & menu_list_source
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
End Sub

Public Sub sort_all_fields_in_form_for_debugging()
    Dim proc_name As String
    proc_name = "cls_form_master_deal_list.sort_all_fields_in_form_for_debugging"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    'to sort the fields, set the form to design mode, and run this module
    
    Dim form_field As cls_field
    load.check_conn_and_variables
    Me.make_all_fields_visible_for_debugging = True
    load.is_debugging = True
    Me.init
    
    For Each form_field In Me.col_all_controls
        form_field.field_visible = True
    Next form_field

    Me.make_all_fields_visible_for_debugging = False
End Sub


