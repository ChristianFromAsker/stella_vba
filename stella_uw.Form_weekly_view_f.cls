VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_weekly_view_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub budget_home_change_afterupdate()
    Const proc_name As String = "weekly_view_f.budget_home_change.Click"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    With global_vars.interfaces.weekly_view_f
        .change_budget_home Me!deal_id, Me.SelTop, Me.budget_home_change
    End With
outro:
    load.call_stack = ""
    Exit Sub

err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "Me!deal_id = " & Me!deal_id, "", True
    Resume outro
End Sub

Private Sub cmd_change_analyst_Click()
    Dim proc_name As String
    proc_name = "weekly_view_f.cmd_change_analyst.Click"
    load.call_stack = proc_name
    
    load.check_conn_and_variables
    
    With load.deal_details
        If .is_init = False Then .init Me!deal_id
        load.master_deal_list.quick_change .txt_analyst_id.field_name, "Analyst", load.sources.menu_lists.uws, Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
End Sub

Private Sub cmd_change_is_repeat_buyer_Click()
    Dim proc_name As String
    proc_name = "weekly_view_f.cmd_change_is_repeat_buyer.Click"
    load.call_stack = proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    Dim deal_id As Long
    Dim str_sql As String
        
    deal_id = Me!deal_id
    
    With global_vars.interfaces.weekly_view_f
        If Me.Controls(.txt_is_repeat_buyer.field_name) = "yes" Then
            str_sql = "UPDATE " & load.sources.deals_table & " SET is_repeat_buyer = 'no' WHERE deal_id = " & deal_id
        Else
            str_sql = "UPDATE " & load.sources.deals_table & " SET is_repeat_buyer = 'yes' WHERE deal_id = " & deal_id
        End If
    End With
    conn.Execute str_sql
    
    fix_rs.weekly_view_f 0, Me.RecordSource
    
outro:
    load.call_stack = "" ' if initiating sub
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "me!deal_id = " & Me!deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Private Sub cmd_change_nbi_prepper_Click()
    Dim proc_name As String
    proc_name = "weekly_view_f.cmd_change_nbi_prepper.Click"
    load.call_stack = proc_name
    
    load.check_conn_and_variables
        
    ' CHECK IF CHANGE OF NBI PREPPER CAN SET DEAL STATUS TO NBN
    
    With load.deal_details
        If .is_init = False Then .init Me!deal_id
        load.master_deal_list.quick_change .txt_nbi_prepper_id.field_name, "NBI prepper", load.sources.menu_lists.uws, Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
End Sub

Private Sub cmd_change_primary_uw_Click()
    Dim proc_name As String
    proc_name = "weekly_view_f.cmd_change_primary_uw.Click"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    With load.deal_details
        If .is_init = False Then .init Me!deal_id
        load.master_deal_list.quick_change .txt_primary_uw_id.field_name, "Primary UW", load.sources.menu_lists.uws, Me!deal_id
    End With

outro:
    load.call_stack = ""
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

Private Sub cmd_change_requote_info_Click()
    Dim proc_name As String
    proc_name = "weekly_view_f.cmd_change_requote_info.Click"
    load.call_stack = proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    With global_vars.data_fields.requote_info
        global_vars.interfaces.quick_change_f.setup _
        Me!deal_id _
        , .field_name_in_table _
        , "Requote info" _
        , "txt" _
        , utilities.twips_converter(1.2, "inch") _
        , Me.Name _
        , load.sources.menu_lists.budget_homes _
        , "" _
        , Nz(Me.Controls("re_quote_info"), "")
    End With
    
outro:
    load.call_stack = ""
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

Private Sub cmd_change_submission_notes_Click()
    Dim proc_name As String
    proc_name = "weekly_view_f.cmd_change_submission_notes.Click"
    load.call_stack = proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    With global_vars.data_fields.submission_notes
        global_vars.interfaces.quick_change_f.setup _
        Me!deal_id _
        , .field_name_in_table _
        , "Submission notes" _
        , "txt" _
        , utilities.twips_converter(1.2, "inch") _
        , Me.Name _
        , load.sources.menu_lists.budget_homes _
        , "" _
        , Nz(Me.Controls("submission_notes"), "")
    End With
    
outro:
    load.call_stack = ""
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

Private Sub cmd_export_deals_Click()
    Dim proc_name As String
    proc_name = "weekly_view_f.cmd_export_deals.Click"
    load.call_stack = proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_milestone As String
    Dim str_sql As String
    
    With load.system_info.texts
        open_forms.working_on_it_f .exporting_excels_heading, .exporting_excels_details, 6000
    End With
    
    load.check_secondary_access_app
    str_sql = Me.RecordSource
    Central.open_external_resource_app "reports.accdb", False, load.system_info.system_paths.common_path & "reports.accdb"

    With load.secondary_access_app
        
        str_milestone = "5"
        .Run "generate_weekly_view_extract", str_sql, load.system_info.app_continent
        
        .CloseCurrentDatabase
        
        str_milestone = "6"
        .OpenCurrentDatabase load.system_info.system_paths.common_path & "placeholder.accdb", False
        .Visible = False
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
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Private Sub deal_status_change_AfterUpdate()
    Const proc_name As String = "weekly_view.deal_status_change.AfterUpdate"
    utilities.call_stack_add_item proc_name
    
    If Me!deal_status_change = -1 Then
        GoTo outro
    End If
    
    load.check_conn_and_variables
    
    With global_vars.interfaces.weekly_view_f
        open_forms.working_on_it_f "Updating. Please hold.", "Currencies, deal folders, all that jazz", 3000
        .change_deal_status Me!deal_id, Me.SelTop, Me.deal_status_change
    End With
outro:
    open_forms.working_on_it_f__close
    Exit Sub
End Sub

Private Sub deal_status_Click()
    Forms(Me.Name).Controls("deal_status_change").SetFocus
End Sub

Private Sub Form_Load()
    DoCmd.MoveSize Right:=9250, Down:=0, Width:=21000, Height:=6000
End Sub
Private Sub cmdOpenProject_Click()
    load.check_conn_and_variables
    open_forms.deal_details_f Me!deal_id
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub header_budget_continent_AfterUpdate()
    Dim proc_name As String
    proc_name = "weekly_view_f.header_budget_continent.AfterUpdate"
    load.call_stack = vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    Dim control_name As String
    Dim rs As ADODB.Recordset
    Dim str_sql As String
        
    control_name = "header_budget_region"
    str_sql = load.sources.menu_lists.budget_regions & " WHERE budget_continent_id = " & Me.header_budget_continent
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        
        Do While Me.Controls(control_name).ListCount > 0
            Me.Controls(control_name).RemoveItem (0)
        Loop
        
        Me.Controls(control_name).AddItem "0; '_all'"
        Do While rs.EOF = False
            Me.Controls(control_name).AddItem rs!id & ";" & rs!menu_item
            rs.MoveNext
        Loop
        .Close
        Me.Controls(control_name) = 0
    End With
    
    Set rs = Nothing
    
    fix_rs.weekly_view_f Me.Controls(control_name), , Me!selector_risk_type, Me!header_budget_continent, selector_operational_re, Me!header_selector_uws
outro:
    load.call_stack = ""
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

Private Sub header_budget_region_AfterUpdate()
    Dim proc_name As String
    proc_name = "weekly_view_f.header_budget_region.Click"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    fix_rs.weekly_view_f Me!header_budget_region _
    , _
    , Me!selector_risk_type _
    , Me!header_budget_continent _
    , selector_operational_re _
    , Me!header_selector_uws _
    , global_vars.interfaces.weekly_view_f.header_cmd_foreign_deals.field_value
        
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
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Sub

Private Sub frm_foreign_deals_Click()
    Me!deal_name.SetFocus
End Sub

Private Sub header_selector_uws_AfterUpdate()
    Dim proc_name As String
    proc_name = "weekly_view_f.header_selector_uws.AfterUpdate"
    load.call_stack = vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    fix_rs.weekly_view_f Me!header_budget_region _
    , _
    , Me!selector_risk_type _
    , Me!header_budget_continent _
    , selector_operational_re _
    , Me!header_selector_uws _
    , global_vars.interfaces.weekly_view_f.header_cmd_foreign_deals.field_value
    
outro:
    load.call_stack = ""
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

Private Sub nbi_prepper_change_AfterUpdate()
    Dim proc_name As String
    proc_name = "weekly_view.nbi_prepper_change.AfterUpdate"
    load.call_stack = proc_name
    
    load.check_conn_and_variables
    
    global_vars.interfaces.weekly_view_f.change_nbi_prepper Me!deal_id, Me.SelTop, Me.nbi_prepper_change
    
outro:
    open_forms.working_on_it_f__close
    Exit Sub
End Sub

Private Sub selector_operational_re_AfterUpdate()
    Dim proc_name As String
    proc_name = "weekly_view_f.header_selector_operational_re.AfterUpdate"
    load.call_stack = vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    fix_rs.weekly_view_f Me!header_budget_region _
    , _
    , Me!selector_risk_type _
    , Me!header_budget_continent _
    , selector_operational_re _
    , Me!header_selector_uws _
    , global_vars.interfaces.weekly_view_f.header_cmd_foreign_deals.field_value
    
outro:
    load.call_stack = ""
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

Private Sub selector_risk_type_AfterUpdate()
    Dim proc_name As String
    proc_name = "weekly_view_f.header_selector_risk_type.AfterUpdate"
    load.call_stack = vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    fix_rs.weekly_view_f Me!header_budget_region _
    , _
    , Me!selector_risk_type _
    , Me!header_budget_continent _
    , selector_operational_re _
    , Me!header_selector_uws _
    , global_vars.interfaces.weekly_view_f.header_cmd_foreign_deals.field_value
        
outro:
    load.call_stack = ""
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
Private Sub tog_foreign_deals_Click()
    Const proc_name As String = "weekly_view_f.tog_foreign_deals.Click"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    With global_vars.interfaces.weekly_view_f.tog_foreign_deals
        If .toggle_position = 1 Then
            .cmd_toggle.field_caption = "no"
            .cmd_toggle.field_bg_color = load.colors.light_red
            .cmd_toggle.field_value = False
            .toggle_position = 3
        Else
            .cmd_toggle.field_caption = "yes"
            .cmd_toggle.field_bg_color = load.colors.light_green
            .cmd_toggle.field_value = True
            .toggle_position = 1
        End If
        .change Me.Name
    End With
    
    With Forms(Me.Name)
        fix_rs.weekly_view_f .Controls("header_budget_region") _
        , _
        , .Controls("selector_risk_type") _
        , .Controls("header_budget_continent") _
        , .Controls("selector_operational_re") _
        , .Controls("header_selector_uws") _
        , global_vars.interfaces.weekly_view_f.tog_foreign_deals.cmd_toggle.field_value

        .Controls("deal_name").SetFocus
    End With
outro:
    load.call_stack = ""
    open_forms.working_on_it_f__close
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Private Sub tog_view_density_Click()
    Const proc_name As String = "weekly_view_f.tog_view_density.Click"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim control_height As Long
    Dim control_height_cmd As Long
    Dim margin_top As Long
    Dim margin_bottom As Long
    Dim font_size As Long
    
    load.check_conn_and_variables
    
    With global_vars.interfaces.weekly_view_f.tog_view_density
        If .toggle_position = 1 Then
            .cmd_toggle.field_caption = "normal"
            .cmd_toggle.field_bg_color = load.colors.frost_70
            .toggle_position = 2
            control_height = utilities.twips_converter(0.6, "inch")
            control_height_cmd = utilities.twips_converter(0.17, "inch")
            font_size = 11
            margin_bottom = 0
            margin_top = 0
        ElseIf .toggle_position = 2 Then
            .cmd_toggle.field_caption = "dense"
            .cmd_toggle.field_bg_color = load.colors.frost_70
            .toggle_position = 3
            control_height = utilities.twips_converter(0.5, "inch")
            control_height_cmd = utilities.twips_converter(0.13, "inch")
            font_size = 9
            margin_bottom = 0
            margin_top = 0
        Else
            .cmd_toggle.field_caption = "light"
            .cmd_toggle.field_bg_color = load.colors.frost_70
            .toggle_position = 1
            control_height = utilities.twips_converter(0.8, "inch")
            control_height_cmd = utilities.twips_converter(0.17, "inch")
            font_size = 11
            margin_bottom = utilities.twips_converter(0.08, "inch")
            margin_top = utilities.twips_converter(0.08, "inch")
        End If
        .change Me.Name
        global_vars.interfaces.weekly_view_f.density = .toggle_position
    End With
    Me.Controls("deal_name").SetFocus
    
    With global_vars.interfaces.weekly_view_f
        For Each form_field In .col_text_fields_all
            With form_field
                .field_height = control_height
                .font_size = font_size
                .margin_internal_top = margin_top
                .margin_internal_bottom = margin_bottom
            End With
        Next form_field
        For Each form_field In .col_cmds_change
            With form_field
                .field_height = control_height_cmd
            End With
        Next form_field
        For Each form_field In .col_txts_with_dependencies
            With form_field
                .field_height = control_height_cmd
                .font_size = font_size
            End With
        Next form_field
        
        .cmd_change_nbi_prepper.field_top = margin_top
        .txt_deal_status_change.field_top = margin_top
        .txt_nbi_prepper_change.field_top = margin_top
        .re_init_bottom_cmds
        utilities.paint_control Me.Name, .col_cmds_change
        utilities.paint_control Me.Name, .col_text_fields_all
        utilities.paint_control Me.Name, .col_txts_with_dependencies
    End With
    
    Me.Detail.Height = control_height
    
outro:
    load.call_stack = ""
    open_forms.working_on_it_f__close
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
