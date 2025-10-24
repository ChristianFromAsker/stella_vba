Attribute VB_Name = "open_forms"
Option Compare Database
Option Explicit

Public Sub web_search_f(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "open_forms.web_search_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    load.check_conn_and_variables
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_name As String
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    form_name = load.form_names.web_search_f
    If CurrentProject.AllForms(form_name).IsLoaded = False Then
        DoCmd.OpenForm form_name, acNormal
        DoCmd.MoveSize 500, 500, 10000, 5000
    End If
    With Forms(form_name)
        .Caption = "compliance web search"
        !lbl_deal_id.Caption = deal_id
        str_sql = "SELECT deal_id, deal_name FROM " & load.sources.deals_table & " WHERE deal_id = " & deal_id
        Set rs = utilities.create_adodb_rs(conn, str_sql)
        rs.Open
            .Caption = "compliance web search for " & rs!deal_name
        rs.Close
        Set rs = Nothing
        .SetFocus
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
    
    form_name = load.form_names.uw_positions_f
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
            .FormatConditions.delete
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
Public Sub deal_tags_f(ByVal deal_id As Long)
    load.check_conn_and_variables
    load.call_stack = load.call_stack & vbNewLine & "open_forms.deal_tags_f"
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_name As String, str_sql As String, rs As ADODB.Recordset
    form_name = "deal_tags_f"
    If CurrentProject.AllForms(form_name).IsLoaded = False Then
        DoCmd.OpenForm form_name, acNormal
        DoCmd.MoveSize 500, 500, 18000, 4200
    End If
    With Forms(form_name)
        .Caption = "deal tags"
        !deal_id = deal_id
        str_sql = "SELECT deal_id, deal_name FROM " & load.sources.risk_details_f_view & " WHERE deal_id = " & deal_id
        Set rs = utilities.create_adodb_rs(conn, str_sql)
        rs.Open
            !deal_name = rs!deal_name
        rs.Close
    End With
    
    
    With load.deal_tags
        .init deal_id
        .paint
    End With

outro:
    Set rs = Nothing
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "open_form." & form_name
        .milestone = "str_sql = " & str_sql
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
End Sub
Public Sub add_many_deals_f()
    load.call_stack = load.call_stack & vbNewLine & "open_forms.add_many_deals_f"
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    Dim i As Integer, rs As ADODB.Recordset, str_sql As String, arr_controls() As Variant
    
    Dim form_is_init As Boolean, form_name As String
    form_name = load.add_many_deals.form_name
    
    'rather than closing the form, it is hidden when being closed. Hiding and unhiding is much faster than closing and reopening
    form_is_init = False
    If CurrentProject.AllForms(form_name).IsLoaded = True Then
        With Forms(form_name)
            If .Visible = False Then
                .Visible = True
                form_is_init = True
            End If
        End With
    End If
    
    If form_is_init = False Or load.add_many_deals.shall_be_refreshed = True Then
        DoCmd.OpenForm form_name, acNormal
        load.add_many_deals.shall_be_refreshed = False
        With Forms(form_name)
            .SetFocus
            DoCmd.MoveSize Right:=200, Down:=200, Width:=25200, Height:=14000
            .Caption = "add many deals"
            ReDim arr_controls(0 To 50, 0 To 1)
            
            i = 1
            arr_controls(i, 0) = "header_broker_person"
            arr_controls(i, 1) = load.sources.menu_lists.broker_persons
            
            i = 1 + i
            arr_controls(i, 0) = "header_broker_firm"
            arr_controls(i, 1) = load.sources.menu_lists.broker_firms
            
            i = 1 + i
            arr_controls(i, 0) = "filter_broker_firm"
            arr_controls(i, 1) = load.sources.menu_lists.broker_firms
            
            i = 1 + i
            arr_controls(i, 0) = "header_buyer_law_firm_1"
            arr_controls(i, 1) = load.sources.menu_lists.law_firms
            
            i = 1 + i
            arr_controls(i, 0) = "filter_buyer_law_firm"
            arr_controls(i, 1) = load.sources.menu_lists.law_firms
            
            i = 1 + i
            arr_controls(i, 0) = "header_primary_or_xs"
            arr_controls(i, 1) = load.sources.menu_lists.layer_types
            
            i = 1 + i
            arr_controls(i, 0) = "header_risk_type"
            arr_controls(i, 1) = load.sources.menu_lists.risk_types
            
            i = 1 + i
            arr_controls(i, 0) = "header_stage"
            arr_controls(i, 1) = load.sources.menu_lists.stages
            
            i = 1 + i
            arr_controls(i, 0) = "header_target_super_sector_id"
            arr_controls(i, 1) = load.sources.menu_lists.super_sectors
            
            i = i + 1
            arr_controls(i, 0) = "header_nbi_prepper_id"
            arr_controls(i, 1) = load.sources.menu_lists.uws
            
            i = i + 1
            arr_controls(i, 0) = "header_budget_home"
            arr_controls(i, 1) = load.sources.menu_lists.budget_homes
            
            arr_controls(0, 0) = i
            
            For i = 1 To arr_controls(0, 0)
                Do While .Controls(arr_controls(i, 0)).ListCount > 0
                    .Controls(arr_controls(i, 0)).RemoveItem (0)
                Loop
            Next i
            
            .Controls(load.add_many_deals.filter_broker_firm.field_name).AddItem "-1;_all"
            .Controls(load.add_many_deals.filter_buyer_law_firm.field_name).AddItem "-1;_all"
            
            For i = 1 To arr_controls(0, 0)
                .Controls(arr_controls(i, 0)).AddItem "-1;' '"
                str_sql = arr_controls(i, 1)
                Set rs = utilities.create_adodb_rs(conn, str_sql)
                rs.Open
                Do While rs.EOF = False
                    If rs!menu_item <> "" And rs!menu_item <> " " And rs!menu_item <> "_all" Then
                        .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                    End If
                    rs.MoveNext
                Loop
                rs.Close
            Next i
            Set rs = Nothing
            If load.deal_details.is_init = False Then load.deal_details.init 1
            .Controls("header_" & load.deal_details.txt_budget_home_id.field_name_add_many_deals_f) = buget_homes.rp_us
            
        End With
    End If
    
    With load.add_many_deals
        .default_filters
        .paint
    End With
    
    fix_rs.add_many_deals_f ""
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "open_form.add_many_deals_f"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
End Sub
Public Sub sub_control_f()
    'intro
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim rs As ADODB.Recordset, str_sql As String, i As Integer
    
    Dim str_form As String
    str_form = "sub_control_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    
    With Forms(str_form)
        .Picture = load.form_backgrounds.submission_control_f
        .SetFocus
        DoCmd.MoveSize Right:=200, Down:=200, Width:=21000, Height:=14000
        
        'add items to combo boxes
        'remove any existing lists
        i = 1
        Dim arr_controls() As Variant
        ReDim arr_controls(0 To 50, 0 To 1)
        arr_controls(i, 0) = "header_target_legal_jurisdiction_id"
        arr_controls(i, 1) = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " & load.sources.jurisdictions_view & " WHERE jurisdiction_type = 'country' ORDER BY jurisdiction"
        
        i = i + 1
        arr_controls(i, 0) = "header_spa_law"
        arr_controls(i, 1) = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " & load.sources.jurisdictions_view & " WHERE jurisdiction_type = 'country' ORDER BY jurisdiction"
        
        i = i + 1
        arr_controls(i, 0) = "header_deal_currency"
        arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & load.sources.menu_list_table & " WHERE item_type = 'currency' ORDER BY menu_item "
        
        i = i + 1
        arr_controls(i, 0) = "header_broker_firm_id"
        arr_controls(i, 1) = "SELECT broker_firm_id id, business_name menu_item FROM " & load.sources.broker_firms_view & " ORDER BY business_name"
        
        i = i + 1
        arr_controls(i, 0) = "header_buyer_law_firm_1_id"
        arr_controls(i, 1) = load.sources.menu_lists.law_firms
        
        i = i + 1
        arr_controls(i, 0) = "header_seller_law_firm"
        arr_controls(i, 1) = load.sources.menu_lists.law_firms
        
        i = i + 1
        arr_controls(i, 0) = "header_transaction_style"
        arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & load.sources.menu_list_table & " WHERE item_type = 'transaction_style'"
        
        i = i + 1
        arr_controls(i, 0) = "header_risk_type_id"
        arr_controls(i, 1) = load.sources.menu_lists.risk_types
        
        i = i + 1
        arr_controls(i, 0) = "header_risk_type_id_selector"
        arr_controls(i, 1) = load.sources.menu_lists.risk_types_incl_major
        
        i = i + 1
        arr_controls(i, 0) = "header_primary_or_xs_id"
        arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & load.sources.menu_list_table & " WHERE item_type = 'layer_type'"
        
        i = i + 1
        arr_controls(i, 0) = "header_target_super_sector_id"
        arr_controls(i, 1) = "SELECT sector_id id, sector_name menu_item FROM " & sources.sectors_table & " WHERE is_deleted = 0 AND sector_type = " & load.super_sector & " ORDER BY sector_name"
        
        i = i + 1
        arr_controls(i, 0) = "header_super_sector_selector"
        arr_controls(i, 1) = "SELECT sector_id id, sector_name menu_item FROM " & sources.sectors_table & " WHERE is_deleted = 0 AND sector_type = " & load.super_sector & " ORDER BY sector_name"
        
        i = i + 1
        arr_controls(i, 0) = "header_deal_status_id"
        arr_controls(i, 1) = load.sources.menu_lists.deal_statuses
        
        i = i + 1
        arr_controls(i, 0) = "header_nbi_prepper"
        arr_controls(i, 1) = "SELECT id, uw_initials menu_item FROM " & sources.underwriters_view & " WHERE is_employed_id = 93 AND user_type = 150 ORDER BY uw_initials"
        
        i = i + 1
        arr_controls(i, 0) = "header_budget_home"
        arr_controls(i, 1) = "SELECT entity_id id, entity_business_name menu_item FROM " & load.sources.entities_table & " WHERE is_deleted = 0 AND entity_type = 475 ORDER BY entity_business_name"

        arr_controls(0, 0) = i
        
        For i = 1 To arr_controls(0, 0)
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        
        'add 'all options'
        !header_risk_type_id_selector.AddItem "0;_all"
        !header_budget_home.AddItem "0;_all"
        !header_super_sector_selector.AddItem "0;_all"
        
        'add space to super_sector
        !header_super_sector_selector.AddItem "-10;'Real estate+'"
        !header_super_sector_selector.AddItem "-5;-----"
        
        
        'add Nordics as an RP entity
        !header_budget_home.AddItem load.nordic_rp_entity & ";Nordics"
        
        'add new values
        On Error GoTo err_handler
        If load.is_debugging = True Then On Error GoTo 0
        For i = 1 To arr_controls(0, 0)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
        
        'set default values
        !deal_name_search = ""
        If current_uw.budget_home_id = load.rp_denmark _
            Or current_uw.budget_home_id = load.rp_finland _
            Or current_uw.budget_home_id = load.rp_norway _
            Or current_uw.budget_home_id = load.rp_sweden _
            Then
            !header_budget_home = load.nordic_rp_entity
        Else
            !header_budget_home = current_uw.budget_home_id
        End If
        !header_super_sector_selector = 0
        !header_risk_type_id_selector = 0
    End With

outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "open_forms.sub_control_f"
        .milestone = "str_sql = " & str_sql
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub


Public Sub kpi_dashboard_f()
    Dim str_form As String
    str_form = "kpi_dashboard_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    
    With Forms(str_form)
        'time periods combo box
        'remove current values
        With !time_period
            Do While .ListCount > 0
                .RemoveItem (0)
            Loop
            .AddItem global_vars.time_periods.last_seven_days & ";'last 7 days'"
            .AddItem global_vars.time_periods.last_14_days & ";'last 14 days'"
            .AddItem global_vars.time_periods.last_30_days & ";'last 30 days'"
            .AddItem global_vars.time_periods.last_3_months & ";'last 3 months'"
            .AddItem global_vars.time_periods.last_12_months & ";'last 12 months'"
            .AddItem global_vars.time_periods.previous_month & ";'previous month'"
            .AddItem global_vars.time_periods.previous_year & ";'previous year'"
            .AddItem global_vars.time_periods.this_month & ";'this month'"
            .AddItem global_vars.time_periods.ytd & ";'YTD'"
            .value = global_vars.time_periods.last_14_days
        End With
            
        DoCmd.MoveSize Right:=0, Down:=0, Width:=15000, Height:=12000
        .Caption = "key performance indicator tracking"
    End With
End Sub
Public Sub main_menu_f()
    Dim str_form As String
    str_form = "MenuMainF"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    If load.main_menu.is_init = False Then load.main_menu.init
    
    With Forms(str_form)
        Dim strWelcome As String, str_name As String
        If current_uw.nick = "-1" Then
            str_name = Split(current_uw.full_name, " ")(0)
        Else
            str_name = current_uw.nick
        End If
        strWelcome = "Hi " & str_name & "!"
        !lblHeadingMainMenu.Caption = strWelcome
        
        !cmdNewTra.SetFocus
        !search_input = "Click here to search for deals"
        If current_uw.admin_access = 93 Then
            !admin.Visible = True
        End If
        
        .Picture = load.form_backgrounds.MenuMainF
        
        load.main_menu.update_deal_statuses_for_search
        
        .Controls("cmd_weekly_view").ControlTipText = "Shows recent deals on a weekly basis. Deals are sorted after due date, " _
        & "or, if not available, when they were moved from the NDA stage." & vbNewLine & vbNewLine _
        & "The view shows a snapshot of activity per week."
        
        .SetFocus
        DoCmd.MoveSize Right:=0, Down:=0, Width:=11100, Height:=7000
        .Caption = "Main Menu"
        
        load.main_menu.debugging_buttons
    End With
    load.main_menu.paint_main_menu
End Sub
Public Sub weekly_view_f()
    Dim proc_name As String
    proc_name = "open_forms.weekly_view_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim cmd_foreign_deals As cls_field
    Dim control_name As String
    Dim current_control As Access.Control
    Dim form_field As cls_field
    Dim row_source As String
    Dim rs As ADODB.Recordset
    Dim str_form As String
    
    str_form = global_vars.interfaces.weekly_view_f.form_name
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    
    With global_vars.interfaces.weekly_view_f
        .init
        .paint_visible_only .col_text_fields_all
        .paint_visible_only .col_header_labels_all
        If load.system_info.app_continent = load.system_info.continents.americas Then
            For Each form_field In .col_text_fields_us
                form_field.field_visible = True
            Next form_field
            For Each form_field In .col_header_labels_us
                form_field.field_visible = True
            Next form_field
            
            .paint .col_text_fields_us
            .paint .col_header_labels_us
        Else
            For Each form_field In .col_text_fields
                form_field.field_visible = True
            Next form_field
            For Each form_field In .col_header_labels
                form_field.field_visible = True
            Next form_field
            
            .paint .col_text_fields
            .paint .col_header_labels
        End If
        .init_cmds
        .paint .col_cmds_change
        .header_cmd_foreign_deals.field_value = True
    End With
    
    With Forms(str_form)
        'add values to combo box
        Dim arr_controls() As Variant, i As Integer
        ReDim arr_controls(0 To 10, 0 To 1)
        
        i = 1
        arr_controls(i, 0) = "header_budget_continent"
        arr_controls(i, 1) = load.sources.menu_lists.budget_continents

        i = i + 1
        arr_controls(i, 0) = "header_budget_region"
        arr_controls(i, 1) = load.sources.menu_lists.budget_regions & " WHERE budget_continent_id = " & load.current_uw.budget_continent_id
        
        i = i + 1
        arr_controls(i, 0) = "selector_risk_type"
        arr_controls(i, 1) = load.sources.menu_lists.risk_types_major
        
        i = i + 1
        arr_controls(i, 0) = "selector_operational_re"
        arr_controls(i, 1) = "SELECT DISTINCT(target_sector_group_2) id, ' ' menu_item FROM " & load.sources.weekly_view_f
        
        arr_controls(0, 0) = i
        
        For i = 1 To arr_controls(0, 0)
            .Controls(arr_controls(i, 0)).RowSource = ""
        Next i
        
        control_name = global_vars.interfaces.weekly_view_f.header_selector_uws.field_name
        .Controls(control_name).RowSource = "0;_all;" & load.sources.menu_lists.row_sources_uws
        
        row_source = "-2;'Declined - aim for xs'; -1;'---'; " & load.sources.menu_lists.row_soruce_deal_statuses
        .Controls(global_vars.interfaces.weekly_view_f.txt_deal_status_change.field_name).RowSource = row_source
        
        control_name = global_vars.interfaces.weekly_view_f.txt_nbi_prepper_change.field_name
        .Controls(control_name).RowSource = "-1;' ';" & load.sources.menu_lists.row_sources_uws
        
        'add 'all option'
        !header_budget_region.AddItem "0;_all"
        !selector_risk_type.AddItem "0;_all"
        !selector_operational_re.AddItem "_all;_all"
        
        
        For i = 1 To arr_controls(0, 0)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            rs.Open
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
        
        'set default options
        !header_budget_continent = load.current_uw.budget_continent_id
        !header_selector_uws = 0
        
        Set cmd_foreign_deals = global_vars.interfaces.weekly_view_f.header_cmd_foreign_deals
        
        'load recordset
        If load.system_info.app_continent = load.system_info.continents.americas Then
            !selector_risk_type = 0
            !selector_operational_re = "_all"
            !header_budget_region = 0
            fix_rs.weekly_view_f 0, , 0, load.current_uw.budget_continent_id, "_all", , cmd_foreign_deals.field_value
        Else
            !selector_risk_type = 100
            !selector_operational_re = "_all"
            !header_budget_region = load.current_uw.budget_region_id
            fix_rs.weekly_view_f current_uw.budget_region_id, , 100, load.current_uw.budget_continent_id, "_all", , cmd_foreign_deals.field_value
        End If
        
        'fix row formatting
        Dim objFormatConds As FormatCondition
        
        'format rows based on week number
        Set current_control = .Controls(global_vars.interfaces.weekly_view_f.txt_row_bg.field_name)
        Set objFormatConds = current_control.FormatConditions.Add(acExpression _
        , , "[week_no_number] = " & Format(Date, "WW", vbMonday) - 1 _
        & " Or [week_no_number] = " & Format(Date, "WW", vbMonday) - 3 _
        & " Or [week_no_number] = " & Format(Date, "WW", vbMonday) - 5 _
        & " Or [week_no_number] = " & Format(Date, "WW", vbMonday) - 7 _
        & " Or [week_no_number] = " & Format(Date, "WW", vbMonday) - 9 _
        & " Or [week_no_number] = " & Format(Date, "WW", vbMonday) - 11)
                
        current_control.FormatConditions(0).BackColor = load.colors.week_1
        
        'format deal status
        Set current_control = .Controls("deal_status")
        Set objFormatConds = current_control.FormatConditions.Add(acExpression, , "[status_id] = " & load.deal_statuses.declined)
        current_control.FormatConditions(0).BackColor = load.colors.light_red
        
        Set objFormatConds = current_control.FormatConditions.Add(acExpression, , "[status_id] = " & load.deal_statuses.nbi _
        & " OR [status_id] = " & load.deal_statuses.preferred _
        & " OR [status_id] = " & load.deal_statuses.expensed _
        & " OR [status_id] = " & load.deal_statuses.uw)
        current_control.FormatConditions(1).BackColor = load.colors.quoted
            
        Set objFormatConds = current_control.FormatConditions.Add(acExpression, , "[status_id] = " & load.deal_statuses.submission)
        current_control.FormatConditions(2).BackColor = load.colors.submission_color
        
        'format risk type and buyer business name
        With global_vars.interfaces.weekly_view_f
            If load.system_info.app_continent = load.system_info.continents.americas Then
                'risk type
                Set current_control = Forms(str_form).Controls(.txt_deal_name.field_name)
    
                Set objFormatConds = current_control.FormatConditions.Add(acExpression, _
                , "[" & .txt_risk_type_major_id.field_name & "] = " & global_vars.risk_types.wi)
                current_control.FormatConditions(0).BackColor = load.colors.risk_type_wi

                Set objFormatConds = current_control.FormatConditions.Add(acExpression, _
                , "[" & .txt_risk_type_major_id.field_name & "] = " & global_vars.risk_types.tax)
                current_control.FormatConditions(1).BackColor = load.colors.risk_type_tax

                Set objFormatConds = current_control.FormatConditions.Add(acExpression, _
                , "[" & .txt_risk_type_major_id.field_name & "] = " & global_vars.risk_types.contingency)
                current_control.FormatConditions(2).BackColor = load.colors.risk_type_cont
                
                'buyer business name
                Set current_control = Forms(str_form).Controls(.txt_buyer_business_name.field_name)
                Set objFormatConds = current_control.FormatConditions.Add(acExpression, _
                , "[" & .txt_is_repeat_buyer.field_name & "] = ""yes""")
                current_control.FormatConditions(0).BackColor = load.colors.Is_repeat_buyer
            End If
        End With
        
        'format submission notes based on whether the word 'joinder' is in there
        With global_vars.interfaces.weekly_view_f
            If load.system_info.app_continent = load.system_info.continents.americas Then
                Set current_control = Forms(str_form).Controls(.txt_submission_notes.field_name)
    
'                Set objFormatConds = current_control.FormatConditions.Add(acExpression, _
'                , instr("[" & .txt_submission_notes & "], ) = " & "joinder")
                ' current_control.FormatConditions(0).BackColor = load.colors.yellow
            End If
        End With
        
        Set current_control = Nothing
        
        'movesize
        Dim form_width As Long
        form_width = 23000
        If load.system_info.app_continent = load.system_info.continents.americas Then
            form_width = 27000
        End If
        DoCmd.MoveSize Right:=600, Down:=600, Width:=form_width, Height:=14000
        
        !selector_operational_re.SetFocus
    End With
    
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
Public Sub live_deals_f()
    Dim proc_name As String
    proc_name = "open_forms.live_deals_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim arr_controls() As Variant
    Dim current_control As Access.Control
    Dim deal_count As Long
    Dim form_field As cls_field
    Dim i As Integer
    Dim objFormatConds As FormatCondition
    Dim str_form As String
    
    str_form = "live_deals_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    
    With global_vars.interfaces.live_deals_f
        .init
'        .paint_visible_only .col_text_fields_all
'        .paint_visible_only .col_header_labels_all
        If load.system_info.app_continent = load.system_info.continents.americas Then
'            For Each form_field In .col_text_fields_us
'                form_field.field_visible = True
'            Next form_field
'            For Each form_field In .col_header_labels_us
'                form_field.field_visible = True
'            Next form_field
'
            .paint .col_text_fields_us
            .paint .col_header_labels_us
        Else
'            For Each form_field In .col_text_fields
'                form_field.field_visible = True
'            Next form_field
'            For Each form_field In .col_header_labels
'                form_field.field_visible = True
'            Next form_field
'
            .paint .col_text_fields
            .paint .col_header_labels
        End If
'        .init_cmds
'        .paint .col_cmds_change
    End With
    
    With Forms(str_form)
        'add values to combo box
        ReDim arr_controls(0 To 10, 0 To 1)
        
        i = 1
        arr_controls(i, 0) = "header_selector_budget_continent"
        arr_controls(i, 1) = load.sources.menu_lists.budget_continents

        i = i + 1
        arr_controls(i, 0) = "header_selector_budget_region"
        arr_controls(i, 1) = load.sources.menu_lists.budget_regions & " WHERE budget_continent_id = " & load.current_uw.budget_continent_id
        
        i = i + 1
        arr_controls(i, 0) = "header_selector_risk_type"
        arr_controls(i, 1) = load.sources.menu_lists.risk_types_major
        
        i = i + 1
        arr_controls(i, 0) = "header_selector_operational_re"
        arr_controls(i, 1) = "SELECT DISTINCT(target_sector_group_2) id, ' ' menu_item FROM " & load.sources.live_deals_view
        
        i = i + 1
        arr_controls(i, 0) = "header_selector_uws"
        arr_controls(i, 1) = load.sources.menu_lists.uws
        
        arr_controls(0, 0) = i
        
        For i = 1 To arr_controls(0, 0)
            .Controls(arr_controls(i, 0)).RowSource = ""
        Next i
        
        'add manual options to drop-downs'
        .Controls(global_vars.interfaces.live_deals_f.header_selector_budget_region.field_name).AddItem "0;_all"
        .Controls(global_vars.interfaces.live_deals_f.header_selector_risk_type.field_name).AddItem "0;_all"
        .Controls(global_vars.interfaces.live_deals_f.header_selector_operational_re.field_name).AddItem "_all;_all"
        With .Controls(global_vars.interfaces.live_deals_f.header_selector_uws.field_name)
            .AddItem "0;_all"
            .AddItem "-1;'---'"
            .AddItem load.current_uw.uw_id & ";" & load.current_uw.initials
            .AddItem "-1;'---'"
        End With
        
        Dim rs As ADODB.Recordset
        For i = 1 To arr_controls(0, 0)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            rs.Open
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
        
        'set default values
        .Controls(global_vars.interfaces.live_deals_f.header_selector_budget_continent.field_name) = load.current_uw.budget_continent_id
        .Controls(global_vars.interfaces.live_deals_f.header_selector_uws.field_name) = 0
        
        Dim ctl_budget_region As Access.Control
        Dim ctl_operational_re As Access.Control
        Dim ctl_risk_type As Access.Control
        
        Set ctl_budget_region = .Controls(global_vars.interfaces.live_deals_f.header_selector_budget_region.field_name)
        Set ctl_operational_re = .Controls(global_vars.interfaces.live_deals_f.header_selector_operational_re.field_name)
        Set ctl_risk_type = .Controls(global_vars.interfaces.live_deals_f.header_selector_risk_type.field_name)
        
        If load.system_info.app_continent = load.system_info.continents.americas Then
            ctl_risk_type = 0
            ctl_operational_re = "_all"
            ctl_budget_region = 0
            deal_count = fix_rs.live_deals_f(0, , 0, load.current_uw.budget_continent_id)
        Else
            ctl_risk_type = 100
            ctl_operational_re = "operational"
            ctl_budget_region = load.current_uw.budget_region_id
            deal_count = fix_rs.live_deals_f(current_uw.budget_region_id, , 100, load.current_uw.budget_continent_id)
        End If
        
        'fix row formatting
'        ReDim arr_controls(0 To 14)
'        i = 1
'        arr_controls(i) = "broker_info": i = i + 1
'        arr_controls(i) = "budget_home": i = i + 1
'        arr_controls(i) = "buyer_business_name": i = i + 1
'        arr_controls(i) = "change_roles": i = i + 1
'        arr_controls(i) = "deal_info": i = i + 1
'        If load.system_info.app_continent <> load.system_info.continents.americas Then
'            arr_controls(i) = "deal_name": i = i + 1
'        End If
'        arr_controls(i) = "ev_deal": i = i + 1
'        arr_controls(i) = "nbi_deadline": i = i + 1
'        arr_controls(i) = "roles": i = i + 1
'        arr_controls(i) = "submission_limits": i = i + 1
'        arr_controls(i) = "submission_notes": i = i + 1
'        arr_controls(i) = "target_business_name": i = i + 1
'        arr_controls(i) = "target_description": i = i + 1
'        arr_controls(i) = "week_no"
'
'        arr_controls(0) = i
'
'        For i = 0 To arr_controls(0)
'            Set objFormatConds = .Controls(arr_controls(i)).FormatConditions.Add(acExpression, , "[week_no_number] = " & Format(Date, "WW", vbMonday) - 1 _
'                & " Or [week_no_number] = " & Format(Date, "WW", vbMonday) - 3 _
'                & " Or [week_no_number] = " & Format(Date, "WW", vbMonday) - 5 _
'                & " Or [week_no_number] = " & Format(Date, "WW", vbMonday) - 7 _
'                & " Or [week_no_number] = " & Format(Date, "WW", vbMonday) - 9 _
'                & " Or [week_no_number] = " & Format(Date, "WW", vbMonday) - 11)
'
'            .Controls(arr_controls(i)).FormatConditions(0).BackColor = load.colors.week_1
'        Next i
'
        With .Controls("deal_status")
            Set objFormatConds = .FormatConditions.Add(acExpression, , "[status_id] = " & load.deal_statuses.declined)
            .FormatConditions(0).BackColor = load.colors.light_red

            Set objFormatConds = .FormatConditions.Add(acExpression, , "[status_id] = " & load.deal_statuses.nbi _
                & " OR [status_id] = " & load.deal_statuses.preferred _
                & " OR [status_id] = " & load.deal_statuses.expensed _
                & " OR [status_id] = " & load.deal_statuses.uw)
            .FormatConditions(1).BackColor = load.colors.quoted

            Set objFormatConds = .FormatConditions.Add(acExpression, , "[status_id] = " & load.deal_statuses.submission)
            .FormatConditions(2).BackColor = load.colors.submission_color
        End With

        With global_vars.interfaces.live_deals_f
            Set current_control = Forms(str_form).Controls(.txt_deal_name.field_name)
            If load.system_info.app_continent = load.system_info.continents.americas Then
                Set objFormatConds = current_control.FormatConditions.Add(acExpression, , "[" & .txt_risk_type_major_id.field_name & "] = " & global_vars.risk_types.wi)
                current_control.FormatConditions(0).BackColor = load.colors.risk_type_wi
                current_control.FormatConditions(0).FontBold = True

                Set objFormatConds = current_control.FormatConditions.Add(acExpression, , "[" & .txt_risk_type_major_id.field_name & "] = " & global_vars.risk_types.tax)
                current_control.FormatConditions(1).BackColor = load.colors.risk_type_tax
                current_control.FormatConditions(1).FontBold = True

                Set objFormatConds = current_control.FormatConditions.Add(acExpression, , "[" & .txt_risk_type_major_id.field_name & "] = " & global_vars.risk_types.contingency)
                current_control.FormatConditions(2).BackColor = load.colors.risk_type_cont
                current_control.FormatConditions(2).FontBold = True
            End If
        End With
        
        With global_vars.interfaces.live_deals_f
            Set current_control = Forms(str_form).Controls(.txt_ev_deal.field_name)
            If load.system_info.app_continent = load.system_info.continents.americas Then
                Set objFormatConds = current_control.FormatConditions.Add(acExpression, , "[" & .txt_primary_or_xs_id.field_name & "] = 438")
                current_control.FormatConditions(0).BackColor = load.colors.frost
                
                Set objFormatConds = current_control.FormatConditions.Add(acExpression, , "[" & .txt_primary_or_xs_id.field_name & "] = 439")
                current_control.FormatConditions(1).BackColor = load.colors.frost_50
                
                Set objFormatConds = current_control.FormatConditions.Add(acExpression, , "[" & .txt_primary_or_xs_id.field_name & "] = 440")
                current_control.FormatConditions(2).BackColor = load.colors.frost_30
                
            End If
        End With
        
        open_forms.live_deals_f_move_size deal_count
        
        ctl_operational_re.SetFocus
    End With
    
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
Public Sub live_deals_f_move_size(ByVal deal_count As Long)
    Dim proc_name As String
    proc_name = "open_forms.live_deals_f_move_size"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
        
    Dim form_height As Long
    Dim form_width As Long
    
    form_height = deal_count * utilities.twips_converter(0.8, "inch") + utilities.twips_converter(2.6, "inch")
    If form_height > windows_apis.PositionFormFillVertical(1, "center") * 0.9 Then
        form_height = windows_apis.PositionFormFillVertical(1, "center") * 0.9
    End If
    form_width = 22000
'        If load.system_info.app_continent = load.system_info.continents.americas Then
'            form_width = 27000
'        End If
    DoCmd.MoveSize Right:=600, Down:=0, Width:=form_width, Height:=form_height
            
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
Public Sub working_on_it_f(ByVal input_text As String, Optional ByVal text_details As String, Optional ByVal form_height As Long)
    load.call_stack = load.call_stack & vbNewLine & "open_forms.working_on_it_f"
    Dim str_form As String
    On Error Resume Next
    str_form = load.form_names.working_on_it_f
    
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    Else
        'working on it already showing, so do nothing
        GoTo outro
    End If
    
    With Forms(str_form)
        !input_text.value = input_text
        !text_details = text_details
        !placeholder.SetFocus
    End With
    Dim int_height As Long
    int_height = 4000
    If form_height > 0 Then int_height = form_height
    DoCmd.MoveSize Right:=200, Down:=200, Width:=10000, Height:=int_height
    
outro:
    Exit Sub
    
End Sub
Public Sub working_on_it_f__close()
    On Error Resume Next
    If CurrentProject.AllForms(load.form_names.working_on_it_f).IsLoaded = True Then DoCmd.Close acForm, load.form_names.working_on_it_f
End Sub
Public Sub my_deals_f(ByVal str_condition_r As String)
    Dim proc_name As String
    proc_name = "open_forms.my_deals_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim intHeight As Integer, x As Integer, y As Integer, int_user As Integer
    Dim str_form As String
    
    str_form = "my_deals_f"
    'If Environ("username") = "christian.kartnes" Then str_form = "my_deals_f_new"
    
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    End If
    Dim deal_count As Long
    deal_count = fix_rs.my_deals_f
    If deal_count = 0 Then
        GoTo outro
    End If
    With Forms(str_form)
        .Caption = "Overview of your deals. There are " & deal_count & " of them."
        .SetFocus
        DoCmd.MoveSize Right:=11000, Down:=400, Width:=12500, Height:=6000
    End With
    
outro:
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_condition = " & str_condition_r
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Sub
Public Sub policies_f(ByVal deal_id As Long)
    'intro
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String, rs As ADODB.Recordset, i As Integer
    Dim str_form As String
    str_form = load.policies.form_name
    
    str_sql = "SELECT deal_id, deal_name, total_rp_premium_on_deal, total_rp_limit_on_deal, lowest_rp_attpoint" _
    & " FROM " & load.sources.deals_view & " WHERE deal_id = " & deal_id
    
    Dim deal_name As String
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        deal_name = rs!deal_name
    
        If CurrentProject.AllForms(str_form).IsLoaded = False Then
            DoCmd.OpenForm str_form
        End If
        
        Dim policy_count As Integer
        policy_count = CInt(fix_rs.policies_f(deal_id))
        
        'Move focus to top of list
        Dim form_height As Long
        With Forms(str_form)
            .Caption = "Policies for " & deal_name & " (" & deal_id & ")"
            !total_rp_premium_on_deal = rs!total_rp_premium_on_deal
            !total_rp_limit_on_deal = rs!total_rp_limit_on_deal
            !lowest_rp_attpoint = rs!lowest_rp_attpoint
            .SelTop = 1
            form_height = 2200 + utilities.twips_converter(3.6, "cm") * policy_count
            If form_height > 8000 Then form_height = 8000
            .SetFocus
            DoCmd.MoveSize 500, 500, 16000, form_height
        End With
    rs.Close
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: open_forms.policies_f" & vbNewLine _
        & "Parameters: deal_id = " & deal_id & vbNewLine _
        & "App: " & load.system_info.app_name, , load.system_info.error_msg_heading
    GoTo outro
End Sub
Public Sub deal_details_f(ByVal deal_id As Long)
    load.call_stack = load.call_stack & vbNewLine & "open_forms.deal_details_f"
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim box_input As String
    Dim i As Integer
    Dim init_deal_details As Boolean
    Dim item_id As Long
    Dim item_text As String
    Dim rs As ADODB.Recordset
    Dim str_form As String
    Dim str_sql As String
    Dim timer_start As Single
    Dim y As Integer
    
    timer_start = Timer
    
    str_sql = "SELECT deal_id FROM " & load.sources.deals_view & " WHERE deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        If rs.BOF And rs.EOF Then
            MsgBox "Deal with deal_id " & deal_id & " cannot be opened. Snip (screenshot) this to Christian" _
                & " (christian.kartnes@rpgroup.com) and Tom (tom.evans@rpgroup.com).", , "Deal cannot be opened."
            GoTo outro
        End If
    rs.Close
    
    str_form = "deal_details_f"
    
    'rather than closing deal_details_f, it is hidden when being closed. Hiding and unhiding is much faster than closing and reopening
    init_deal_details = False
    If CurrentProject.AllForms(str_form).IsLoaded = True Then
        With Forms(str_form)
            If .Visible = False Then
                .Visible = True
                init_deal_details = True
            End If
            .SetFocus
        End With
    End If
    
    If load.deal_details.is_init = False Or load.deal_details.shall_be_refreshed = True Then
        init_deal_details = True
        load.deal_details.init deal_id
    End If
    
    timer_start = Timer
    
    'if form is not loaded, load it, and add items to all combo boxes
    If CurrentProject.AllForms(str_form).IsLoaded = False Or load.deal_details.shall_be_refreshed = True Then
        load.deal_details.shall_be_refreshed = False
        init_deal_details = True
        DoCmd.OpenForm str_form
        Forms(str_form).SetFocus
        DoCmd.MoveSize Right:=50, Down:=50, Width:=21000, Height:=12500
        
        With Forms(str_form)
            'add values to combo boxes.
            
            'add countries to country drop-downs. Separate process to reduce queries to database.
            ReDim arr_controls(0 To 100, 0 To 1)
            i = 1
            arr_controls(i, 0) = "spa_law"
            
            i = i + 1
            arr_controls(i, 0) = "insured_registered_country_id"
            
            i = i + 1
            arr_controls(i, 0) = "target_country"
            
            i = i + 1
            arr_controls(i, 0) = "target_main_jurisdiction_id"
            
            arr_controls(0, 0) = i
            
            For i = 1 To arr_controls(0, 0)
                .Controls(arr_controls(i, 0)).RowSource = ""
                box_input = ""
                For y = 0 To UBound(load.country_list)
                    If load.country_list(y, 1) <> "_all" Then
                        box_input = box_input & load.country_list(y, 0) & ";" & load.country_list(y, 1) & ";"
                    End If
                Next y
                .Controls(arr_controls(i, 0)).RowSource = box_input
            Next i
            
            'add items based on pre-loaded broker_firm array (to reduce queries to the db)
            With .Controls("broker_firm_id")
                .RowSource = ""
                For y = 1 To UBound(load.broker_firm_array)
                    .AddItem load.broker_firm_array(y, 0) & ";" & load.broker_firm_array(y, 1) & ";" & load.broker_firm_array(y, 2)
                Next y
            End With
            
            'add items based on pre-loaded yes_no array (to reduce queries to the db)
            ReDim arr_controls(0 To 8, 0 To 1)
            i = 0
            'yes nos
            arr_controls(i, 0) = "was_quoted_id": i = i + 1
            arr_controls(i, 0) = "ClosingBooked": i = i + 1
            arr_controls(i, 0) = "CounselInvoiceReceived": i = i + 1
            arr_controls(i, 0) = "VDRReceived": i = i + 1
            arr_controls(i, 0) = "rr_done": i = i + 1
            arr_controls(i, 0) = "c5_emails_filed": i = i + 1
            arr_controls(i, 0) = "is_underwritten": i = i + 1
            arr_controls(i, 0) = "is_test_deal_id": i = i + 1
            arr_controls(i, 0) = "closing_set_received_id": i = i + 1
            
            box_input = ""
            For y = 1 To UBound(load.yes_no_array)
                box_input = box_input & load.yes_no_array(y, 0) & ";" & load.yes_no_array(y, 1) & ";"
            Next y
                
            For i = 0 To UBound(arr_controls)
                .Controls(arr_controls(i, 0)).RowSource = box_input
            Next i
                
             
            'add law firms to drop-downs
            ReDim arr_controls(0 To 100, 0 To 1)
            i = i
            arr_controls(i, 0) = load.deal_details.txt_uw_counsel_1_id.field_name
            arr_controls(i, 1) = 93
            
            i = i + 1
            arr_controls(i, 0) = "SellerLegalFirm"
            arr_controls(i, 1) = 94
            
            i = i + 1
            arr_controls(i, 0) = "buyer_law_firm_1_id"
            arr_controls(i, 1) = 94
            
            i = i + 1
            arr_controls(i, 0) = "buyer_law_firm_2_id"
            arr_controls(i, 1) = 94
            
            i = i + 1
            arr_controls(i, 0) = "uw_law_firm_id"
            arr_controls(i, 1) = 93
            
            arr_controls(0, 0) = i
            
            For i = 1 To arr_controls(0, 0)
                .Controls(arr_controls(i, 0)).RowSource = ""
                box_input = ""
                For y = 1 To load.array_law_firms(0, 0)
                    item_text = load.array_law_firms(y, 1)
                    item_id = load.array_law_firms(y, 0)
                    'check if only counsel lawyers shall be added to the control.
                
                    If arr_controls(i, 1) = menu_list.yes Then
                        If load.array_law_firms(y, 2) = menu_list.yes Then
                            box_input = box_input & item_id & ";'" & item_text & "';"
                        End If
                    Else
                        box_input = box_input & item_id & ";'" & item_text & "';"
                    End If
                Next y
                 .Controls(arr_controls(i, 0)).RowSource = box_input
            Next i
            
            'add deal statuses
            Dim menu_item As Scripting.Dictionary
            box_input = ""
            With .Controls(load.deal_details.txt_deal_status_id.field_name)
                For Each menu_item In global_vars.col_deal_statuses
                    box_input = box_input & menu_item("id") & ";'" & menu_item("menu_item") & "'" & ";"
                Next menu_item
                .RowSource = box_input
            End With
    
            'add the rest of the combo boxes

            Dim menu_item_continent As String
            menu_item_continent = "menu_item"
            If load.system_info.app_continent = load.system_info.continents.americas Then menu_item_continent = "menu_item_us menu_item"
            
            ReDim arr_controls(0 To 100, 0 To 1)
            i = 1
            arr_controls(i, 0) = "deal_currency"
            arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & load.sources.menu_list_table & " WHERE item_type = 'Currency' AND is_deleted = 0 ORDER BY menu_item"
            
            i = i + 1
            arr_controls(i, 0) = "stage_id"
            arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & load.sources.menu_list_table & " WHERE item_type = 'stage' ORDER BY menu_item"
            
            i = i + 1
            arr_controls(i, 0) = "vat_home"
            arr_controls(i, 1) = "SELECT id, entity_business_name menu_item FROM " & load.sources.rp_entity_info_view & " WHERE entity_type = 475"
            
            i = i + 1
            arr_controls(i, 0) = "budget_home_id"
            arr_controls(i, 1) = "SELECT id, entity_business_name menu_item FROM " & load.sources.rp_entity_info_view & " WHERE entity_type = 475 ORDER BY entity_business_name"
            
            i = i + 1
            arr_controls(i, 0) = "primary_insurer"
            arr_controls(i, 1) = "SELECT id, insurer_business_name menu_item FROM " & load.sources.insurers_view & " WHERE is_active = 93 ORDER BY insurer_business_name"
            
            i = i + 1
            arr_controls(i, 0) = "UwCounselPerson1"
            arr_controls(i, 1) = "SELECT id, personal_name menu_item FROM " & load.sources.lawyers_view & " ORDER BY personal_name"
            
            i = i + 1
            arr_controls(i, 0) = "UwCounselPerson2"
            arr_controls(i, 1) = "SELECT id, personal_name menu_item FROM " & load.sources.uw_financial_persons_view & " ORDER BY personal_name"
            
            i = i + 1
            arr_controls(i, 0) = "broker_person"
            arr_controls(i, 1) = load.sources.menu_lists.broker_persons
            
            i = i + 1
            arr_controls(i, 0) = "target_super_sector_id"
            arr_controls(i, 1) = "SELECT sector_id id, sector_name menu_item FROM " & sources.sectors_table & " WHERE is_deleted = 0 AND sector_type = " & load.super_sector & " ORDER BY sector_name"
            
            i = i + 1
            arr_controls(i, 0) = "risk_type_id"
            arr_controls(i, 1) = load.sources.menu_lists.risk_types
            
            i = i + 1
            arr_controls(i, 0) = "risk_feel_id"
            arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & load.sources.menu_list_table & " WHERE is_deleted = 0 AND item_type = 'DealQuality' ORDER BY Setting1"
            
            i = i + 1
            arr_controls(i, 0) = "primary_or_xs_id"
            arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & load.sources.menu_list_table & " WHERE is_deleted = 0 AND item_type = 'layer_type' ORDER BY menu_item; "
            
            i = i + 1
            arr_controls(i, 0) = "BuyerFinancialCounsel"
            arr_controls(i, 1) = "SELECT financial_advisor_id id, firm_name menu_item FROM " & load.sources.financial_advisors_view & "  ORDER BY firm_name"
            
            i = i + 1
            arr_controls(i, 0) = load.deal_details.txt_uw_financial_advisor_id.field_name
            arr_controls(i, 1) = "SELECT id, firm_name menu_item FROM " & load.sources.financial_advisors_view & " WHERE is_counsel = " & load.yes & " ORDER BY firm_name"
            
            'controls with logic
            
            'three column boxes
            i = i + 1
            arr_controls(i, 0) = "nbi_prepper_id"
            arr_controls(i, 1) = load.sources.menu_lists.uws
            
            i = i + 1
            arr_controls(i, 0) = "primary_uw"
            arr_controls(i, 1) = load.sources.menu_lists.uws
            
            i = i + 1
            arr_controls(i, 0) = "secondary_uw"
            arr_controls(i, 1) = load.sources.menu_lists.uws
            
            i = i + 1
            arr_controls(i, 0) = "analyst_id"
            arr_controls(i, 1) = load.sources.menu_lists.uws
            
            i = i + 1
            arr_controls(i, 0) = "internal_approver_quote_id"
            arr_controls(i, 1) = load.sources.menu_lists.uws_that_have_quote_righst
            
            i = i + 1
            arr_controls(i, 0) = "internal_approver_binding_id"
            arr_controls(i, 1) = load.sources.menu_lists.uws_that_have_binding_rights
            
            arr_controls(0, 0) = i
            
            'add new values
            For i = 1 To arr_controls(0, 0)
                box_input = ""
                str_sql = arr_controls(i, 1)
                Set rs = utilities.create_adodb_rs(conn, str_sql)
                rs.Open
                    Do While rs.EOF = False
                        box_input = box_input & rs!id & ";'" & rs!menu_item & "'" & ";"
                        rs.MoveNext
                    Loop
                rs.Close
                .Controls(arr_controls(i, 0)).RowSource = box_input
            Next i
            
            'fix email list
            Do While !emails.ListCount > 0
                !emails.RemoveItem (0)
            Loop
            !emails.AddItem 0 & "; "
            !emails.AddItem 1 & "; nbi approval email"
            !emails.AddItem 3 & "; closing pr email"
            !emails.AddItem 4 & "; booking email for Navins"
            !emails.AddItem 5 & "; Capacity Referral Email"
            !emails = 0
            
            !deal_status_id.BackColor = colors.white
            If load.current_uw.admin_access <> load.yes Then
                !budget_home_id.BackColor = colors.locked_field
            End If
            
            !deal_id.SetFocus
        End With
    End If
    
    'CK wants to not re-init the form, but for now it is required to make the notification area run
    load.deal_details.set_bg_of_txt_fields_white
    With load.deal_details
        .refresh_deal_details deal_id
        .paint_deal_details
    End With
    
    'populate fields for collapsable sections if they not collapsed
    With load.deal_details
        If .header_admin.user_has_expanded_section = True Then
            .put_data_into_form .col_txt_fields_for_admin_section, deal_id
        End If
        If .header_closing_info.user_has_expanded_section = True Then
            .put_data_into_form .col_txt_fields_for_closing_info, deal_id
        End If
        If .header_submission_notes.user_has_expanded_section = True Then
            .put_data_into_form .col_txt_fields_for_submission_notes_section, deal_id
        End If
    End With
    
    'fix sub sector
    With Forms(str_form)
        Do While !target_sub_sector_id.ListCount > 0
            !target_sub_sector_id.RemoveItem (0)
        Loop
    End With
    Dim rs_sector As ADODB.Recordset
    str_sql = "SELECT deal_id, target_super_sector_id FROM " & sources.deals_table & " WHERE deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        If IsNull(rs!target_super_sector_id) = False Then
            str_sql = "SELECT sector_id id, sector_name menu_item FROM " & sources.sectors_table & " WHERE parent_sector_id = " & rs!target_super_sector_id & " ORDER BY sector_name"
            Set rs_sector = utilities.create_adodb_rs(conn, str_sql)
            With rs_sector
                .Open
                Do Until .EOF
                    Forms(str_form)!target_sub_sector_id.AddItem rs_sector!id & ";'" & rs_sector!menu_item & "'"
                    .MoveNext
                Loop
                .Close
            End With
        End If
    rs.Close
    
    'Populate fields not dealt with above
    Central.populate_deal_details_f deal_id
    
    load.deal_tags.put_tags_on_deal_details_f deal_id

outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    open_forms.working_on_it_f__close
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "open_forms.deal_details_f"
        .milestone = "str_sql = " & str_sql
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub deal_list_f()
    Dim proc_name As String
    proc_name = "open_forms.deal_list_f"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    'used for testing speed of proc
    Dim timer_start As Single
    timer_start = Timer
            
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim str_form As String
    
    str_form = load.master_deal_list.form_name
    
    If load.master_deal_list.is_init = False Then load.master_deal_list.init
        
    'rather than closing the form, it is hidden when being closed. Hiding and unhiding is much faster than closing and reopening
    Dim form_is_init As Boolean
    form_is_init = False
    If CurrentProject.AllForms(str_form).IsLoaded = True Then
        With Forms(str_form)
            If .Visible = False Then
                .Visible = True
                form_is_init = True
            End If
        End With
    End If
    
    'open form if necessary
    If form_is_init = False Or load.master_deal_list.shall_be_refreshed = True Then
        DoCmd.OpenForm str_form, acNormal
        DoCmd.MoveSize Right:=300, Down:=300, Width:=25500, Height:=5000
        load.master_deal_list.shall_be_refreshed = False
        With Forms(str_form)
            .Caption = "Master Deal List | " & UCase(Left(load.system_info.app_continent, 1)) & Right(load.system_info.app_continent, Len(load.system_info.app_continent) - 1)
            'add values to combo boxes
            Dim arr_controls() As Variant
            ReDim arr_controls(0 To 50, 0 To 1)
            i = 1
            
            With load.master_deal_list
                arr_controls(i, 0) = .txt_filter_broker_firm.field_name
                arr_controls(i, 1) = load.sources.menu_lists.broker_firms
                                
                i = i + 1
                arr_controls(i, 0) = .txt_filter_law_firm.field_name
                arr_controls(i, 1) = "SELECT id, FirmName menu_item FROM " & load.sources.law_firms_view & " ORDER BY FirmName"
                
                i = i + 1
                arr_controls(i, 0) = .txt_filter_super_sector.field_name
                arr_controls(i, 1) = "SELECT sector_id id, sector_name menu_item FROM " & load.sources.sectors_table & " WHERE sector_type = 494 AND is_deleted = 0 ORDER BY sector_name"
                
                i = i + 1
                arr_controls(i, 0) = .txt_filter_spa_law.field_name
                arr_controls(i, 1) = "SELECT jurisdiction id, jurisdiction menu_item FROM " & sources.jurisdictions_view & " WHERE jurisdiction_type = 'country' ORDER BY jurisdiction_type, jurisdiction"
                
                i = i + 1
                arr_controls(i, 0) = .txt_filter_insured_registered_country.field_name
                arr_controls(i, 1) = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " & sources.jurisdictions_view & " WHERE jurisdiction_type = 'country' ORDER BY jurisdiction_type, jurisdiction"
                
                i = i + 1
                arr_controls(i, 0) = .txt_filter_primary_or_xs.field_name
                arr_controls(i, 1) = load.sources.menu_lists.layer_types
                
                i = i + 1
                arr_controls(i, 0) = .txt_filter_risk_type.field_name
                arr_controls(i, 1) = load.sources.menu_lists.risk_types_incl_major
                
                i = i + 1
                arr_controls(i, 0) = .txt_filter_primary_uw.field_name
                arr_controls(i, 1) = load.sources.menu_lists.uws
                
                i = i + 1
                arr_controls(i, 0) = .txt_filter_second_uw.field_name
                arr_controls(i, 1) = load.sources.menu_lists.uws
                
                i = i + 1
                arr_controls(i, 0) = .txt_filter_nbi_prepper.field_name
                arr_controls(i, 1) = load.sources.menu_lists.uws
                
                i = i + 1
                arr_controls(i, 0) = .txt_filter_was_quoted.field_name
                arr_controls(i, 1) = load.sources.menu_lists.yes_no
            End With
            
            arr_controls(0, 0) = i
            
            'remove existing values
            For i = 1 To arr_controls(0, 0)
                Do While .Controls(arr_controls(i, 0)).ListCount > 0
                    .Controls(arr_controls(i, 0)).RemoveItem (0)
                Loop
            Next i
            
            'add _all options
            .Controls(load.master_deal_list.txt_filter_broker_firm.field_name).AddItem "-1; _All"
            .Controls(load.master_deal_list.txt_filter_insured_registered_country.field_name).AddItem "-1; _All"
            .Controls(load.master_deal_list.txt_filter_law_firm.field_name).AddItem "-1; _All"
            .Controls(load.master_deal_list.txt_filter_nbi_prepper.field_name).AddItem "-1; _All"
            .Controls(load.master_deal_list.txt_filter_primary_or_xs.field_name).AddItem "-1; _All"
            .Controls(load.master_deal_list.txt_filter_primary_uw.field_name).AddItem "-1; _All"
            .Controls(load.master_deal_list.txt_filter_risk_type.field_name).AddItem "-1; _All"
            .Controls(load.master_deal_list.txt_filter_second_uw.field_name).AddItem "-1; _All"
            .Controls(load.master_deal_list.txt_filter_sub_sector.field_name).AddItem load.sub_sector_all & ";_All"
            .Controls(load.master_deal_list.txt_filter_super_sector.field_name).AddItem "-1; _All"
            .Controls(load.master_deal_list.txt_filter_was_quoted.field_name).AddItem "-1; _All"
            
            'add new values
            For i = 1 To arr_controls(0, 0)
                str_sql = arr_controls(i, 1)
                Set rs = utilities.create_adodb_rs(conn, str_sql)
                rs.Open
                Do While rs.EOF = False
                    If rs!menu_item = "_n/a" _
                    Or rs!menu_item = " " _
                    Or rs!menu_item = "" _
                    Or rs!menu_item = "_all" Then
                    Else
                        .Controls(arr_controls(i, 0)).AddItem rs!id & ";'" & rs!menu_item & "'"
                    End If
                    rs.MoveNext
                Loop
                rs.Close
            Next i
            
            'Color code deal names
            Dim objFormatConds As FormatCondition
            With .Controls("deal_name")
                .FormatConditions.delete
                
                Set objFormatConds = .FormatConditions.Add(acExpression, , "[deal_status_id] = 5")
                .FormatConditions(0).BackColor = load.colors.light_green
            
                Set objFormatConds = .FormatConditions.Add(acExpression, , "[deal_status_id] = 4")
                .FormatConditions(1).BackColor = load.colors.deal_status_preferred
            End With
        End With
    End If
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    open_forms.working_on_it_f__close
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "open_forms.deal_list_f"
        .milestone = "str_sql = " & str_sql
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub


