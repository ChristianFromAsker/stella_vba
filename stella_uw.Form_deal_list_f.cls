VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_deal_list_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmd_all_statuses_Click()
    Const proc_name As String = "deal_list_f.cmd_all_statuses.Click"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    Dim form_field As cls_field
    Dim col As Collection
    Dim current_status As Long
    
    current_status = load.master_deal_list.cmd_status_nda.field_bg_color
    Set col = New Collection
    For Each form_field In load.master_deal_list.col_cmds_for_statuses_individual
        With form_field
            If current_status = load.colors.active_filter Then
                .field_bg_color = load.colors.inactive_filter
                .font_color = load.colors.black
                .field_caption = Split(.field_caption, ":")(0) & ": off"
            Else
                .field_bg_color = load.colors.active_filter
                .font_color = load.colors.white
                .field_caption = Split(.field_caption, ":")(0) & ": on"
            End If
        End With
        col.Add form_field
    Next form_field
    load.master_deal_list.cmd_date_filter.field_value = "create_date"
    col.Add load.master_deal_list.cmd_date_filter
    
    utilities.paint_control Me.Name, col
    
outro:
    load.call_stack = ""
    Exit Sub
err_handler:
    Resume outro
End Sub

Private Sub cmd_bound_Click()
    load.check_conn_and_variables
    Dim form_field As cls_field
    
    For Each form_field In load.master_deal_list.col_cmds_for_statuses_individual
        With form_field
            .field_bg_color = load.colors.inactive_filter
            .font_color = load.colors.black
            .field_caption = Split(.field_caption, ":")(0) & ": off"
        End With
    Next form_field
    
    With load.master_deal_list
        With .cmd_status_signed
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
        With .cmd_status_closed
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
        .cmd_date_filter.field_value = "inception_date"
    End With
    
    utilities.paint_control Me.Name, load.master_deal_list.col_cmds_for_statuses_individual
    
    load.master_deal_list.cmd_date_filter.field_value = "inception_date"
    cmd_date_filter.Caption = "filter by inception date"
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub cmd_date_filter_Click()
    load.call_stack = "deal_list_f.cmd_date_filter.click"
    load.check_conn_and_variables
    
    With load.master_deal_list.cmd_date_filter
        If .field_value = "create_date" Then
            .field_value = "inception_date"
            Me!cmd_date_filter.Caption = "Filtering by Inception Date"
        Else
            If load.system_info.app_continent = load.system_info.continents.americas Then
                Me!cmd_date_filter.Caption = .field_caption_us
            Else
                Me!cmd_date_filter.Caption = .field_caption
            End If
            .field_value = "create_date"
        End If
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub cmd_export_deals_Click()
    Dim proc_name As String
    proc_name = "deal_list_f.cmd_export_deals.click"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
    
    With load.system_info.texts
        open_forms.working_on_it_f .exporting_excels_heading, .exporting_excels_details, 6000
    End With
    
    If load.system_info.app_continent = load.system_info.continents.americas Then
        generate_reports.generate_master_list_extract Me.RecordSource, "us"
    Else
        generate_reports.generate_master_list_extract Me.RecordSource, "eur"
    End If
    open_forms.working_on_it_f__close
    
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

Private Sub cmd_export_policies_Click()
    Dim proc_name As String
    proc_name = "deal_list_f.cmd_export_policies.click"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.check_conn_and_variables
     
    With load.system_info.texts
        open_forms.working_on_it_f .exporting_excels_heading, .exporting_excels_details, 6000
    End With
    
    Dim str_sql As String
    str_sql = Forms!deal_list_f.RecordSource
    str_sql = Split(str_sql, "FROM")(0) & "FROM " & load.sources.global_policy_list_view & " WHERE" & Split(str_sql, "WHERE")(1)
    
    generate_reports.generate_policy_list_extract str_sql
    
    open_forms.working_on_it_f__close
    
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

Private Sub cmd_filter_buyer_Click()
    load.master_deal_list.change_parties_buttons load.master_deal_list.cmd_filter_buyer
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub cmd_filter_seller_Click()
    load.master_deal_list.change_parties_buttons load.master_deal_list.cmd_filter_seller
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub cmd_filter_target_Click()
    load.master_deal_list.change_parties_buttons load.master_deal_list.cmd_filter_target
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub cmd_non_bound_Click()
    load.check_conn_and_variables
    Dim form_field As cls_field
    
    For Each form_field In load.master_deal_list.col_cmds_for_statuses_individual
        With form_field
            .field_bg_color = load.colors.inactive_filter
            .font_color = load.colors.black
            .field_caption = Split(.field_caption, ":")(0) & ": off"
        End With
    Next form_field
    
    With load.master_deal_list
        With .cmd_status_declined
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
        With .cmd_status_cancelled
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
        With .cmd_status_lost
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
        With .cmd_status_collapsed
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
    End With
    
    utilities.paint_control Me.Name, load.master_deal_list.col_cmds_for_statuses_individual
    
    load.master_deal_list.cmd_date_filter.field_value = "create_date"
    cmd_date_filter.Caption = "filter by NDA date"
    
outro:
    load.call_stack = ""
    Exit Sub

End Sub

Private Sub cmd_open_workspace_Click()
    load.check_conn_and_variables
    
    With load.system_info.texts
        open_forms.working_on_it_f .opening_workspace_heading, .opening_workspace_details, 4000
    End With
    
    imanage.open_workspace Me!deal_id, load.current_uw.user_name & "@rpgroup.com"

outro:
    open_forms.working_on_it_f__close
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub cmd_qc_primary_or_xs_Click()
    Dim proc_name As String
    proc_name = "deal_list_f.cmd_qc_primary_or_xs.Click"
    load.call_stack = proc_name
    
    load.check_conn_and_variables
    
    With load.deal_details
        If .is_init = False Then .init Me!deal_id
        load.master_deal_list.quick_change .txt_primary_or_xs_id.field_name, "primary or xs", load.sources.menu_lists.layer_types, Me!deal_id
    End With

outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub cmd_qc_deal_status_Click()
    Dim proc_name As String
    proc_name = "deal_list_f.cmd_qc_deal_status.Click"
    load.call_stack = proc_name
    
    load.check_conn_and_variables
    
    load.master_deal_list.quick_change "deal_status_id", "deal status", load.sources.menu_lists.deal_statuses, Me!deal_id
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub cmd_qc_primary_uw_Click()
    Dim proc_name As String
    proc_name = "deal_list_f.cmd_qc_primary_uw.Click"
    load.call_stack = proc_name
    
    load.check_conn_and_variables
    
    With load.deal_details
        If .is_init = False Then .init Me!deal_id
        load.master_deal_list.quick_change .txt_primary_uw_id.field_name, "primary uw", load.sources.menu_lists.uws, Me!deal_id
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub
Private Sub cmd_qc_second_uw_Click()
    Dim proc_name As String
    proc_name = "deal_list_f.cmd_qc_second_uw.Click"
    load.call_stack = proc_name
    
    load.check_conn_and_variables
    
    load.master_deal_list.quick_change "secondary_uw", "second uw", load.sources.menu_lists.uws, Me!deal_id
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub cmd_regions_middle_east_Click()
    load.check_conn_and_variables
    load.master_deal_list.filter_cmd_change load.master_deal_list.cmd_regions_middle_east
outro:
    load.call_stack = ""
End Sub

Private Sub cmd_reset_Click()
    load.check_conn_and_variables
    With load.master_deal_list
        .default_filters
        .paint_master_deal_list
    End With
    
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub cmd_status_expensed_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_expensed
outro:
    load.call_stack = ""
End Sub

Private Sub cmd_submission_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_submission
outro:
    load.call_stack = ""
End Sub
Private Sub cmd_collapsed_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_collapsed
outro:
    load.call_stack = ""
End Sub

Private Sub cmd_working_Click()
    load.check_conn_and_variables
    Dim form_field As cls_field
    
    For Each form_field In load.master_deal_list.col_cmds_for_statuses_individual
        With form_field
            .field_bg_color = load.colors.inactive_filter
            .font_color = load.colors.black
            .field_caption = Split(.field_caption, ":")(0) & ": off"
        End With
    Next form_field
    
    With load.master_deal_list
        With .cmd_status_nda
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
        With .cmd_status_submission
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
        With .cmd_status_nbi
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
        With .cmd_status_preferred
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
        With .cmd_status_uw
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
        With .cmd_status_expensed
            .field_bg_color = load.colors.active_filter
            .font_color = load.colors.white
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End With
    End With
    
    utilities.paint_control Me.Name, load.master_deal_list.col_cmds_for_statuses_individual
    
    load.master_deal_list.cmd_date_filter.field_value = "create_date"
    cmd_date_filter.Caption = "filter by NDA date"
outro:
    load.call_stack = ""
End Sub
Private Sub fix_status_buttons(ByVal status_input As Variant, ByVal status_buttons As Variant)
    load.check_conn_and_variables
    'For status_buttons:
    '0 = control name
    '1 = default caption
    '2 = control type
    '3 = caption for on
    '4 = caption for off
    Dim i As Integer: i = 0
    For i = 1 To UBound(status_input)
        With Me.Controls(status_buttons(i, 0))
            If status_input(i) = 1 Then
                .Caption = status_buttons(i, 3)
                .BackColor = colors.active_filter
                .ForeColor = colors.white
            Else
                .Caption = status_buttons(i, 4)
                .BackColor = colors.inactive_filter
                .ForeColor = colors.black
            End If
        End With
    Next i
outro:
    load.call_stack = ""
End Sub

Private Sub deal_name_search_AfterUpdate()
    load.check_conn_and_variables
    load.master_deal_list.txt_filter_deal_name.field_value = Me!deal_name_search

outro:
    load.call_stack = ""
End Sub

Private Sub deal_notes_search_AfterUpdate()
    load.check_conn_and_variables
    load.master_deal_list.txt_filter_deal_notes.field_value = Me!deal_notes_search

outro:
    load.call_stack = ""
End Sub

Private Sub filter_broker_firm_AfterUpdate()
    Dim proc_name As String, field_name As String
    proc_name = "master_deal_list_f.filter_broker_firm.AfterUpdate"
    field_name = load.master_deal_list.txt_filter_broker_firm.field_name
    load.call_stack = proc_name
    load.check_conn_and_variables
    
    load.master_deal_list.txt_filter_broker_firm.field_value = Me.Controls(field_name)
    
outro:
    load.call_stack = ""
End Sub

Private Sub filter_date_month_from_AfterUpdate()
    load.check_conn_and_variables
    load.master_deal_list.txt_filter_date_month_from.field_value = Me.filter_date_month_from
outro:
    load.call_stack = ""
End Sub
Private Sub filter_date_month_to_AfterUpdate()
    load.check_conn_and_variables
    load.master_deal_list.txt_filter_date_month_to.field_value = Me.filter_date_month_to
outro:
    load.call_stack = ""
End Sub
Private Sub filter_date_year_from_AfterUpdate()
    load.check_conn_and_variables
    load.master_deal_list.txt_filter_date_year_from.field_value = Me.filter_date_year_from
outro:
    load.call_stack = ""
End Sub
Private Sub filter_date_year_to_AfterUpdate()
    load.check_conn_and_variables
    load.master_deal_list.txt_filter_date_year_to.field_value = Me.filter_date_year_to
outro:
    load.call_stack = ""
End Sub

Private Sub filter_insured_registered_country_AfterUpdate()
    Dim proc_name As String
    proc_name = "master_deal_list_f.filter_insured_registered_country.AfterUpdate"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    load.check_conn_and_variables
    
    With load.master_deal_list.txt_filter_insured_registered_country
        .field_value = Me.Controls(.field_name).Column(0)
    End With
End Sub

Private Sub filter_nbi_prepper_AfterUpdate()
    Dim proc_name As String
    proc_name = "master_deal_list_f.filter_nbi_prepper.AfterUpdate"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    load.check_conn_and_variables
    
    With load.master_deal_list.txt_filter_nbi_prepper
        .field_value = Me.Controls(.field_name).Column(0)
    End With
End Sub

Private Sub filter_super_sector_AfterUpdate()
    load.check_conn_and_variables
    If Me.filter_super_sector = 111 Then GoTo outro ' 111 is _All
        
    'populate com_sub_sector based on selection
    Dim str_sql As String, rs As ADODB.Recordset
    'remove existing items
    With Me.filter_sub_sector
        Do While .ListCount > 0
            .RemoveItem (0)
        Loop
        .AddItem load.sub_sector_all & ";_All"
        str_sql = "SELECT sector_id id, sector_name menu_item FROM " & sources.sectors_table & " WHERE parent_sector_id = " & Me.filter_super_sector & " ORDER BY sector_name"
        Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
            Do While rs.EOF = False
                .AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
        rs.Close
        .value = load.sub_sector_all
    End With
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    open_forms.working_on_it_f__close
    load.call_stack = ""
    Exit Sub
    
err_handler:
    MsgBox load.system_info.error_instruction & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: deal_details_f.target_super_sector_id_AfterUpdate()" & vbNewLine _
        & "Parameters: me!dela_id = " & Me!deal_id & vbNewLine _
        & "App: " & load.system_info.app_name, , load.system_info.error_msg_heading
    GoTo outro
End Sub

Private Sub filter_was_quoted_AfterUpdate()
    load.check_conn_and_variables
    load.master_deal_list.txt_filter_was_quoted.field_value = Me.filter_was_quoted
outro:
    load.call_stack = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    load.check_conn_and_variables
    Forms(load.master_deal_list.form_name).Visible = False
    If load.stella_is_closing = True Or load.master_deal_list.shall_be_refreshed = True Then
        Cancel = False
        GoTo outro
    End If
    If load.is_debugging = False Then Cancel = True

outro:
    load.call_stack = ""
    Exit Sub
End Sub
Private Sub Form_Load()
    DoCmd.MoveSize Right:=9250, Down:=0, Width:=21000, Height:=6000
End Sub
Private Sub cmd_open_folder_Click()
    load.call_stack = "deal_list_f.cmd_open_folder.Click"
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
Private Sub cmdOpenProject_Click()
    load.call_stack = "deal_list_f.cmdOpenProject.Click"
    load.check_conn_and_variables
    open_forms.deal_details_f Me!deal_id
    load.call_stack = ""
End Sub

Private Sub cmd_regions_usa_Click()
    load.check_conn_and_variables
    load.master_deal_list.filter_cmd_change load.master_deal_list.cmd_regions_usa
outro:
    load.call_stack = ""
End Sub
Private Sub cmd_regions_nordics_Click()
    load.check_conn_and_variables
    load.master_deal_list.filter_cmd_change load.master_deal_list.cmd_regions_nordics
outro:
    load.call_stack = ""
End Sub
Private Sub cmd_regions_canada_Click()
    load.check_conn_and_variables
    load.master_deal_list.filter_cmd_change load.master_deal_list.cmd_regions_canada
outro:
    load.call_stack = ""
End Sub
Private Sub cmd_regions_asiapac_Click()
    load.check_conn_and_variables
    load.master_deal_list.filter_cmd_change load.master_deal_list.cmd_regions_asiapac
outro:
    load.call_stack = ""
End Sub
Private Sub cmd_regions_dach_Click()
    load.check_conn_and_variables
    load.master_deal_list.filter_cmd_change load.master_deal_list.cmd_regions_dach
outro:
    load.call_stack = ""
End Sub
Private Sub cmd_regions_benelux_Click()
    load.check_conn_and_variables
    load.master_deal_list.filter_cmd_change load.master_deal_list.cmd_regions_benelux
outro:
    load.call_stack = ""
End Sub
Private Sub cmd_regions_med_Click()
    load.check_conn_and_variables
    load.master_deal_list.filter_cmd_change load.master_deal_list.cmd_regions_med
outro:
    load.call_stack = ""
End Sub
Private Sub cmd_regions_uk_Click()
    load.check_conn_and_variables
    load.master_deal_list.filter_cmd_change load.master_deal_list.cmd_regions_uk
outro:
    load.call_stack = ""
End Sub
Private Sub cmd_regions_all_Click()
    Const proc_name As String = "deal_list_f.cmd_regions_all.Click"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo err_handler
    
    load.check_conn_and_variables
    
    Dim form_field As cls_field
    Dim current_status As Boolean
    current_status = load.master_deal_list.cmd_regions_nordics.field_value
        
    For Each form_field In load.master_deal_list.col_cmds_for_regions_section
        With form_field
            If current_status = True Then
                .field_bg_color = load.colors.inactive_filter
                .font_color = load.colors.black
                .field_value = False
            Else
                .field_bg_color = load.colors.active_filter
                .font_color = load.colors.white
                .field_value = True
            End If
        End With
    Next form_field
    
    utilities.paint_control Me.Name, load.master_deal_list.col_cmds_for_regions_section
outro:
    load.call_stack = ""
    Exit Sub
err_handler:
    Resume Next
End Sub

Private Sub cmdDeclined_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_declined
outro:
    load.call_stack = ""
End Sub
Private Sub cmdCancel_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_cancelled
outro:
    load.call_stack = ""
End Sub
Private Sub cmdLost_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_lost
outro:
    load.call_stack = ""
End Sub
Private Sub cmdNBI_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_nbi
outro:
    load.call_stack = ""
End Sub
Private Sub cmdNDA_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_nda
outro:
    load.call_stack = ""
End Sub
Private Sub cmdPreferred_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_preferred
outro:
    load.call_stack = ""
End Sub
Private Sub cmdSigned_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_signed
outro:
    load.call_stack = ""
End Sub
Private Sub cmd_closed_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_closed
outro:
    load.call_stack = ""
End Sub
Private Sub cmdUW_Click()
    load.master_deal_list.status_buttons_colors load.master_deal_list.cmd_status_uw
outro:
    load.call_stack = ""
End Sub
Private Sub cmd_refresh_Click()
    Const proc_name As String = "deal_list_f.cmd_refresh.Click"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    load.call_stack = proc_name
    
    load.check_conn_and_variables
    load.master_deal_list.search
    
    Dim str_sql As String
    str_sql = "INSERT INTO " & load.sources.log_activity_table & "(activity, app_name, app_continent, user_name, activity_source)" _
    & " VALUES('searched from deal_list_f'" _
    & ", '" & load.system_info.app_name & "'" _
    & ", '" & load.system_info.app_continent & "'" _
    & ", '" & Environ("username") & "'" _
    & ", '" & proc_name & "')"
    load.conn.Execute str_sql
    
outro:
    load.call_stack = ""
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Private Sub fix_deal_list_order(ByVal data_field As String, Optional ByVal button_name As String)
    load.check_conn_and_variables
    Dim str_sql() As String, str_direction As String, control_name As String
    
    str_sql = Split(Me.RecordSource, "WHERE")
    str_sql = Split(str_sql(1), "ORDER")
    
    If button_name = "" Then
        control_name = "sort_" & data_field
    Else
        control_name = button_name
    End If
    
    If Me.Controls(control_name).Caption = "<<<" Then
        Me.Controls(control_name).Caption = ">>>"
        str_direction = "ASC"
    Else
        Me.Controls(control_name).Caption = "<<<"
        str_direction = "DESC"
    End If
    
    load.master_deal_list.update_data str_sql(0) & " ORDER BY " & data_field & " " & str_direction
outro:
    load.call_stack = ""
End Sub

Private Sub lbl_filter_buyer_info_Click()
    MsgBox "For 'buyer', searching will happen on Buyer Business Name, Ultimate Buyer and Insured Legal Name." _
    & vbNewLine & vbNewLine & "For 'seller', searching will happen on Seller Business Name and Ultimate Seller" _
    & vbNewLine & vbNewLine & "For 'target', searching will happen on Target Business Name and Target Legal Name", , "Filter info for parties"
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub lbl_filter_law_firm_info_Click()
    MsgBox "Will show deals where firm is either buy- or sell-side", , "Filter info for Law Firm"
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub lbl_filter_uw_info_Click()
    MsgBox "Will show deals where either UW is present", , "Filter info for UW"
        
outro:
    load.call_stack = ""
    Exit Sub
    
End Sub

Private Sub sort_broker_firm_Click()
    fix_deal_list_order "broker_firm"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_broker_person_Click()
    fix_deal_list_order "broker_person"
outro:
    load.call_stack = ""
End Sub

Private Sub sort_buyer_business_name_Click()
    fix_deal_list_order "buyer_business_name"
outro:
    load.call_stack = ""
End Sub

Private Sub sort_create_date_Click()
    fix_deal_list_order "create_date"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_deal_name_Click()
    fix_deal_list_order "deal_name"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_ev_Click()
    fix_deal_list_order "ev_eur", "sort_ev"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_inception_date_Click()
    fix_deal_list_order "inception_date"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_budget_home_Click()
    fix_deal_list_order "budget_home"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_deal_status_Click()
    load.check_conn_and_variables
    Dim sort_field As String
    sort_field = "deal_status"
    If load.system_info.app_continent = load.system_info.continents.americas Then
        sort_field = "deal_status_us"
    End If
    fix_deal_list_order sort_field, "sort_deal_status"
outro:
    load.call_stack = ""
End Sub

Private Sub sort_is_sanction_checks_done_Click()
    fix_deal_list_order "is_sanction_checks_done", "sort_is_sanction_checks_done"
outro:
    load.call_stack = ""
End Sub

Private Sub sort_limit_Click()
    fix_deal_list_order "total_rp_limit_on_deal_eur", "sort_limit"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_nbi_deadline_Click()
    fix_deal_list_order "nbi_deadline", "sort_nbi_deadline"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_premium_Click()
    fix_deal_list_order "total_rp_premium_on_deal_eur", "sort_premium"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_primary_or_xs_Click()
    fix_deal_list_order "primary_or_xs"
End Sub
Private Sub sort_primary_uw_Click()
    fix_deal_list_order "primary_uw_initials", "sort_primary_uw"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_risk_type_Click()
    fix_deal_list_order "risk_type"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_spa_law_Click()
    fix_deal_list_order "spa_law"
outro:
    load.call_stack = ""
End Sub
Private Sub sort_target_business_name_Click()
    fix_deal_list_order "target_business_name", load.master_deal_list.sort_target_business_name.field_name
outro:
    load.call_stack = ""
End Sub
