Attribute VB_Name = "generate_reports"
Option Compare Database
Option Explicit
Public Sub generate_test()
    Debug.Print "test was run"
    'this is called by calling app to test that macros can run
End Sub
Public Sub generate_us_policy_list(str_sql As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_us_policy_list"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.check_conn_and_variables
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_for_us_policy_export
    generate_reports.generate_report_old str_sql, report_frame
    
outro:
    Exit Sub
    
err_handler:
    GoTo outro
End Sub
Public Sub generate_master_deal_list_for_uw(str_sql As String, Optional app_continent As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_master_deal_list_for_uw"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    If app_continent <> "" Then
        Load.system_info.app_continent = app_continent
        Load.system_info.init_system_info
        Load.init_conn
    End If
    Load.check_conn_and_variables
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_for_master_deal_list_for_uw
    generate_reports.generate_report_old str_sql, report_frame
    
outro:
    Exit Sub
    
err_handler:
    GoTo outro
    
End Sub
Public Sub master_view_report_per_deal(str_sql As String, Optional app_continent As String)
    Dim proc_name As String
    proc_name = "generate_reports.master_view_report_per_deal"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    If app_continent <> "" Then
        Load.system_info.app_continent = app_continent
        Load.init_conn
    End If
    
    Load.check_conn_and_variables
    If is_debugging = True Then On Error GoTo 0
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_for_master_deal_export
    generate_reports.generate_report_old str_sql, report_frame
    
outro:
    Exit Sub
    
err_handler:
    GoTo outro
        
End Sub
Public Sub generate_cm_admin_extract(ByVal input_year As Long)
    Dim proc_name As String
    proc_name = "generate_reports.generate_cm_admin_extract"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.check_conn_and_variables
    If is_debugging = True Then On Error GoTo 0
    
    cm_admin.prepare_cm_admin_extract (input_year)
   
outro:
    Exit Sub
    
err_handler:
    GoTo outro
    
End Sub
Public Sub generate_global_deal_report(str_sql As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_global_deal_report"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.check_conn_and_variables
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_for_global_deal_list
    generate_reports.generate_report_old str_sql, report_frame
   
outro:
    Exit Sub
    
err_handler:
    GoTo outro
        
End Sub
Public Sub generate_eur_us_deals_report(str_sql As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_eur_us_deals_report"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.check_conn_and_variables
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_for_global_deal_list
    
    generate_reports.generate_report_old str_sql, report_frame
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
    
End Sub
Public Sub generate_weekly_view_extract(str_sql As String, ByVal app_continent As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_weekly_view_extract"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    If app_continent <> "" Then
        Load.system_info.app_continent = app_continent
        Load.system_info.init_system_info
        Load.init_conn
    End If
    
    Load.check_conn_and_variables
    
    Dim report_frame As Collection
    Set report_frame = generate_reports.report_frame_for_weekly_view
    generate_reports.generate_report str_sql, report_frame
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
    
End Sub
Public Sub generate_bound_report_per_deal(str_sql As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_bound_report_per_deal"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.check_conn_and_variables
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_per_deal
    generate_reports.generate_report_old str_sql, report_frame
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
    
End Sub
Public Sub generate_global_policy_report(str_sql As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_global_policy_report"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.check_conn_and_variables
    If is_debugging = True Then On Error GoTo 0
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_for_global_policy_list
    generate_reports.generate_report_old str_sql, report_frame

outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub generate_bound_report_per_policy(str_sql As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_bound_report_per_policy"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.check_conn_and_variables
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_per_policy
    generate_reports.generate_report_old str_sql, report_frame
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
    
End Sub
Public Sub generate_report_for_deals_with_missing_actions(str_sql As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_report_for_deals_with_missing_actions"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.check_conn_and_variables
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_for_deals_with_missing_actions
    generate_reports.generate_report_old str_sql, report_frame
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub generate_claims_report_per_policy(str_sql As String, Optional ByVal app_continent As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_claims_report_per_policy"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    Load.check_conn_and_variables
    If is_debugging = True Then On Error GoTo 0
    
    If app_continent <> "" Then
        Load.system_info.app_continent = app_continent
        Load.system_info.init_system_info
        Load.init_conn
    End If
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_for_claims_per_policy
    generate_reports.generate_report_old str_sql, report_frame
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
    
End Sub
Public Sub generate_claims_report_per_risk(str_sql As String, ByVal app_continent As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_claims_report_per_risk"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.system_info.app_continent = app_continent
    Load.system_info.init_system_info
    Load.init_conn
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_for_claims_per_risk
    generate_reports.generate_report_old str_sql, report_frame
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub generate_claims_report_per_carrier(str_sql As String, ByVal app_continent As String)
    Dim proc_name As String
    proc_name = "generate_reports.generate_claims_report_per_carrier"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.system_info.app_continent = app_continent
    Load.system_info.init_system_info
    Load.init_conn
    
    Dim report_frame() As Variant
    report_frame = generate_reports.report_frame_for_claims_per_carrier
    generate_reports.generate_report_old str_sql, report_frame
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub generate_report(ByVal str_sql As String, report_frame As Collection)
    Dim proc_name As String
    proc_name = "generate_reports.generate_report"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim arr_data() As Variant
    Dim arr_data_2() As Variant
    Dim arr_headings() As String
    Dim output As Variant
    Dim rs As ADODB.Recordset
    Dim row_count As Long
    Dim column_header As cls_field
    Dim column_count As Long
    Dim row_counter As Long
    Dim column_counter As Long
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        If .RecordCount = 0 Then
            'something
            GoTo outro
        End If
    
        row_count = CLng(rs.RecordCount)
        column_count = report_frame.Count
        
        'add data to the array
        ReDim arr_data(1 To row_count, 1 To column_count)
        .MoveFirst
        For row_counter = 1 To row_count
            column_counter = 1
            For Each column_header In report_frame
                output = ""
                
                If column_header.calculated_at_runtime = False Then
                    arr_data(row_counter, column_counter) = rs.Fields(column_header.field_name_deals_v)
                Else
                    If column_header.field_name_deals_v = "program_summary" Then
                        output = Nz(.Fields(column_header.field_name_deals_v), "")
                        output = Replace(output, "<br>", vbNewLine)
                    ElseIf column_header.field_name_deals_v = "highest_exposure_eur" Then
                        If !claimed_amount_eur - !lowest_rp_attpoint_eur < 0 Then
                            output = 0
                        Else
                            output = !claimed_amount_eur - !lowest_rp_attpoint_eur
                        End If
                    End If
                    arr_data(row_counter, column_counter) = output
                End If
                column_counter = 1 + column_counter
            Next column_header
            rs.MoveNext
        Next row_counter
        .Close
    End With
    Set rs = Nothing
    
    Dim ws As Excel.Worksheet
    Dim column_count_reduction
    column_count_reduction = 0
    
    '4 June 2025, CK: I suddenly got an 'out of memory error'. Turns out the error was caused by the fx_date data point. _
    This is just removed for now. The arr_data_2 was for error testing, but letting it be here for now.
    
    ReDim arr_data_2(1 To UBound(arr_data, 1), 1 To UBound(arr_data, 2) - column_count_reduction)
    For row_counter = 1 To UBound(arr_data, 1)
        For column_counter = 1 To UBound(arr_data, 2) - column_count_reduction
            arr_data_2(row_counter, column_counter) = arr_data(row_counter, column_counter)
        Next column_counter
    Next row_counter
    
    Erase arr_data
    
    Set ws = generate_reports.create_excel(report_frame, row_count)
    ReDim arr_headings(1 To report_frame.Count)
    column_counter = 1
    For Each column_header In report_frame
        arr_headings(column_counter) = column_header.field_caption
        column_counter = 1 + column_counter
    Next column_header
    With ws
        .Range(ws.Cells(1, 1), ws.Cells(1, column_count)) = arr_headings
        With .Range(ws.Cells(2, 1), ws.Cells(UBound(arr_data_2) + 1, column_count - column_count_reduction))
            .Value = arr_data_2
            .WrapText = True
            .EntireRow.AutoFit
            .Borders.Value = 1
        End With
    End With
    
outro:
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub generate_report_old(ByVal str_sql As String, report_frame() As Variant)
    Dim proc_name As String
    proc_name = "generate_reports.generate_report"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim arr_data() As Variant
    Dim arr_data_2() As Variant
    Dim arr_headings() As String
    Dim output As Variant
    Dim rs As ADODB.Recordset
    Dim row_count As Long
    Dim column_count As Long
    Dim row_counter As Long
    Dim column_counter As Long
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        If .RecordCount = 0 Then
            'something
            GoTo outro
        End If
    
        row_count = CLng(rs.RecordCount)
        
        'report_frame 0,0 is column count (assuming first is 1 and not 0)
        column_count = report_frame(0, 0)
        
        ReDim arr_headings(1 To 1, 1 To column_count + 1)
        For column_counter = 1 To column_count + 1
            arr_headings(1, column_counter) = report_frame(column_counter, 0)
        Next column_counter
        
        'add data to the array
        ReDim arr_data(1 To row_count, 1 To column_count)
        .MoveFirst
        For row_counter = 1 To row_count
            For column_counter = 1 To column_count
                output = ""
                'The fourth value of the report_frame dictates if a field is calculated. 1 means yes.
                If report_frame(column_counter, 4) = 0 Then
                    arr_data(row_counter, column_counter) = rs.Fields(report_frame(column_counter, 1))
                Else
                    If report_frame(column_counter, 1) = "ev_eur_band" Then
                        If IsNull(.Fields("ev_eur")) = False Then
                            'ev_eur is a calcualted field, so is returned as string from MySQL. Remove thousand seprators (,) and convert to currency.
                            Select Case CCur(Replace(.Fields("ev_eur"), ",", ""))
                            Case Is > 2000000000: output = "2000m<"
                            Case Is > 1000000000: output = "1000<2000m"
                            Case Is > 500000000: output = "500<1000m"
                            Case Is > 250000000: output = "250<500m"
                            Case Is > 150000000: output = "150<250m"
                            Case Is > 50000000: output = "50<150m"
                            Case Else: output = "0<50m"
                            End Select
                        End If
                    ElseIf report_frame(column_counter, 1) = "limit_eur_band" Then
                        If IsNull(.Fields("total_rp_limit_on_deal_eur")) = False Then
                            Select Case CCur(Replace(.Fields("total_rp_limit_on_deal_eur"), ",", ""))
                            Case Is > 100000000: output = ">100m"
                            Case Is > 80000000: output = ">80-100m"
                            Case Is > 60000000: output = ">60-80m"
                            Case Is > 40000000: output = ">40-60m"
                            Case Is > 20000000: output = ">20-40m"
                            Case Else: output = "0-20m"
                            End Select
                        End If
                    ElseIf report_frame(column_counter, 1) = "rol" Then
                        output = "=RC[-1] / RC[-3]"
                    ElseIf report_frame(column_counter, 1) = "program_summary" Then
                        If IsNull(.Fields(report_frame(column_counter, 1))) = False Then
                            output = Replace(.Fields(report_frame(column_counter, 1)), "<br>", vbNewLine)
                        End If
                    ElseIf report_frame(column_counter, 1) = "highest_exposure_eur" Then
                        If CLng(Replace(Nz(!claimed_amount_eur, 0), ",", "")) - !lowest_rp_attpoint_eur < 0 Then
                            output = 0
                        Else
                            output = CLng(Replace(Nz(!claimed_amount_eur, 0), ",", "")) - !lowest_rp_attpoint_eur
                        End If
                    End If
                    arr_data(row_counter, column_counter) = output
                End If
            Next column_counter
            rs.MoveNext
        Next row_counter
        .Close
    End With
    Set rs = Nothing
    
    Dim ws As Excel.Worksheet
    Dim column_count_reduction
    column_count_reduction = 0
    
    '4 June 2025, CK: I suddenly got an 'out of memory error'. Turns out the error was caused by the fx_date data point. _
    This is just removed for now. The arr_data_2 was for error testing, but letting it be here for now.
    
    ReDim arr_data_2(1 To UBound(arr_data, 1), 1 To UBound(arr_data, 2) - column_count_reduction)
    For row_counter = 1 To UBound(arr_data, 1)
        For column_counter = 1 To UBound(arr_data, 2) - column_count_reduction
            arr_data_2(row_counter, column_counter) = arr_data(row_counter, column_counter)
        Next column_counter
    Next row_counter
    
    Erase arr_data
    
    Set ws = generate_reports.create_excel_old(report_frame, row_count)
    With ws
        .Range(ws.Cells(1, 1), ws.Cells(1, column_count)) = arr_headings
        With .Range(ws.Cells(2, 1), ws.Cells(UBound(arr_data_2) + 1, column_count - column_count_reduction))
            .Value = arr_data_2
            .WrapText = True
            .EntireRow.AutoFit
            .Borders.Value = 1
        End With
    End With
    
outro:
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Function create_excel(col_data_fields As Collection, row_count As Long) As Excel.Worksheet
    Dim proc_name As String
    proc_name = "generate_reports.create_excel"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
        
    Dim app_excel As Excel.Application
    Dim column_counter As Long
    Dim column_count As Long
    Dim column_header As cls_field
    Dim str_milestone As String
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set app_excel = New Excel.Application
    app_excel.Visible = True
    Set wb = app_excel.Workbooks.Add
    app_excel.Calculation = xlCalculationManual
    wb.Activate
    Set ws = app_excel.Worksheets(1)
    str_milestone = "1e"
    app_excel.ActiveWindow.DisplayGridlines = False
    
    With ws
        column_counter = 1
        For Each column_header In col_data_fields
            With .Columns(column_counter)
                .ColumnWidth = column_header.excel_column_width
                .HorizontalAlignment = column_header.excel_horizontal_alignment
            End With
            column_counter = column_counter + 1
        Next column_header
        
        .Rows(1).Interior.Color = colors.black
        .Rows(1).Font.Color = colors.white
        .Rows(1).Font.Bold = True
        .Rows(1).WrapText = True
        .Rows(1).RowHeight = 40
        .Range(ws.Cells(1, 1), ws.Cells(row_count + 2, col_data_fields.Count)).Font.Name = "Montserrat"
        .Range(ws.Cells(1, 1), ws.Cells(row_count + 2, col_data_fields.Count)).Font.Size = 9
        .Rows("2:" & row_count + 2).RowHeight = 16
    End With
    
    '18 November 2024, CK: .visible is taking forever and does not seem necessary.
    '28 November 2024, CK: Now, without the below, the app would stay hidden, so I renabled it.
    '13 December 2024, CK: again, .visible is taking forever. I need to investigate .xml files. Trying to move the line up
    
    str_milestone = "4"
    app_excel.WindowState = xlMaximized
    Set create_excel = ws
    
outro:
    Exit Function
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = "str_milestine = " & str_milestone
        .params = "row_count = " & row_count
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
End Function
Public Function create_excel_old(report_frame() As Variant, row_count As Long) As Excel.Worksheet
    Dim proc_name As String
    proc_name = "generate_reports.create_excel_old"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_milestone As String
    
    Dim app_excel As Excel.Application
    Dim column_counter As Long
    Dim column_count As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    
    str_milestone = "1a"
    Set app_excel = New Excel.Application
    app_excel.Visible = True
    str_milestone = "1b"
    Set wb = app_excel.Workbooks.Add
    app_excel.Calculation = xlCalculationManual
    str_milestone = "1c"
    wb.Activate
    str_milestone = "1d"
    Set ws = app_excel.Worksheets(1)
    str_milestone = "1e"
    app_excel.ActiveWindow.DisplayGridlines = False
    
    str_milestone = "2"
    column_count = report_frame(0, 0)
    With ws
        'set default column properties
        For column_counter = 1 To column_count
            .Columns(column_counter).ColumnWidth = report_frame(column_counter, 2)
            .Columns(column_counter).HorizontalAlignment = xlLeft
            If report_frame(column_counter, 5) = True Then
                .Columns(column_counter).NumberFormat = "#,##0"
            End If
        Next column_counter
        
        'set deviating column properties
        'use if necessary
        str_milestone = "3a"
        .Rows(1).Interior.Color = colors.black
        str_milestone = "3b"
        .Rows(1).Font.Color = colors.white
        str_milestone = "3c"
        .Rows(1).Font.Bold = True
        str_milestone = "3d"
        .Rows(1).WrapText = True
        str_milestone = "3e"
        .Rows(1).RowHeight = 40
        str_milestone = "3f"
        .Range(ws.Cells(1, 1), ws.Cells(row_count + 2, column_count)).Font.Name = "Montserrat"
        str_milestone = "3g"
        .Range(ws.Cells(1, 1), ws.Cells(row_count + 2, column_count)).Font.Size = 9
        str_milestone = "3h"
        .Rows("2:" & row_count + 2).RowHeight = 16
    End With
    '18 November 2024, CK: .visible is taking forever and does not seem necessary.
    '28 November 2024, CK: Now, without the below, the app would stay hidden, so I renabled it.
    '13 December 2024, CK: again, .visible is taking forever. I need to investigate .xml files. Trying to move the line up
    
    str_milestone = "4"
    app_excel.WindowState = xlMaximized
    Set create_excel_old = ws
    
outro:
    Exit Function
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = "str_milestine = " & str_milestone
        .params = "row_count = " & row_count
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
End Function
Public Function report_frame_for_master_deal_export()
    Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Integer
    x = 1
    report_frame(x, 0) = "deal id"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "deal name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "broker firm"
    report_frame(x, 1) = "broker_firm"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target business name"
    report_frame(x, 1) = "target_business_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target legal jurisdiction"
    report_frame(x, 1) = "target_legal_jurisdiction"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target super sector"
    report_frame(x, 1) = "target_super_sector"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target sub sector"
    report_frame(x, 1) = "target_sub_sector"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "status"
    report_frame(x, 1) = "deal_status"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "was quoted?"
    report_frame(x, 1) = "was_quoted"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "budget home"
    report_frame(x, 1) = "budget_home"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "budget region"
    report_frame(x, 1) = "budget_region"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "SPA law"
    report_frame(x, 1) = "spa_law"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "primary or xs"
    report_frame(x, 1) = "primary_or_xs"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Inception date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal currency"
    report_frame(x, 1) = "deal_currency"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "EV (EUR)"
    report_frame(x, 1) = "ev_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Primary retention (EUR)"
    report_frame(x, 1) = "retention_eur"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Total RP limit (EUR)"
    report_frame(x, 1) = "total_rp_limit_on_deal_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Lowest RP attachment (EUR)"
    report_frame(x, 1) = "lowest_rp_attpoint_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Total RP premium (EUR)"
    report_frame(x, 1) = "total_rp_premium_on_deal_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Blended ROL"
    report_frame(x, 1) = "rol"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 1
    x = x + 1
    report_frame(x, 0) = "risk type"
    report_frame(x, 1) = "risk_type"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "buyer business names"
    report_frame(x, 1) = "buyer_business_name"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Insured registered country"
    report_frame(x, 1) = "insured_registered_country"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Ultimate Buyer"
    report_frame(x, 1) = "UltimateBuyer"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Seller Law Firm"
    report_frame(x, 1) = "seller_law_firm"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Buyer Law Firm 1"
    report_frame(x, 1) = "buyer_law_firm_1"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Buyer Law Firm 2"
    report_frame(x, 1) = "buyer_law_firm_2"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Submission date"
    report_frame(x, 1) = "submission_date"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Create date"
    report_frame(x, 1) = "create_date"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Comments"
    report_frame(x, 1) = "comments"
    report_frame(x, 2) = 30
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_for_master_deal_export = report_frame

End Function
Public Function report_frame_bound_us_deals_for_booking_team()
    Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Integer
    x = 1
    report_frame(x, 0) = "deal id"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "deal name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "broker firm"
    report_frame(x, 1) = "broker_firm"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target business name"
    report_frame(x, 1) = "target_business_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "status"
    report_frame(x, 1) = "deal_status"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "budget home"
    report_frame(x, 1) = "budget_home"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "SPA law"
    report_frame(x, 1) = "spa_law"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "primary or xs"
    report_frame(x, 1) = "primary_or_xs"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Inception date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal currency"
    report_frame(x, 1) = "deal_currency"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "EV (deal currency)"
    report_frame(x, 1) = "ev"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Initial retention"
    report_frame(x, 1) = "drop_start"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Dropped retention"
    report_frame(x, 1) = "drop_end"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Drop period (months)"
    report_frame(x, 1) = "drop_period"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Total RP limit"
    report_frame(x, 1) = "total_rp_limit_on_deal"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Lowest RP attachment"
    report_frame(x, 1) = "lowest_rp_attpoint"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Total RP premium"
    report_frame(x, 1) = "total_rp_premium_on_deal"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "risk type"
    report_frame(x, 1) = "risk_type"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "buyer business names"
    report_frame(x, 1) = "buyer_business_name"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Insured registered country"
    report_frame(x, 1) = "insured_registered_country"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Create date"
    report_frame(x, 1) = "create_date"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Comments"
    report_frame(x, 1) = "comments"
    report_frame(x, 2) = 30
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_bound_us_deals_for_booking_team = report_frame

End Function
Public Function report_frame_per_deal()
    Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Integer
    x = 1
    report_frame(x, 0) = "Deal ID"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Broker firm"
    report_frame(x, 1) = "broker_firm"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Target"
    report_frame(x, 1) = "target_business_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target super sector"
    report_frame(x, 1) = "target_super_sector"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target sub sector"
    report_frame(x, 1) = "target_sub_sector"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Status"
    report_frame(x, 1) = "deal_status"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "RP Entity"
    report_frame(x, 1) = "budget_home"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "SPA law"
    report_frame(x, 1) = "spa_law"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Layers"
    report_frame(x, 1) = "primary_or_xs"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Inception date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal currency"
    report_frame(x, 1) = "deal_currency"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "EV (EUR)"
    report_frame(x, 1) = "ev_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "EV (EUR) Band"
    report_frame(x, 1) = "ev_eur_band"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 1
    x = x + 1
    report_frame(x, 0) = "Primary retention (EUR)"
    report_frame(x, 1) = "retention_eur"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Lowest RP attachment (EUR)"
    report_frame(x, 1) = "lowest_rp_attpoint_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Total RP limit (EUR)"
    report_frame(x, 1) = "total_rp_limit_on_deal_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Lmit (EUR) Band"
    report_frame(x, 1) = "limit_eur_band"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 1
    x = x + 1
    report_frame(x, 0) = "Total RP premium (EUR)"
    report_frame(x, 1) = "total_rp_premium_on_deal_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Blended ROL"
    report_frame(x, 1) = "rol"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 1
    x = x + 1
    report_frame(x, 0) = "risk type"
    report_frame(x, 1) = "risk_type"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Buyer business names"
    report_frame(x, 1) = "buyer_business_name"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Insured registered country"
    report_frame(x, 1) = "insured_registered_country"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Ultimate Buyer"
    report_frame(x, 1) = "UltimateBuyer"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Create date"
    report_frame(x, 1) = "create_date"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Comments"
    report_frame(x, 1) = "comments"
    report_frame(x, 2) = 30
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 1
    
    
    report_frame(0, 0) = x
    
    report_frame_per_deal = report_frame

End Function
Public Function report_frame_for_global_policy_list() As Variant
    Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Integer
    x = 1
    report_frame(x, 0) = "Deal ID"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Policy No"
    report_frame(x, 1) = "policy_no"
    report_frame(x, 2) = 15
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Layer no"
    report_frame(x, 1) = "layer_no"
    report_frame(x, 2) = 6
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Layer group"
    report_frame(x, 1) = "layer_group"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Broker firm"
    report_frame(x, 1) = "broker_firm"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Target business name"
    report_frame(x, 1) = "target_business_name"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target super sector"
    report_frame(x, 1) = "target_super_sector"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target sub sector"
    report_frame(x, 1) = "target_sub_sector"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target sector group"
    report_frame(x, 1) = "target_sector_group"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Target country"
    report_frame(x, 1) = "target_jurisdiction"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Target region"
    report_frame(x, 1) = "target_region"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal Status"
    report_frame(x, 1) = "deal_status"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Parent Deal Status"
    report_frame(x, 1) = "parent_deal_status"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "budget home"
    report_frame(x, 1) = "budget_home"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "budget region"
    report_frame(x, 1) = "budget_region"
    report_frame(x, 2) = 15
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "budget continent"
    report_frame(x, 1) = "budget_continent"
    report_frame(x, 2) = 15
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "SPA law"
    report_frame(x, 1) = "spa_law"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "SPA law region"
    report_frame(x, 1) = "spa_law_region"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Primary or xs"
    report_frame(x, 1) = "primary_or_xs"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Inception date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Closing date"
    report_frame(x, 1) = "closing_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal currency"
    report_frame(x, 1) = "deal_currency"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    '4 June 2025, CK: This particular columns seems to cause memory error
'    x = x + 1
'    report_frame(x, 0) = "fx date"
'    report_frame(x, 1) = "fx_date"
'    report_frame(x, 2) = 12
'    report_frame(x, 3) = xlLeft
'    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "fx deal to EUR"
    report_frame(x, 1) = "fx_deal_eur"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "fx local to EUR"
    report_frame(x, 1) = "fx_local_eur"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "fx USD to EUR"
    report_frame(x, 1) = "fx_usd_eur"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "EV (EUR)"
    report_frame(x, 1) = "ev_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "primary retention"
    report_frame(x, 1) = "retention"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "primary retention (eur)"
    report_frame(x, 1) = "retention_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "underlying limit (eur)"
    report_frame(x, 1) = "underlying_limit_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "policy limit (eur)"
    report_frame(x, 1) = "policy_limit_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "fundamental limit (eur)"
    report_frame(x, 1) = "policy_fundamental_limit_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "general limit (eur)"
    report_frame(x, 1) = "policy_general_limit_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "lowest rp attachment (eur)"
    report_frame(x, 1) = "lowest_rp_attpoint_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "policy premium"
    report_frame(x, 1) = "policy_premium"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "policy premium (eur)"
    report_frame(x, 1) = "policy_premium_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "fundamental premium (eur)"
    report_frame(x, 1) = "fundamental_premium_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "general premium (eur)"
    report_frame(x, 1) = "general_premium_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "risk type"
    report_frame(x, 1) = "risk_type"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "major risk type"
    report_frame(x, 1) = "risk_type_major"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Buyer business names"
    report_frame(x, 1) = "buyer_business_name"
    report_frame(x, 2) = 26
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Create date"
    report_frame(x, 1) = "create_date"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    x = x + 1
    report_frame(x, 0) = "Navins home"
    report_frame(x, 1) = "navins_home"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    x = x + 1
    report_frame(x, 0) = "Issuing entity"
    report_frame(x, 1) = "issuing_entity"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_for_global_policy_list = report_frame
End Function
Public Function report_frame_per_policy() As Variant
    Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Integer
    x = 1
    report_frame(x, 0) = "Deal ID"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Policy No (C5)"
    report_frame(x, 1) = "policy_no"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Policy No (Stella)"
    report_frame(x, 1) = "stella_policy_no"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Broker firm"
    report_frame(x, 1) = "broker_firm"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Target business name"
    report_frame(x, 1) = "target_business_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target super sector"
    report_frame(x, 1) = "target_super_sector"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target sub sector"
    report_frame(x, 1) = "target_sub_sector"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Status"
    report_frame(x, 1) = "deal_status"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "budget home"
    report_frame(x, 1) = "budget_home"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "SPA law"
    report_frame(x, 1) = "spa_law"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Primary or xs"
    report_frame(x, 1) = "primary_or_xs"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Inception date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal currency"
    report_frame(x, 1) = "deal_currency"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "EV (EUR)"
    report_frame(x, 1) = "ev_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "EV (EUR) Band"
    report_frame(x, 1) = "ev_eur_band"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 1
    x = x + 1
    report_frame(x, 0) = "Primary retention"
    report_frame(x, 1) = "retention_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Layer no"
    report_frame(x, 1) = "layer_no"
    report_frame(x, 2) = 6
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Underlying limit (EUR)"
    report_frame(x, 1) = "underlying_limit_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Policy limit (EUR)"
    report_frame(x, 1) = "policy_limit_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Lowest RP attachment (EUR)"
    report_frame(x, 1) = "lowest_rp_attpoint_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "policy premium"
    report_frame(x, 1) = "policy_premium"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "policy premium (eur)"
    report_frame(x, 1) = "policy_premium_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "risk type"
    report_frame(x, 1) = "risk_type"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Buyer business names"
    report_frame(x, 1) = "buyer_business_name"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Ultimate Buyer"
    report_frame(x, 1) = "UltimateBuyer"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Create date"
    report_frame(x, 1) = "create_date"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Navins home"
    report_frame(x, 1) = "navins_home"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_per_policy = report_frame
End Function
Public Function report_frame_for_master_deal_list_for_uw() As Variant
 Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    'position 5: format as number
    Dim x As Integer
    x = 1
    report_frame(x, 0) = "NBIL Due Date"
    report_frame(x, 1) = "nbi_deadline_us"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Deal Name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Risk Type"
    report_frame(x, 1) = "risk_type"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Shop"
    report_frame(x, 1) = "broker_firm"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Person"
    report_frame(x, 1) = "broker_person"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Target Business Name"
    report_frame(x, 1) = "target_business_name"
    report_frame(x, 2) = 24
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Target Description"
    report_frame(x, 1) = "target_desc"
    report_frame(x, 2) = 24
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Buyer Business Name"
    report_frame(x, 1) = "buyer_business_name"
    report_frame(x, 2) = 24
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Deal Currency"
    report_frame(x, 1) = "deal_currency"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "EV"
    report_frame(x, 1) = "ev"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    report_frame(x, 5) = True
    
    x = 1 + x
    report_frame(x, 0) = "Max Limit Quoted"
    report_frame(x, 1) = "max_limit_quoted"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    report_frame(x, 5) = True
    
    x = 1 + x
    report_frame(x, 0) = "Total RP Premium on Deal (deal currency)"
    report_frame(x, 1) = "total_rp_premium_on_deal"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    report_frame(x, 5) = True
    
    x = 1 + x
    report_frame(x, 0) = "Buyer's Law Firm"
    report_frame(x, 1) = "buyer_law_firm_1"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "1st UW"
    report_frame(x, 1) = "primary_uw_initials"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "2nd UW"
    report_frame(x, 1) = "second_uw_initials"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Primary or xs"
    report_frame(x, 1) = "primary_or_xs"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Stage"
    report_frame(x, 1) = "stage"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Requote info"
    report_frame(x, 1) = "re_quote_info"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Submission Notes"
    report_frame(x, 1) = "submission_notes"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Deal Structure"
    report_frame(x, 1) = "deal_info"
    report_frame(x, 2) = 24
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Deal Status"
    report_frame(x, 1) = "deal_status"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Create Date"
    report_frame(x, 1) = "create_date"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    x = 1 + x
    report_frame(x, 0) = "Deal ID"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_for_master_deal_list_for_uw = report_frame
  
End Function

Public Function report_frame_for_deals_with_missing_actions() As Variant
    Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Integer
    x = 1
    report_frame(x, 0) = "deal id"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = 1 + 1
    report_frame(x, 0) = "deal name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "deal status"
    report_frame(x, 1) = "deal_status"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "budget home"
    report_frame(x, 1) = "budget_home"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "primary uw"
    report_frame(x, 1) = "primary_uw_full_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "secondary uw"
    report_frame(x, 1) = "second_uw_full_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "signing premium"
    report_frame(x, 1) = "signing_invoice_amount"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "signing premium received date"
    report_frame(x, 1) = "signing_premium_received_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "closing premium"
    report_frame(x, 1) = "closing_invoice_amount"
    report_frame(x, 2) = 18
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "closing premium received date"
    report_frame(x, 1) = "closing_premium_received_date"
    report_frame(x, 2) = 18
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "inception date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "emails filed"
    report_frame(x, 1) = "c5_emails_filed_hr"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "rr done"
    report_frame(x, 1) = "rr_done_hr"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "counsel invoice handled"
    report_frame(x, 1) = "counsel_invoice_handled_hr"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "closing booked"
    report_frame(x, 1) = "is_closing_booked"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "vdr received"
    report_frame(x, 1) = "vdr_received_hr"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_for_deals_with_missing_actions = report_frame
    
End Function
Public Function report_frame_for_us_policy_export() As Variant
    Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Long
      
    x = 1
    report_frame(x, 0) = "NBIL Due Date"
    report_frame(x, 1) = "nbi_deadline"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal Name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Risk Type"
    report_frame(x, 1) = "risk_type"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Shop"
    report_frame(x, 1) = "broker_firm"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Person"
    report_frame(x, 1) = "broker_person"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Target business name"
    report_frame(x, 1) = "target_business_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Target Description"
    report_frame(x, 1) = "target_desc"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Buyer Business Name"
    report_frame(x, 1) = "buyer_business_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal Currency"
    report_frame(x, 1) = "deal_currency"
    report_frame(x, 2) = 6
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "EV (deal currency)"
    report_frame(x, 1) = "ev"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Max Limit Requested (deal currency)"
    report_frame(x, 1) = "max_limit_quoted"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Policy Limit (deal currency)"
    report_frame(x, 1) = "policy_limit"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Policy Prremium (deal currency)"
    report_frame(x, 1) = "policy_premium"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Buyer's Law Firm"
    report_frame(x, 1) = "buyer_law_firm_1"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "1st UW"
    report_frame(x, 1) = "primary_uw"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "2nd UW"
    report_frame(x, 1) = "secondary_uw"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Primary or Xs"
    report_frame(x, 1) = "primary_or_xs"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Stage"
    report_frame(x, 1) = "stage"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Notes & Comments"
    report_frame(x, 1) = "comments"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Status"
    report_frame(x, 1) = "deal_status"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Create Date"
    report_frame(x, 1) = "create_date"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Inception Date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "policy premium"
    report_frame(x, 1) = "policy_premium"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Navins home"
    report_frame(x, 1) = "navins_home"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Policy No"
    report_frame(x, 1) = "policy_no"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal ID"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Policy ID"
    report_frame(x, 1) = "policy_id"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_for_us_policy_export = report_frame
    
End Function
Public Function report_frame_for_claims_per_policy() As Variant
    Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Long
    x = 1

    report_frame(x, 0) = "deal id"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "claim id"
    report_frame(x, 1) = "claim_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "deal name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Insured Legal Name"
    report_frame(x, 1) = "insured_legal_name"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "policy no"
    report_frame(x, 1) = "policy_no"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "broker firm"
    report_frame(x, 1) = "broker_firm"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "spa law"
    report_frame(x, 1) = "spa_law"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target jurisdiction"
    report_frame(x, 1) = "target_jurisdiction"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
     report_frame(x, 0) = "target super sector"
    report_frame(x, 1) = "target_super_sector"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target sub sector"
    report_frame(x, 1) = "target_sub_sector"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Buyer Business Name"
    report_frame(x, 1) = "buyer_business_name"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "insured jurisdiction"
    report_frame(x, 1) = "insured_jurisdiction"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "inception date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "ev (eur)"
    report_frame(x, 1) = "ev_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "claimed amount (eur)"
    report_frame(x, 1) = "claimed_amount_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "retention (eur)"
    report_frame(x, 1) = "retention_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "underlying limit (eur)"
    report_frame(x, 1) = "underlying_limit_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "policy limit (eur)"
    report_frame(x, 1) = "policy_limit_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "max loss exposure (eur)"
    report_frame(x, 1) = "max_exposure_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "rp estimate (eur)"
    report_frame(x, 1) = "estimated_loss_eur"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "indemnity paid (eur)"
    report_frame(x, 1) = "indemnity_paid_eur"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "final payment (eur)"
    report_frame(x, 1) = "final_loss_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "primary, xs or both?"
    report_frame(x, 1) = "primary_or_xs"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "notification date"
    report_frame(x, 1) = "ClaimDate"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "claim closed date"
    report_frame(x, 1) = "ClaimClosedDate"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "claim currency"
    report_frame(x, 1) = "claim_currency"
    report_frame(x, 2) = 10
    x = x + 1
    report_frame(x, 0) = "internal advisor fees (eur)"
    report_frame(x, 1) = "internal_advisor_fees_eur"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "client claim category 1"
    report_frame(x, 1) = "ClaimCategory1Client_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "client claim category 2"
    report_frame(x, 1) = "ClaimCategory2Client_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "client claim category 3"
    report_frame(x, 1) = "ClaimCategory3Client_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "client claim category 4"
    report_frame(x, 1) = "ClaimCategory4Client_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "relevant exclusion 1"
    report_frame(x, 1) = "RelevantExclusion1_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "relevant exclusion 2"
    report_frame(x, 1) = "RelevantExclusion2_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "relevant exclusion 3"
    report_frame(x, 1) = "RelevantExclusion3"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "relevant exclusion 4"
    report_frame(x, 1) = "RelevantExclusion4"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "risk likelihood"
    report_frame(x, 1) = "risk_likelihood_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "risk consequence"
    report_frame(x, 1) = "RiskFeel_hr"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "claim handler"
    report_frame(x, 1) = "claim_handler_hr"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "claim status"
    report_frame(x, 1) = "claim_status_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_for_claims_per_policy = report_frame
End Function
Public Function report_frame_for_claims_per_carrier() As Variant
    Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Long
    x = 1

    'deal info
    report_frame(x, 0) = "deal name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "policy no"
    report_frame(x, 1) = "policy_no"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "broker firm"
    report_frame(x, 1) = "broker_firm"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "spa law"
    report_frame(x, 1) = "spa_law"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target jurisdiction"
    report_frame(x, 1) = "target_jurisdiction"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "insured jurisdiction"
    report_frame(x, 1) = "insured_jurisdiction"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "inception date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "primary, xs or both?"
    report_frame(x, 1) = "primary_or_xs"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "ev (eur)"
    report_frame(x, 1) = "ev_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "retention (eur)"
    report_frame(x, 1) = "retention_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "underlying limit (eur)"
    report_frame(x, 1) = "underlying_limit_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "policy limit (eur)"
    report_frame(x, 1) = "policy_limit_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "carrier name"
    report_frame(x, 1) = "insurer_legal_name"
    report_frame(x, 2) = 24
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "carrier quota"
    report_frame(x, 1) = "carrier_quota"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "carrier limit (eur)"
    report_frame(x, 1) = "carrier_limit_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    
    'claims info
    report_frame(x, 0) = "claim currency"
    report_frame(x, 1) = "claim_currency"
    report_frame(x, 2) = 10
    x = x + 1
    report_frame(x, 0) = "total claimed amount (eur)"
    report_frame(x, 1) = "claimed_amount_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "max loss exposure for policy (eur)"
    report_frame(x, 1) = "total_max_exposure_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "carrier max exposure (eur)"
    report_frame(x, 1) = "carrier_max_exposure_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "rp total estimate (eur)"
    report_frame(x, 1) = "estimated_loss_eur"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "rp carrier estimate (eur)"
    report_frame(x, 1) = "carrier_estimated_loss_eur"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "internal advisor fees (eur)"
    report_frame(x, 1) = "internal_advisor_fees_eur"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "total final payment (eur)"
    report_frame(x, 1) = "final_loss_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "carrier final payment (eur)"
    report_frame(x, 1) = "carrier_final_loss_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "indemnity paid (eur)"
    report_frame(x, 1) = "indemnity_paid_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "interests paid (eur)"
    report_frame(x, 1) = "interests_paid_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "expenses paid (eur)"
    report_frame(x, 1) = "expenses_paid_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "total paid (eur)"
    report_frame(x, 1) = "total_paid_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "notification date"
    report_frame(x, 1) = "claim_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "claim closed date"
    report_frame(x, 1) = "claim_closed_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "client claim category 1"
    report_frame(x, 1) = "ClaimCategory1Client_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "client claim category 2"
    report_frame(x, 1) = "ClaimCategory2Client_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "client claim category 3"
    report_frame(x, 1) = "ClaimCategory3Client_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "client claim category 4"
    report_frame(x, 1) = "ClaimCategory4Client_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "relevant exclusion 1"
    report_frame(x, 1) = "RelevantExclusion1_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "relevant exclusion 2"
    report_frame(x, 1) = "RelevantExclusion2_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "relevant exclusion 3"
    report_frame(x, 1) = "relevant_exclusion_3"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "relevant exclusion 4"
    report_frame(x, 1) = "relevant_exclusion_4"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "risk likelihood"
    report_frame(x, 1) = "risk_likelihood_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "risk consequence"
    report_frame(x, 1) = "RiskFeel_hr"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "claim handler"
    report_frame(x, 1) = "claim_handler_hr"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    
    'admin
    report_frame(x, 0) = "deal id"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "claim id"
    report_frame(x, 1) = "claim_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "claim status"
    report_frame(x, 1) = "claim_status_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_for_claims_per_carrier = report_frame
End Function
Public Function report_frame_for_claims_per_risk() As Variant
    Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Long
    x = 1

    report_frame(x, 0) = "deal id"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "claim_id"
    report_frame(x, 1) = "claim_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Claim status"
    report_frame(x, 1) = "claim_status_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Broker firm"
    report_frame(x, 1) = "broker_firm"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "SPA law"
    report_frame(x, 1) = "spa_law"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target registered country"
    report_frame(x, 1) = "target_registered_country"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Insured registered country"
    report_frame(x, 1) = "insured_registered_country"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Inception Date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "EV (EUR)"
    report_frame(x, 1) = "ev_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    'ws_budget.Columns(i).NumberFormat = "#,###,##0"
    x = x + 1
    report_frame(x, 0) = "Primary retention (EUR)"
    report_frame(x, 1) = "retention_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    'ws_budget.Columns(i).NumberFormat = "#,###,##0"
    x = x + 1
    report_frame(x, 0) = "Lowest RP attachment (EUR)"
    report_frame(x, 1) = "lowest_rp_attpoint_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    'ws_budget.Columns(i).NumberFormat = "#,###,##0"
    x = x + 1
    report_frame(x, 0) = "Total limit on risk (EUR)"
    report_frame(x, 1) = "total_rp_limit_on_deal_eur"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    'ws_budget.Columns(i).NumberFormat = "#,###,##0"
    x = x + 1
    report_frame(x, 0) = "Claimed amount (EUR)"
    report_frame(x, 1) = "claimed_amount_eur"
    report_frame(x, 2) = 18
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    'ws_budget.Columns(i).NumberFormat = "#,###,##0"
    x = x + 1
    report_frame(x, 0) = "Highest exposure (EUR)"
    report_frame(x, 1) = "highest_exposure_eur"
    report_frame(x, 2) = 18
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 1
    'ws_budget.Columns(i).NumberFormat = "#,###,##0"
    x = x + 1
    report_frame(x, 0) = "RP Estimate (EUR)"
    report_frame(x, 1) = "estimated_loss_eur"
    report_frame(x, 2) = 18
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    'ws_budget.Columns(i).NumberFormat = "#,###,##0"
    x = x + 1
    report_frame(x, 0) = "Final Loss (EUR)"
    report_frame(x, 1) = "final_loss_eur"
    report_frame(x, 2) = 18
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    'ws_budget.Columns(i).NumberFormat = "#,###,##0"
    x = x + 1
    report_frame(x, 0) = "Primary, xs or both?"
    report_frame(x, 1) = "primary_or_xs"
    report_frame(x, 2) = 18
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    'ws_budget.Columns(i).NumberFormat = "#,###,##0"
    x = x + 1
    report_frame(x, 0) = "Notification date"
    report_frame(x, 1) = "ClaimDate"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Claim closed date"
    report_frame(x, 1) = "claim_closed_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Claim Currency"
    report_frame(x, 1) = "claim_currency"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Internal Advisor Fees"
    report_frame(x, 1) = "internal_advisor_fees"
    report_frame(x, 2) = 18
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    'ws_budget.Columns(i).NumberFormat = "#,###,##0"
    x = x + 1
    report_frame(x, 0) = "Target super sector"
    report_frame(x, 1) = "target_super_sector"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Target sub sector"
    report_frame(x, 1) = "target_sub_sector"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Claim category 1"
    report_frame(x, 1) = "claim_category_1_client_hr"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Claim category 2"
    report_frame(x, 1) = "claim_category_2_client_hr"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Claim category 3"
    report_frame(x, 1) = "claim_category_3_client_hr"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Claim category 4"
    report_frame(x, 1) = "claim_category_4_client_hr"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Relevant exclusion 1"
    report_frame(x, 1) = "relevant_exclusion_1_hr"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Relevant exclusion 2"
    report_frame(x, 1) = "relevant_exclusion_2_hr"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Relevant exclusion 3"
    report_frame(x, 1) = "relevant_exclusion_3"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Relevant exclusion 4"
    report_frame(x, 1) = "relevant_exclusion_4"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Risk consequence"
    report_frame(x, 1) = "risk_consequence"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Risk likelihood"
    report_frame(x, 1) = "risk_likelihood_hr"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Claim handler"
    report_frame(x, 1) = "claim_handler_hr"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Primary UW"
    report_frame(x, 1) = "primary_uw_full_name"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Second UW"
    report_frame(x, 1) = "second_uw_full_name"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Comments"
    report_frame(x, 1) = "Comments"
    report_frame(x, 2) = 48
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_for_claims_per_risk = report_frame
End Function
Public Function report_frame_for_weekly_view()
    Dim col_frame As New Collection
    
    Set col_frame = Nothing
    
    With global_vars.data_fields
        col_frame.Add .deal_id
        col_frame.Add .week_no_number
        col_frame.Add .deal_name
        col_frame.Add .risk_type
        col_frame.Add .ev_deal
        col_frame.Add .quote_due_date
        col_frame.Add .quote_due_time
        col_frame.Add .broker_firm
        col_frame.Add .target_desc
        col_frame.Add .target_business_name
        col_frame.Add .buyer_business_name
        col_frame.Add .is_repeat_buyer
        col_frame.Add .max_limit_quoted
        col_frame.Add .buyer_law_firm_1
        col_frame.Add .nbi_prepper_full_name
        col_frame.Add .primary_uw_full_name
        col_frame.Add .second_uw_full_name
        col_frame.Add .analyst_full_name
        col_frame.Add .budget_home
        col_frame.Add .deal_status
        col_frame.Add .requote_info
        col_frame.Add .submission_notes
        col_frame.Add .create_date
    End With
    
    Set report_frame_for_weekly_view = col_frame
    
End Function
Public Function report_frame_for_global_deal_list()
    Dim report_frame(100, 6) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Integer
    x = 1
    report_frame(x, 0) = "deal id"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "deal name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 20
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "policy count"
    report_frame(x, 1) = "policy_count"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "broker firm"
    report_frame(x, 1) = "broker_firm"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target super sector"
    report_frame(x, 1) = "target_super_sector"
    report_frame(x, 2) = 31
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target sub sector"
    report_frame(x, 1) = "target_sub_sector"
    report_frame(x, 2) = 27
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "target sector group"
    report_frame(x, 1) = "target_sector_group"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "deal status"
    report_frame(x, 1) = "deal_status"
    report_frame(x, 2) = 9
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "parent deal status"
    report_frame(x, 1) = "parent_deal_status"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "budget home"
    report_frame(x, 1) = "budget_home"
    report_frame(x, 2) = 15
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "budget region"
    report_frame(x, 1) = "budget_region"
    report_frame(x, 2) = 15
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "budget continent"
    report_frame(x, 1) = "budget_continent"
    report_frame(x, 2) = 15
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "SPA law"
    report_frame(x, 1) = "spa_law"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "SPA law region"
    report_frame(x, 1) = "spa_law_region"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "primary or xs"
    report_frame(x, 1) = "primary_or_xs"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Inception date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "SPA signing date"
    report_frame(x, 1) = "spa_signing_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Closing date"
    report_frame(x, 1) = "closing_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Deal currency"
    report_frame(x, 1) = "deal_currency"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
'    report_frame(x, 0) = "fx date"
'    report_frame(x, 1) = "fx_date"
'    report_frame(x, 2) = 12
'    report_frame(x, 3) = xlLeft
'    report_frame(x, 4) = 0
'    x = x + 1
    report_frame(x, 0) = "fx deal to EUR"
    report_frame(x, 1) = "fx_deal_eur"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "fx local to EUR"
    report_frame(x, 1) = "fx_local_eur"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "fx USD to EUR"
    report_frame(x, 1) = "fx_usd_eur"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "EV (EUR)"
    report_frame(x, 1) = "ev_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    x = x + 1
    report_frame(x, 0) = "Retention (EUR)"
    report_frame(x, 1) = "retention_eur"
    report_frame(x, 2) = 10
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    x = x + 1
    report_frame(x, 0) = "Total RP limit (EUR)"
    report_frame(x, 1) = "total_rp_limit_on_deal_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    x = x + 1
    report_frame(x, 0) = "Lowest RP attachment (EUR)"
    report_frame(x, 1) = "lowest_rp_attpoint_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Total RP premium (EUR)"
    report_frame(x, 1) = "total_rp_premium_on_deal_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Blended ROL"
    report_frame(x, 1) = "rol"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 1
    x = x + 1
    report_frame(x, 0) = "risk type"
    report_frame(x, 1) = "risk_type"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "major risk type"
    report_frame(x, 1) = "risk_type_major"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Insured registered country"
    report_frame(x, 1) = "insured_registered_country"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Target country"
    report_frame(x, 1) = "target_jurisdiction"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Target region"
    report_frame(x, 1) = "target_region"
    report_frame(x, 2) = 14
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Was quoted?"
    report_frame(x, 1) = "was_quoted"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Create date"
    report_frame(x, 1) = "create_date"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Sellside law firm"
    report_frame(x, 1) = "seller_law_firm"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Buyer law firm 1"
    report_frame(x, 1) = "buyer_law_firm_1"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Buyer law firm 2"
    report_frame(x, 1) = "buyer_law_firm_2"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "uw fee amount (EUR)"
    report_frame(x, 1) = "uw_fee_amount_eur"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "counsel fee amount (EUR)"
    report_frame(x, 1) = "counsel_fee_amount_eur"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "uw fee we keep (EUR)"
    report_frame(x, 1) = "uw_fee_we_keep_eur"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Signing Invoice Amount (EUR)"
    report_frame(x, 1) = "signing_invoice_amount_eur"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Total tower size"
    report_frame(x, 1) = "program_limit_eur"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_for_global_deal_list = report_frame

End Function
Public Function global_report_frame(ByVal report_type As Long) As Variant
    Dim report_frame(100, 4) As Variant
    'position 0: Column header in excel
    'Position 1: field name in recordset
    'Position 2: column width in excel
    'Position 3: Horizontal alignment in excel
    'Position 4: 0 for data field, 1 for being calculated
    Dim x As Integer
    x = 1
    report_frame(x, 0) = "deal id"
    report_frame(x, 1) = "deal_id"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = 1
    report_frame(x, 0) = "deal name"
    report_frame(x, 1) = "deal_name"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "primary uw"
    report_frame(x, 1) = "primary_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "secondary uw"
    report_frame(x, 1) = "secondary_hr"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlCenter
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "Total RP premium (EUR)"
    report_frame(x, 1) = "total_rp_premium_on_deal_eur"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "create date"
    report_frame(x, 1) = "create_date"
    report_frame(x, 2) = 16
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "inception date"
    report_frame(x, 1) = "inception_date"
    report_frame(x, 2) = 12
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "emails filed"
    report_frame(x, 1) = "c5_emails_filed_hr"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "rr done"
    report_frame(x, 1) = "rr_done_hr"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "counsel invoice handled"
    report_frame(x, 1) = "counsel_invoice_handled_hr"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "closed"
    report_frame(x, 1) = "is_closed_hr"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "closing booked"
    report_frame(x, 1) = "is_closing_booked_hr"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    x = x + 1
    report_frame(x, 0) = "vdr received"
    report_frame(x, 1) = "vdr_received_hr"
    report_frame(x, 2) = 8
    report_frame(x, 3) = xlLeft
    report_frame(x, 4) = 0
    
    report_frame(0, 0) = x
    
    report_frame_for_deals_with_missing_actions = report_frame
    
End Function
