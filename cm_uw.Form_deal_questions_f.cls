VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_deal_questions_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub cmd_change_Click()
    Const proc_name As String = "deal_questions_f.cmd_change.Click"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    load.check_conn_and_variables
    
    load.event_id = utilities.get_event_id
    If is_debugging = True Then On Error GoTo 0
    
    Dim answer As Integer
    Dim deal_id As Long
    Dim q_id As Long
    Dim q_number As Integer
    Dim str_sql As String
    
    deal_id = Me!header_deal_id
    
    'answer can only be yes or no, so change it to the other
    answer = 93
    If Me!answer_id = menu_list.yes Then answer = menu_list.no
    
    q_id = Me!id
    'figure out where to focus form after recordset update, to avoid form going to top
    q_number = Me.SelTop
    
    'save new data
    str_sql = "UPDATE " & sources.deal_questions_table & " SET answer_id = " & answer & " WHERE id = " & q_id
    conn.Execute str_sql
    load.rs_counter = load.rs_counter + 1
    
    'log change
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .change_source = proc_name
        .changer_id = Environ("username")
        .data_set = load.sources.deal_questions_table
        .deal_id = deal_id
        .field_name = "answer_id"
        .new_value = answer
        .record_id = q_id
    End With
    log_object.data_logger log_object
    Set log_object = Nothing
    
    'update recordset
    Dim rs_deal_questions As ADODB.Recordset
    
    str_sql = "SELECT * FROM " & sources.deal_questions_view _
    & " WHERE deal_id = " & deal_id _
    & " AND (" _
        & "question_type_id = " & question_types.binder _
        & " OR question_type_id = " & question_types.inar _
        & " OR question_type_id = " & question_types.confirmation _
    & ")" _
    & " ORDER BY sort_order, template_question_category"
    
    Set rs_deal_questions = utilities.create_adodb_rs(conn, str_sql)
    With rs_deal_questions
        .Open
        fix_rs.deal_questions_f deal_id, rs_deal_questions
        'fix_rs.deal_questions_f deal_id
        central.update_referrals_for_new_response_to_deal_question q_id, deal_id, rs_deal_questions
        fix_rs.deal_referrals_f deal_id
        If CurrentProject.AllForms("deal_referrals_f").IsLoaded = True Then
            open_forms.deal_referrals_f_resize
        End If
        central.check_status_area_on_control_panel_f deal_id, rs_deal_questions
        .Close
    End With
    Me.SelTop = q_number
    
outro:
    If Not rs_deal_questions Is Nothing Then
        If rs_deal_questions.State = 1 Then rs_deal_questions.Close
        Set rs_deal_questions = Nothing
    End If
    load.call_stack = ""
    load.event_id = ""
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "", "deal_id = " & deal_id, "", True
    Resume outro
End Sub
Private Sub cmd_expand_q_Click()
    MsgBox Me!question, , "Full question text"
End Sub
Private Sub Form_Load()
    DoCmd.MoveSize Right:=12500, Down:=3900, Width:=10000, Height:=10000
    'intro
    On Error GoTo err_handler
    load.check_conn_and_variables
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "deal_questions_f.Form.Load"
        .milestone = ""
        .params = "str_sql = " & ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Private Sub header_free_text_search_AfterUpdate()
    load.check_conn_and_variables
    With global_vars.deal
        If .is_init = False Then .init Forms("control_panel_f")!deal_id
    End With
    central.search_in_deal_questions_f global_vars.deal.deal_id
End Sub
Private Sub header_free_text_search_Click()
    Me!header_free_text_search = "hit enter to search"
End Sub
Private Sub sort_by_category_Click()
    load.check_conn_and_variables
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_sql_temp() As String, str_sql As String, data_field As String, str_control As String
    str_control = "sort_by_category"
    data_field = "template_question_category"
    str_sql_temp = Split(Me.RecordSource, "WHERE")
    str_sql_temp = Split(str_sql_temp(1), "ORDER")
    Dim str_direction As String
    If Me.Controls(str_control).Caption = "<<<" Then
        Me.Controls(str_control).Caption = ">>>"
        str_direction = "ASC"
    Else
        Me.Controls(str_control).Caption = "<<<"
        str_direction = "DESC"
    End If
    
    str_sql = "SELECT * FROM " & sources.deal_questions_view _
    & " WHERE " & str_sql_temp(0) _
    & " AND (" _
        & "question_type_id = " & question_types.binder _
        & " OR question_type_id = " & question_types.inar _
        & " OR question_type_id = " & question_types.confirmation _
    & ")" _
    & " ORDER BY " & data_field & " " & str_direction
    
    Dim rs As ADODB.Recordset
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        Set Me.Recordset = rs
    rs.Close
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "deal_questions_f.sort_by_category.Click"
        .milestone = ""
        .params = "str_sql = " & ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Private Sub sort_by_create_date_Click()
    load.check_conn_and_variables
    Dim str_sql_temp() As String, str_sql As String, data_field As String, str_control As String
    str_control = "sort_by_create_date"
    data_field = "create_date"
    str_sql_temp = Split(Me.RecordSource, "WHERE")
    str_sql_temp = Split(str_sql_temp(1), "ORDER")
    Dim str_direction As String
    If Me.Controls(str_control).Caption = "<<<" Then
        Me.Controls(str_control).Caption = ">>>"
        str_direction = "ASC"
    Else
        Me.Controls(str_control).Caption = "<<<"
        str_direction = "DESC"
    End If
    str_sql = "SELECT * FROM " & sources.deal_questions_view & " WHERE " & str_sql_temp(0) & " ORDER BY " & data_field & " " & str_direction
    Dim rs As ADODB.Recordset
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        Set Me.Recordset = rs
    rs.Close
End Sub

