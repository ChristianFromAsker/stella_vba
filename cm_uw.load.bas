Attribute VB_Name = "load"
Option Compare Database
Option Explicit

'29 Jan 2025, CK: I tried putting the enums in global_vars, but something about the load sequence (I think) made it throw an error.
'enums
Public Enum menu_list
    yes = 93
    no = 94
End Enum
Public Enum question_types
    binder = 450
    country_proceure = 1000
    inar = 451
    auto_referral = 500
    confirmation = 520
    confirmation_uw_positions = 530
    uw_positions = 400
End Enum
Public Enum template_question_ids
    policy_period_question = 100
    question_re_having_reviewed_deal_questions_at_nbi_stage = 5
    question_re_having_reviewed_deal_questions_at_signing_stage = 6
End Enum

'These variables must load first, in this sequence.
Public sources As New cls_sources
Public conn As New ADODB.Connection

'objects
Public system_info As New cls_system
Public colors As New cls_colors
Public control_panel As New cls_control_panel
Public current_uw As New cls_underwriter
Public form_backgrounds As New cls_images
Public service_message_area As New cls_service_message_area
Public add_policy As New cls_form_add_policy_f
Public policies_f As New cls_form_policies_f

Private Type typ_form_names
    working_on_it As String
End Type
Public form_names As typ_form_names

'contants
'statuses (from stella_common. menu_list_t)
Private Type typ_deal_statuses
    nda_stage As Integer
    submission_stage As Integer
    quote_stage As Integer
    signed As Integer
    closed As Integer
End Type
Public deal_statuses As typ_deal_statuses

'country procedure statuses (from stella_common.
Private Type typ_country_procedure_statuses
    not_started As Integer
    cleared_internally As Integer
    in_process_with_third_party As Integer
    cleared As Integer
    not_needed_after_all As Integer
    failed As Integer
End Type
Public country_procedure_statuses As load.typ_country_procedure_statuses

Public Const quote_stage As Integer = 2
Public Const signed As Integer = 6
Public Const declined As Integer = 7
Public Const died As Integer = 8
Public Const lost As Integer = 9
Public Const closed As Integer = 436
Public Const collapsed As Integer = 485

'base values
Public Const placeholder_state As Integer = 442

'binder referral statuses
Public Const binder_referral_not_started = 521
Public Const cleared_internally As Integer = 486
Public Const in_process_with_carrier As Integer = 489
Public Const binder_referral_ok As Integer = 487
Public Const no_binder_referral As Integer = 488
Public Const binder_referral_denied As Integer = 490

'inar referrals
Public Const inar_referral_not_started = 498
Public Const inar_referral_ok As Integer = 500

'control_panel_f placements
'binder overview
Public Const first_binder_row_top = 1700
Public Const second_info_row_top = 5500
Public Const field_height As Long = 288
Public Const first_column_left As Long = 138
Public Const first_column_width As Long = 2268
Public second_column_left As Long
Public Const second_column_width As Long = 2041
Public third_column_left As Long
Public Const third_column_width As Long = 2041
Public fourth_column_left As Long
Public Const fourth_column_width As Long = 2835

'policy overview
Public Const policy_overview_left As Long = 10506

'arrays/collections to replace queries
Public binder_jurisdictions_array()
Public col_template_questions As New Collection
Public jurisdictions As New Collection
Public risk_types As New Collection

'other variables
Public binders_on_policy As Boolean
Public is_debugging As Boolean
Public is_init As Boolean
Public print_milestones As Boolean
Public underwriters() As Variant
Public has_searched As Boolean

'form movesize constants
Public Const second_view_right As Integer = 16000
Public Const binder_view_f_width As Integer = 19500

'misc
Public call_stack As String
Public debug_counter As Long
Public event_id As String
Public rs_counter As Integer
Public secondary_access_app As Access.Application


Public Sub check_secondary_access_app()
    load.call_stack = load.call_stack & vbNewLine & "load.check_secondary_access_app"
    
    'create access app for standby.
    Dim secondary_app_name As String, init_secondary_database As Boolean
    init_secondary_database = False
    If load.secondary_access_app Is Nothing Then
        'app object is destroyed and must be recreated.
        init_secondary_database = True
    Else
        'even if the app object exists, the actual app might be closed due to external influence.
        'Therefore, need to check if the app is available via a try-catch solution
        secondary_app_name = ""
        On Error Resume Next
            secondary_app_name = load.secondary_access_app.Name
        On Error GoTo err_handler
        If load.is_debugging = True Then On Error GoTo 0
        If secondary_app_name = "" Then
            init_secondary_database = True
        End If
    End If
    If init_secondary_database = True Then
        Set load.secondary_access_app = CreateObject("Access.Application")
        With load.secondary_access_app
            .OpenCurrentDatabase load.system_info.system_paths.common_path & "placeholder.accdb", False
            .Visible = False
        End With
    End If

outro:
    Exit Sub
    
err_handler:
    MsgBox "Something went wrong. Snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: load.check_secondary_access_app" & vbNewLine _
        & "Parameters: n/a" _
        & "App: " & load.system_info.app_name, , load.system_info.error_msg_heading
    GoTo outro
End Sub
Public Sub check_conn_and_variables()
    Const proc_name As String = "load.check_conn_and_variables"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim rs As ADODB.Recordset
    Dim str_milestone As String
    Dim str_sql As String
    
    str_milestone = "before load.system_info.init"
    load.system_info.init
    load.sources.init
    
    str_milestone = "before If conn Is Nothing Then load.init_conn"
    If conn Is Nothing Then load.init_conn
    
    'connections often drops while conn.state remain 1. The .close and .open seems to fix that.
    If conn.State <> adStateClosed Then
        conn.Close
    End If
    On Error GoTo conn_fix
    str_milestone = "before conn.Open"
    conn.Open
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    'populate service messages
    If CurrentProject.AllForms("control_panel_f").IsLoaded = True Then
        With load.service_message_area
            .init
            .populate_messages
            .place_boxes
        End With
    End If
    
    If load.is_init = False Then load.init_global_variables
    
outro:
    Exit Sub
    
conn_fix:
    load.init_conn
    Resume Next
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "load.check_conn_and_variables"
        .milestone = str_milestone
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    Resume outro
    
End Sub
Public Sub init_conn()
    Const proc_name As String = "load.init_conn"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim database_name As String
    Dim encrypted_string As String
    Dim encryption_key As Long
    Dim file_path As String
    Dim ip_address As String
    Dim pw_file As Long
    Dim pwd As String
    Dim server_name As String
    Dim str_conn As String
    Dim user_name As String
    
    'default values
    database_name = load.system_info.database_name
    encryption_key = 33
    ip_address = "mysql01-weu-prd.rpgroup.com"
    server_name = "public_avd_db"
    user_name = "stella"
    
    'get password
    pw_file = FreeFile
    
    file_path = load.system_info.system_paths.pws & "\" & server_name & ".txt"
    Open file_path For Input As FreeFile
        encrypted_string = Input(LOF(pw_file), pw_file)
        pwd = CStr(utilities.decrypt_string(encrypted_string, Left(encryption_key, 1), Right(encryption_key, 1)))
    Close pw_file
    
    'reset connection
    If Not load.conn Is Nothing Then Set load.conn = Nothing
    
    'activate connection

'1 September 2025, CK: We are moving to a new driver. This can be removed when done.
    On Error GoTo new_driver
    str_conn = "Driver={MySQL ODBC 8.0 Unicode Driver};" _
        & "Server=" & ip_address & ";" _
        & "DATABASE=" & database_name & ";" _
        & "UID=" & user_name & ";" _
        & "PWD=" & pwd
    Set load.conn = New ADODB.Connection
    
    With conn
        .ConnectionString = str_conn
        .Open
        .CursorLocation = adUseClient
    End With
    
    GoTo outro
    
new_driver:
'end remove

    str_conn = "Driver={MySQL ODBC 9.4 Unicode Driver};" _
        & "Server=" & ip_address & ";" _
        & "DATABASE=" & database_name & ";" _
        & "UID=" & user_name & ";" _
        & "PWD=" & pwd
    Set load.conn = New ADODB.Connection
    
    With conn
        .ConnectionString = str_conn
        .Open
        .CursorLocation = adUseClient
    End With
    
    GoTo outro
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    central.err_handler proc_name, Err.Number, Err.Description, "", "", "init_conn failed", True
    GoTo outro
End Sub
Public Sub init_global_variables()
    load.call_stack = load.call_stack & vbNewLine & "load.init_global_variables"
    is_init = True
    
    With load.form_names
        .working_on_it = "working_on_it_f"
    End With
    
    If load.system_info.is_init = False Then load.system_info.init
    load.sources.init
    load.add_policy.init
    load.init_binder_jurisdictions_array
    load.init_col_template_questions
    load.init_jurisdictions
    load.init_binder_risk_types
    load.current_uw.init_uw
    load.policies_f.init
    
    With load.control_panel
        If .is_init = False Then
            .init
        End If
    End With
    
    With load.deal_statuses
        .closed = 436
        .quote_stage = 2
        .signed = 6
        .submission_stage = 481
        .nda_stage = 1
    End With
    
    With load.country_procedure_statuses
        .cleared = 487
        .cleared_internally = 486
        .failed = 490
        .in_process_with_third_party = 489
        .not_needed_after_all = 488
        .not_started = 521
    End With
    
    'binder overview table on control_panel_f
    second_column_left = load.first_column_left + load.first_column_width
    third_column_left = load.first_column_left + load.first_column_width + load.second_column_width
    fourth_column_left = load.first_column_left + load.first_column_width + load.second_column_width + load.third_column_width
    
    global_vars.init
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "load.init_global_variables"
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
    
End Sub
Public Sub start_cm()
    load.rs_counter = 0
    load.call_stack = load.call_stack & vbNewLine & "load.start_cm"
    load.sources.init
    load.system_info.init
    
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    load.init_conn
    load.init_global_variables
    load.current_uw.init_uw
    'load.check_stella_version
    
    'make access window smaller (but not minimize)
    If load.is_debugging = False Then
        windows_apis.AccessMoveSize 0, 0, 350, 120
    End If
    
    Exit Sub
    
err_handler:
    MsgBox "Something went wrong. Snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: load.start_cm" & vbNewLine _
        & "Parameters: n/a" _
        & "App: " & load.system_info.app_name, , load.system_info.error_msg_heading

End Sub
Public Sub exit_cm()
    If Not colors Is Nothing Then Set colors = Nothing
    If Not current_uw Is Nothing Then Set current_uw = Nothing
    If Not sources Is Nothing Then Set sources = Nothing
    If Not system_info Is Nothing Then Set system_info = Nothing
    If Not load.add_policy Is Nothing Then Set load.add_policy = Nothing
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    If Not load.control_panel Is Nothing Then Set load.control_panel = Nothing
    If Not load.secondary_access_app Is Nothing Then
        load.secondary_access_app.Quit
        Set load.secondary_access_app = Nothing
    End If
End Sub
Public Sub init_binder_jurisdictions_array()
    load.call_stack = load.call_stack & vbNewLine & "load.inint_binder_jurisdictions_array"
    Dim rs As ADODB.Recordset, str_sql As String, i As Long
    
    str_sql = "SELECT id, binder_id, jurisdiction_id " _
    & " FROM " & load.sources.binder_jurisdictions_view _
    & " WHERE is_active = 93 " _
    & " ORDER BY id"
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        ReDim binder_jurisdictions_array(1 To CLng(rs.RecordCount), 0 To 2)
        i = 1
        Do Until rs.EOF = True
            binder_jurisdictions_array(i, 0) = rs!id
            binder_jurisdictions_array(i, 1) = rs!binder_id
            binder_jurisdictions_array(i, 2) = rs!jurisdiction_id
            i = i + 1
            rs.MoveNext
        Loop
    rs.Close
    Set rs = Nothing
End Sub
Public Sub init_col_template_questions()
    Dim proc_name As String
    proc_name = "load.init_col_template_questions"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    
    Dim counter As Integer
    Dim template_question As Scripting.Dictionary
    Dim str_sql As String, rs As ADODB.Recordset
    
    str_sql = "SELECT id, question, good_answer_id, default_answer_id, question_type_id__cm_question_types_t, for_eur_id, for_us_id" _
        & " FROM " & load.sources.template_qs_table _
        & " WHERE is_deleted = 0"
        
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        Set load.col_template_questions = Nothing
        With load.col_template_questions
            Do Until rs.EOF
                Set template_question = New Scripting.Dictionary
                With template_question
                    .Add "id", rs!id.value
                    .Add "question", rs!question.value
                    .Add "good_answer_id", rs!good_answer_id.value
                    .Add "default_answer_id", rs!default_answer_id.value
                    .Add "question_type_id", rs!question_type_id__cm_question_types_t.value
                    .Add "for_eur_id", rs!for_eur_id.value
                    .Add "for_us_id", rs!for_us_id.value
                End With
                .Add template_question
                rs.MoveNext
            Loop
        End With
    rs.Close
    
outro:
    If Not rs Is Nothing Then Set rs = Nothing
    Set rs = Nothing
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "load.init_template_questions"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub init_jurisdictions()
    load.call_stack = load.call_stack & vbNewLine & "load.inint_jurisdictions"
    Dim jurisdiction As Scripting.Dictionary, counter As Integer
    Dim str_sql As String, rs As ADODB.Recordset
    
    str_sql = "SELECT jurisdiction_id, jurisdiction, rp_region_id, jurisdiction_type, parent_jurisdiction_id" _
        & " FROM " & load.sources.jurisdictions_view _
        & " WHERE jurisdiction_type = 'country' OR jurisdiction_type = 'us_state'"
        
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        Set load.jurisdictions = Nothing
        With load.jurisdictions
            Do Until rs.EOF
                Set jurisdiction = New Scripting.Dictionary
                With jurisdiction
                    .Add "jurisdiction_id", rs!jurisdiction_id.value
                    .Add "jurisdiction", rs!jurisdiction.value
                    .Add "rp_region_id", rs!rp_region_id.value
                    .Add "jurisdiction_type", rs!jurisdiction_type.value
                    .Add "parent_jurisdiction_id", rs!parent_jurisdiction_id.value
                End With
                .Add jurisdiction
                rs.MoveNext
            Loop
        End With
    rs.Close
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "load.init_jurisdictions"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With

End Sub
Public Sub init_binder_risk_types()
    load.call_stack = load.call_stack & vbNewLine & "load.inint_bidner_risk_type"
    Dim risk_type As Scripting.Dictionary, counter As Integer
    Dim str_sql As String, rs As ADODB.Recordset
    
    str_sql = "SELECT id, binder_id, risk_type_id" _
        & " FROM " & load.sources.binder_risk_types_view
        
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        Set load.risk_types = Nothing
        With load.risk_types
            Do Until rs.EOF
                Set risk_type = New Scripting.Dictionary
                With risk_type
                    .Add "id", rs!id.value
                    .Add "binder_id", rs!binder_id.value
                    .Add "risk_type_id", rs!risk_type_id.value
                End With
                .Add risk_type
                rs.MoveNext
            Loop
        End With
    rs.Close
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "load.init_risk_types"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With

End Sub
