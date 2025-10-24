Attribute VB_Name = "Paths"
Option Compare Database
Option Explicit
Public Sub fix_folder_backlog()
    Dim proc_name As String
    proc_name = "paths.fix_folder_backlog"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim fso As Object
    Dim rs As ADODB.Recordset
    Dim str_normative_path As String
    Dim str_sql As String
    
    str_sql = "SELECT * FROM " & load.sources.folder_moving_table _
    & " WHERE is_deleted = 0" _
    & " AND is_moved = 'no'"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        If CLng(.RecordCount) > 0 Then
            str_normative_path = ""
            '20 August 2025 CK YOU GOT HERE!
            On Error Resume Next
            'fso.movefolder str_actual_path, str_normative_path
            If load.is_debugging = True Then On Error GoTo 0
        
        End If
        .Close
    End With
    Set rs = Nothing
    Set fso = Nothing
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
Public Sub open_deal_folder(ByVal deal_id As Long)
    Dim proc_name As String
    proc_name = "paths.open_deal_folder"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim missing_fields As New Collection
    Dim str_milestone As String
    Dim str_path As String
    
    str_milestone = "0"
    
    Set missing_fields = mandatory_data.check_data(deal_id, load.mandatory_tests.folder_check)
    If missing_fields.Count > 0 Then
        MsgBox "It appears the Deal Folder for this deal is not created. Reach out to Tom or Christian if you think that's wrong" _
        & vbNewLine & vbNewLine & "For Tom and Christian:" & vbNewLine & "stella_uw.paths.open_deal_folder" _
        , , "Cannot open deal folder for " & deal_id
        GoTo outro
    End If
    
    str_milestone = "1"
    
    str_path = Paths.folder_path_from_scripts(deal_id, "find")
    If str_path = "-1" Or str_path = "" Then
        MsgBox "I've looked in the basement, and attic, and all the storage rooms, but I simply cannot find the deal folder." & vbNewLine & vbNewLine _
        & "Maybe the bats ate it, or someone stole it. Or misplaced it.", , "No deal folder!"
        GoTo outro
    End If
    
    str_milestone = "2"
    Application.FollowHyperlink str_path
    str_milestone = "3"
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .call_stack = load.call_stack
        .routine_name = proc_name
        .milestone = str_milestone
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro

End Sub

Public Function folder_path_from_scripts(ByVal deal_id As Long, normative_or_find As String) As String
    Dim proc_name As String
    proc_name = "paths.folder_path_from_scripts"
    load.call_stack = load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_milestone As String
    Dim str_folder_path As String
    
    load.check_secondary_access_app
    Central.open_external_resource_app "scripts.accdb", False, load.system_info.system_paths.common_path & "scripts.accdb"
    
    With load.secondary_access_app
        If normative_or_find = "normative" Then
            If load.system_info.app_continent = load.system_info.continents.americas Then
                str_folder_path = CStr(.Run("create_normative_folder_path_us", deal_id, load.system_info.app_continent))
            ElseIf load.system_info.app_continent = load.system_info.continents.eurasia Then
                str_folder_path = CStr(.Run("create_normative_folder_path", deal_id, load.system_info.app_continent))
            End If
        ElseIf normative_or_find = "find" Then
            str_folder_path = CStr(.Run("scripts.find_folder", deal_id, load.system_info.app_continent))
        End If
            
        str_milestone = "4"
        .CloseCurrentDatabase
        
        str_milestone = "5"
        .OpenCurrentDatabase load.system_info.system_paths.common_path & "placeholder.accdb", False
    End With
    
    str_milestone = "6"
    folder_path_from_scripts = str_folder_path
    
outro:
    load.secondary_access_app.Visible = False
    Exit Function
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .call_stack = load.call_stack
        .routine_name = proc_name
        .milestone = "str_milestone = " & str_milestone
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
        folder_path_from_scripts = "-1"
    GoTo outro
End Function


