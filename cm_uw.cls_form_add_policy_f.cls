VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_form_add_policy_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public form_name As String

Public Sub init()
    form_name = "add_policy_f"
End Sub
Public Sub update_policy_breakdown()
    load.call_stack = load.call_stack & vbNewLine & "load.form_add_policy_f.update_policy_breakdown"
    With Forms(form_name)
        If !fundamental_limit > !layer_limit Then
            MsgBox "The fundamental limit needs to be a portion of the layer limit. Please adjust."
            !fundamental_limit = 0
        End If
        If !fundamental_premium * !quota > !policy_premium Then
            MsgBox "The fundamental premium needs to be a portion of the policy premium. Please adjust."
            !fundamental_premium = 0
        End If
        
        !overview_fundamental_limit = !fundamental_limit * !quota
        !overview_general_limit = (!layer_limit - !fundamental_limit) * !quota
        !overview_policy_limit = !layer_limit * !quota
        
        !overview_fundamental_premium = !fundamental_premium
        !overview_general_premium = !policy_premium - !fundamental_premium
        !overview_policy_premium = !policy_premium
    End With
End Sub
