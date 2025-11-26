VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_form_deal_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'headers
Public header_nda_stage As New cls_field
Public header_deal_and_target_info As New cls_field
Public header_advisors As New cls_field
Public header_submission_notes As New cls_field
Public header_insurance_info As New cls_field
Public header_buyer_and_seller As New cls_field
Public header_closing_info As New cls_field
Public header_admin As New cls_field
Public header_financial_info As New cls_field
Public header_misc As New cls_field
Public col_headers As New Collection

'labels
    'nda stage
    Public lbl_deal_status_id As New cls_field
    Public lbl_risk_name As New cls_field
    Public lbl_deal_id As New cls_field
    Public lbl_broker_firm_id As New cls_field
    Public lbl_broker_person_id As New cls_field
    Public lbl_spa_law_id As New cls_field
    Public lbl_stage_id As New cls_field
    Public lbl_max_limit_quoted As New cls_field
    Public lbl_quote_due_date As New cls_field
    Public lbl_was_quoted_id As New cls_field
    Public lbl_is_test_deal_id As New cls_field
    
    'deal and target info
    Public lbl_risk_currency_id As New cls_field
    Public lbl_ev As New cls_field
    Public lbl_target_legal_name As New cls_field
    Public lbl_target_business_name As New cls_field
    Public lbl_target_desc As New cls_field
    Public lbl_target_registered_country_id As New cls_field
    Public lbl_target_main_country_id As New cls_field
    Public lbl_target_super_sector_id As New cls_field
    Public lbl_target_sub_sector_id As New cls_field
    Public lbl_target_sic_code As New cls_field
    Public lbl_target_naic_code As New cls_field
    Public lbl_risk_signing_date As New cls_field
    Public lbl_risk_closing_date As New cls_field
    
    'advisors
    Public lbl_seller_law_firm_id As New cls_field
    Public lbl_buyer_law_firm_1_id As New cls_field
    Public lbl_buyer_law_firm_2_id As New cls_field
    Public lbl_buyer_financial_advisor_id As New cls_field
    Public lbl_legal_uw_counsel_firm_id As New cls_field
    Public lbl_uw_counsel_1_id As New cls_field
    Public lbl_financial_uw_counsel_firm_id As New cls_field
    Public lbl_uw_counsel_2_id As New cls_field
    
    'insurance info
    Public lbl_risk_type_id As New cls_field
    Public lbl_nbi_prepper As New cls_field
    Public lbl_primary_uw_id As New cls_field
    Public lbl_secondary_uw_id As New cls_field
    Public lbl_analyst_id As New cls_field
    Public lbl_internal_approver_quote_id As New cls_field
    Public lbl_internal_approver_binding_id As New cls_field
    Public lbl_risk_feel_id As New cls_field
    Public lbl_primary_or_xs_id As New cls_field
    Public lbl_budget_home_id As New cls_field
    Public lbl_program_limit As New cls_field
    Public lbl_primary_insurer_id As New cls_field
    Public lbl_retention As New cls_field
    Public lbl_retention_dropped As New cls_field
    Public lbl_drop_period As New cls_field
    Public lbl_inception_date As New cls_field
    Public lbl_policy_period As New cls_field
    
    'buyer and seller info
    Public lbl_ultimate_seller As New cls_field
    Public lbl_seller_business_name As New cls_field
    Public lbl_ultimate_buyer As New cls_field
    Public lbl_buyer_business_name As New cls_field
    Public lbl_insured_legal_name As New cls_field
    Public lbl_insured_registered_country_id As New cls_field
    Public lbl_insured_main_region_id As New cls_field
        
    'submission notes
    Public lbl_reviewer_notes As New cls_field
    Public lbl_deal_info As New cls_field
    Public lbl_submission_limits As New cls_field
    
    'financial info
    Public lbl_amounts As New cls_field
    Public lbl_due_dates As New cls_field
    Public lbl_received_dates As New cls_field
    Public lbl_uw_fee As New cls_field
    Public lbl_counsel_fee As New cls_field
    Public lbl_net_uw_fee As New cls_field
    Public lbl_signing_premium As New cls_field
    Public lbl_closing_premium As New cls_field
    Public lbl_total_premium As New cls_field
    
    'misc section
    Public lbl_surplus_broker_firm As New cls_field
    Public lbl_surplus_broker_person As New cls_field
    Public lbl_surplus_broker_license_no As New cls_field
    Public lbl_us_producer_no As New cls_field
    Public lbl_is_underwritten_id As New cls_field
    Public lbl_comments As New cls_field
    Public lbl_stella_message_board As New cls_field
    
    'closing info
    Public lbl_closing_set_received As New cls_field
    Public lbl_esg_country As New cls_field
    Public lbl_esg_company As New cls_field
    Public lbl_sanction_checks As New cls_field
    Public lbl_is_counsel_invoice_received_id As New cls_field
    Public lbl_is_vdr_received_id As New cls_field
    Public lbl_vdr_password As New cls_field
    Public lbl_is_rr_completed_id As New cls_field
    Public lbl_are_emails_filed_id As New cls_field
    Public lbl_is_closing_booked_id As New cls_field
    
    'admin
    Public lbl_create_date As New cls_field
    Public lbl_premium_eur As New cls_field
    Public lbl_premium_local As New cls_field
    Public lbl_local_fx_to_eur As New cls_field
    Public lbl_deal_fx_to_eur As New cls_field
    Public lbl_currency_date As New cls_field
    Public lbl_vat_home As New cls_field
    Public lbl_creating_uw As New cls_field

'text fields
    'nda stage
    Public txt_deal_status_id As New cls_field
    Public txt_risk_name As New cls_field
    Public txt_deal_id As New cls_field
    Public txt_broker_firm_id As New cls_field
    Public txt_broker_person As New cls_field
    Public txt_spa_law As New cls_field
    Public txt_stage_id As New cls_field
    Public txt_max_limit_quoted As New cls_field
    Public txt_quote_due_date As New cls_field
    Public txt_was_quoted_id As New cls_field
    Public txt_is_test_deal_id As New cls_field
    
    'deal and target info
    Public txt_risk_currency_id As New cls_field
    Public txt_ev As New cls_field
    Public txt_target_legal_name As New cls_field
    Public txt_target_business_name As New cls_field
    Public txt_target_desc As New cls_field
    Public txt_target_registered_country_id As New cls_field
    Public txt_target_main_country_id As New cls_field
    Public txt_target_super_sector_id As New cls_field
    Public txt_target_sub_sector_id As New cls_field
    Public txt_target_sic_code As New cls_field
    Public txt_target_naic_code As New cls_field
    Public txt_risk_signing_date As New cls_field
    Public txt_risk_closing_date As New cls_field
    
    'advisors
    Public txt_seller_law_firm_id As New cls_field
    Public txt_buyer_law_firm_1_id As New cls_field
    Public txt_buyer_law_firm_2_id As New cls_field
    Public txt_buyer_financial_advisor_id As New cls_field
    Public txt_uw_law_firm_id As New cls_field
    Public txt_uw_counsel_1_id As New cls_field
    Public txt_uw_financial_advisor_id As New cls_field
    Public txt_uw_counsel_2_id As New cls_field
    
    'insurance info
    Public txt_risk_type_id As New cls_field
    Public txt_nbi_prepper_id As New cls_field
    Public txt_nbi_prepper_full_name As New cls_field
    Public txt_primary_uw_id As New cls_field
    Public txt_primary_uw_full_name As New cls_field
    Public txt_secondary_uw_id As New cls_field
    Public txt_secondary_uw_full_name As New cls_field
    Public txt_analyst_id As New cls_field
    Public txt_analyst_full_name As New cls_field
    Public txt_internal_approver_quote_id As New cls_field
    Public txt_internal_approver_quote_full_name As New cls_field
    Public txt_internal_approver_binding_id As New cls_field
    Public txt_internal_approver_binding_full_name As New cls_field
    Public txt_risk_feel_id As New cls_field
    Public txt_primary_or_xs_id As New cls_field
    Public txt_budget_home_id As New cls_field
    Public txt_program_limit As New cls_field
    Public txt_primary_insurer_id As New cls_field
    Public txt_retention As New cls_field
    Public txt_retention_dropped As New cls_field
    Public txt_drop_period As New cls_field
    Public txt_inception_date As New cls_field
    Public txt_policy_period As New cls_field
    
    'buyer and seller info
    Public txt_ultimate_seller As New cls_field
    Public txt_seller_business_name As New cls_field
    Public txt_ultimate_buyer As New cls_field
    Public txt_buyer_business_name As New cls_field
    Public txt_insured_legal_name As New cls_field
    Public txt_insured_registered_country_id As New cls_field
    Public txt_insured_main_region_id As New cls_field
    
    'submission notes
    Public txt_reviewer_notes As New cls_field
    Public txt_deal_info As New cls_field
    Public txt_submission_limits As New cls_field
    
    'financial info
    Public txt_uw_fee As New cls_field
    Public txt_counsel_fee As New cls_field
    Public txt_net_uw_fee As New cls_field
    Public txt_uw_fee_due_date As New cls_field
    Public txt_uw_fee_received_date As New cls_field
    Public txt_counsel_fee_due_date As New cls_field
    Public txt_counsel_fee_received_date As New cls_field
    Public txt_signing_premium_amount As New cls_field
    Public txt_signing_premium_due_date As New cls_field
    Public txt_signing_premium_received_date As New cls_field
    Public txt_closing_premium_amount As New cls_field
    Public txt_closing_premium_due_date As New cls_field
    Public txt_closing_premium_received_date As New cls_field
    Public txt_total_premium As New cls_field
    
    'check points
    Public txt_closing_set_received As New cls_field
    Public txt_esg_country As New cls_field
    Public txt_esg_company As New cls_field
    Public txt_sanction_checks As New cls_field
    Public txt_is_counsel_invoice_received_id As New cls_field
    Public txt_is_vdr_received_id As New cls_field
    Public txt_vdr_password As New cls_field
    Public txt_is_rr_completed_id As New cls_field
    Public txt_are_emails_filed_id As New cls_field
    Public txt_is_closing_booked_id As New cls_field
    
    'admin
    Public txt_create_date As New cls_field
    Public txt_premium_eur As New cls_field
    Public txt_premium_local As New cls_field
    Public txt_local_fx_to_eur As New cls_field
    Public txt_deal_fx_to_eur As New cls_field
    Public txt_currency_date As New cls_field
    Public txt_vat_home As New cls_field
    Public txt_creating_uw As New cls_field
    
    'misc section
    Public txt_is_underwritten_id As New cls_field
    Public txt_surplus_broker_firm As New cls_field
    Public txt_surplus_broker_person As New cls_field
    Public txt_surplus_broker_license_no As New cls_field
    Public txt_us_producer_no As New cls_field
    Public txt_comments As New cls_field
    Public txt_stella_message_board As New cls_field

'command buttons
    'expanders
    Public cmd_expand_admin_section As New cls_field
    Public cmd_expand_closing_info_section As New cls_field
    Public cmd_expand_submission_notes_section As New cls_field
    
    'general
    Public cmd_general_refresh_nbi_templates As New cls_field
    Public cmd_general_open_folder As New cls_field
    Public cmd_general_open_imanage As New cls_field
    Public cmd_general_create_folder As New cls_field
    Public cmd_general_new_claim As New cls_field
    Public cmd_general_open_claims As New cls_field
    Public cmd_general_show_policies As New cls_field
    Public cmd_general_open_cm As New cls_field
    Public cmd_general_search As New cls_field
    Public cmd_generate_rr As New cls_field
    Public cmd_general_tags As New cls_field
    Public cmd_general_uw_positions As New cls_field
    Public cmd_general_delete As New cls_field
    Public cmd_general_emails As New cls_field 'not actually a drop down, but is with the CMDs and works as one

'misc txt fields
    Public notifiation_area As New cls_field
    Public help_text As New cls_field
    
'placements
    Public columns_first_left  As Long
    Public columns_second_left As Long
    Public columns_third_left As Long
    Public rows_first_top As Long
    Public rows_second_top As Long
    Public rows_third_top As Long

'collections for each section
Public col_all_labels As New Collection
Public col_all_text_fields As New Collection

Public col_labels_for_nda_section As New Collection
Public col_txt_fields_for_nda_section As New Collection

Public col_labels_for_deal_and_target_info As New Collection
Public col_txt_fields_for_deal_and_target_info As New Collection

Public col_labels_for_advisors As New Collection
Public col_txt_fields_for_advisors As New Collection

Public col_labels_for_insurance_info As New Collection
Public col_txt_fields_for_insurance_info As New Collection

Public col_labels_for_buyer_and_seller_section As New Collection
Public col_txt_fields_for_buyer_and_seller_section As New Collection

Public col_labels_for_submission_notes_section As New Collection
Public col_txt_fields_for_submission_notes_section As New Collection

Public col_labels_for_financial_info_section As New Collection
Public col_txt_fields_for_financial_info_section As New Collection

Public col_labels_for_closing_info As New Collection
Public col_txt_fields_for_closing_info As New Collection

Public col_labels_for_admin_section As New Collection
Public col_txt_fields_for_admin_section As New Collection

Public col_labels_for_misc_section As New Collection
Public col_txt_fields_for_misc_section As New Collection

Public col_cmd_expanders As New Collection
Public col_cmd_generals As New Collection

Public col_misc_txt_fields As New Collection
Public col_special_america_fields As New Collection

'misc
Public is_init As Boolean
Public shall_be_refreshed As Boolean
Public make_all_fields_visible_for_debugging As Boolean
Public form_name As String
Public deal_tags As New cls_deal_tags

Public Sub init_collections()
    Const proc_name As String = "cls_form_deal_deatils.init_collections"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0

    Dim form_field As cls_field

    'misc fields
    Set Me.col_misc_txt_fields = Nothing
    With Me.col_misc_txt_fields
        .Add Me.notifiation_area
        .Add Me.help_text
    End With
    
'headers
    Set col_headers = Nothing
    With col_headers
        .Add Me.header_nda_stage
        .Add Me.header_deal_and_target_info
        .Add Me.header_advisors
        .Add Me.header_insurance_info
        .Add Me.header_buyer_and_seller
        .Add Me.header_submission_notes
        .Add Me.header_financial_info
        .Add Me.header_closing_info
        .Add Me.header_admin
        .Add Me.header_misc
    End With

'command buttons
    Set Me.col_cmd_expanders = Nothing
    With Me.col_cmd_expanders
        .Add Me.cmd_expand_admin_section
        .Add Me.cmd_expand_closing_info_section
        .Add Me.cmd_expand_submission_notes_section
    End With
    
    Set Me.col_cmd_generals = Nothing
    With Me.col_cmd_generals
        .Add Me.cmd_general_search
        .Add Me.cmd_general_refresh_nbi_templates
        .Add Me.cmd_general_open_folder
        .Add Me.cmd_general_open_imanage
        .Add Me.cmd_general_create_folder
        '.Add Me.cmd_general_new_claim
        '.Add Me.cmd_general_open_claims
        .Add Me.cmd_general_show_policies
        .Add Me.cmd_general_open_cm
        .Add Me.cmd_generate_rr
        .Add Me.cmd_general_tags
        .Add Me.cmd_general_uw_positions
        .Add Me.cmd_general_delete
        .Add Me.cmd_general_emails
    End With

'labels
    Set Me.col_labels_for_nda_section = Nothing
    With Me.col_labels_for_nda_section
        .Add Me.lbl_deal_id
        .Add Me.lbl_deal_status_id
        .Add Me.lbl_risk_name
        .Add Me.lbl_broker_firm_id
        .Add Me.lbl_broker_person_id
        .Add Me.lbl_spa_law_id
        .Add Me.lbl_stage_id
        .Add Me.lbl_max_limit_quoted
        .Add Me.lbl_quote_due_date
        .Add Me.lbl_was_quoted_id
        .Add Me.lbl_is_test_deal_id
    End With
    
    Set Me.col_labels_for_deal_and_target_info = Nothing
    With Me.col_labels_for_deal_and_target_info
        .Add Me.lbl_risk_currency_id
        .Add Me.lbl_ev
        .Add Me.lbl_target_legal_name
        .Add Me.lbl_target_business_name
        .Add Me.lbl_target_registered_country_id
        .Add Me.lbl_target_desc
        .Add Me.lbl_target_main_country_id
        .Add Me.lbl_target_super_sector_id
        .Add Me.lbl_target_sub_sector_id
        .Add Me.lbl_target_sic_code
        .Add Me.lbl_target_naic_code
        .Add Me.lbl_risk_signing_date
        .Add Me.lbl_risk_closing_date
    End With
    
    Set Me.col_labels_for_advisors = Nothing
    With Me.col_labels_for_advisors
        .Add lbl_seller_law_firm_id
        .Add Me.lbl_buyer_law_firm_1_id
        .Add Me.lbl_buyer_law_firm_2_id
        .Add Me.lbl_buyer_financial_advisor_id
        .Add Me.lbl_legal_uw_counsel_firm_id
        .Add Me.lbl_uw_counsel_1_id
        .Add Me.lbl_financial_uw_counsel_firm_id
        .Add Me.lbl_uw_counsel_2_id
    End With
    
    Set Me.col_labels_for_insurance_info = Nothing
    With Me.col_labels_for_insurance_info
        .Add Me.lbl_risk_type_id
        .Add Me.lbl_nbi_prepper
        .Add Me.lbl_primary_uw_id
        .Add Me.lbl_secondary_uw_id
        .Add Me.lbl_analyst_id
        .Add Me.lbl_internal_approver_quote_id
        .Add Me.lbl_internal_approver_binding_id
        .Add Me.lbl_risk_feel_id
        .Add Me.lbl_primary_or_xs_id
        .Add Me.lbl_budget_home_id
        .Add Me.lbl_program_limit
        .Add Me.lbl_primary_insurer_id
        .Add Me.lbl_retention
        .Add Me.lbl_retention_dropped
        .Add Me.lbl_drop_period
        .Add Me.lbl_inception_date
        .Add Me.lbl_policy_period
    End With
    
    Set Me.col_labels_for_buyer_and_seller_section = Nothing
    With Me.col_labels_for_buyer_and_seller_section
        .Add lbl_ultimate_seller
        .Add lbl_seller_business_name
        .Add Me.lbl_ultimate_buyer
        .Add Me.lbl_buyer_business_name
        .Add Me.lbl_insured_legal_name
        .Add Me.lbl_insured_registered_country_id
        .Add Me.lbl_insured_main_region_id
    End With
    
    Set Me.col_labels_for_submission_notes_section = Nothing
    With Me.col_labels_for_submission_notes_section
        .Add Me.lbl_reviewer_notes
        .Add Me.lbl_deal_info
        .Add Me.lbl_submission_limits
    End With
    
    Set Me.col_labels_for_financial_info_section = Nothing
    With Me.col_labels_for_financial_info_section
        .Add Me.lbl_amounts
        .Add Me.lbl_due_dates
        .Add Me.lbl_received_dates
        .Add Me.lbl_uw_fee
        .Add Me.lbl_counsel_fee
        .Add Me.lbl_net_uw_fee
        .Add Me.lbl_signing_premium
        .Add Me.lbl_closing_premium
        .Add Me.lbl_total_premium
    End With
    
    Set Me.col_labels_for_closing_info = Nothing
    With Me.col_labels_for_closing_info
        .Add Me.lbl_sanction_checks
        .Add Me.lbl_esg_country
        .Add Me.lbl_esg_company
        .Add Me.lbl_is_counsel_invoice_received_id
        .Add Me.lbl_is_vdr_received_id
        .Add Me.lbl_vdr_password
        .Add Me.lbl_closing_set_received
        .Add Me.lbl_is_rr_completed_id
        .Add Me.lbl_are_emails_filed_id
        .Add Me.lbl_is_closing_booked_id
    End With
    
    Set Me.col_labels_for_admin_section = Nothing
    With Me.col_labels_for_admin_section
        .Add Me.lbl_create_date
        .Add Me.lbl_premium_eur
        .Add Me.lbl_premium_local
        .Add Me.lbl_local_fx_to_eur
        .Add Me.lbl_deal_fx_to_eur
        .Add Me.lbl_currency_date
        .Add Me.lbl_vat_home
        .Add Me.lbl_creating_uw
    End With
    
    Set Me.col_labels_for_misc_section = Nothing
    With Me.col_labels_for_misc_section
        .Add Me.lbl_surplus_broker_firm
        .Add Me.lbl_surplus_broker_person
        .Add Me.lbl_surplus_broker_license_no
        .Add Me.lbl_us_producer_no
        .Add Me.lbl_is_underwritten_id
        .Add Me.lbl_comments
        .Add Me.lbl_stella_message_board
    End With
    
    Set Me.col_all_labels = Nothing
    With Me.col_all_labels
        For Each form_field In Me.col_labels_for_nda_section
            .Add form_field
        Next form_field
        For Each form_field In Me.col_labels_for_deal_and_target_info
            .Add form_field
        Next form_field
        For Each form_field In Me.col_labels_for_advisors
            .Add form_field
        Next form_field
        For Each form_field In Me.col_labels_for_insurance_info
            .Add form_field
        Next form_field
        For Each form_field In Me.col_labels_for_buyer_and_seller_section
            .Add form_field
        Next form_field
        For Each form_field In Me.col_labels_for_submission_notes_section
            .Add form_field
        Next form_field
        For Each form_field In Me.col_labels_for_financial_info_section
            .Add form_field
        Next form_field
        For Each form_field In Me.col_labels_for_closing_info
            .Add form_field
        Next form_field
        For Each form_field In Me.col_labels_for_admin_section
            .Add form_field
        Next form_field
        For Each form_field In Me.col_labels_for_misc_section
            .Add form_field
        Next form_field
    End With

'text fields
    Set Me.col_txt_fields_for_nda_section = Nothing
    With Me.col_txt_fields_for_nda_section
        .Add Me.txt_deal_status_id
        .Add Me.txt_risk_name
        .Add Me.txt_deal_id
        .Add Me.txt_broker_firm_id
        .Add Me.txt_broker_person
        .Add Me.txt_spa_law
        .Add Me.txt_stage_id
        .Add Me.txt_max_limit_quoted
        .Add Me.txt_quote_due_date
        .Add Me.txt_was_quoted_id
        .Add Me.txt_is_test_deal_id
    End With
    
    Set Me.col_txt_fields_for_deal_and_target_info = Nothing
    With Me.col_txt_fields_for_deal_and_target_info
        .Add Me.txt_risk_currency_id
        .Add Me.txt_ev
        .Add Me.txt_target_legal_name
        .Add Me.txt_target_business_name
        .Add Me.txt_target_registered_country_id
        .Add Me.txt_target_desc
        .Add Me.txt_target_main_country_id
        .Add Me.txt_target_super_sector_id
        .Add Me.txt_target_sub_sector_id
        .Add Me.txt_target_sic_code
        .Add Me.txt_target_naic_code
        .Add Me.txt_risk_signing_date
        .Add Me.txt_risk_closing_date
    End With
    
    Set Me.col_txt_fields_for_advisors = Nothing
    With Me.col_txt_fields_for_advisors
        .Add txt_seller_law_firm_id
        .Add txt_buyer_law_firm_1_id
        .Add txt_buyer_law_firm_2_id
        .Add txt_buyer_financial_advisor_id
        .Add Me.txt_uw_law_firm_id
        .Add Me.txt_uw_counsel_1_id
        .Add Me.txt_uw_financial_advisor_id
        .Add Me.txt_uw_counsel_2_id
    End With
    
    Set Me.col_txt_fields_for_insurance_info = Nothing
    With Me.col_txt_fields_for_insurance_info
        .Add Me.txt_risk_type_id
        .Add Me.txt_nbi_prepper_id
        .Add Me.txt_nbi_prepper_full_name
        .Add Me.txt_primary_uw_id
        .Add Me.txt_primary_uw_full_name
        .Add Me.txt_secondary_uw_id
        .Add Me.txt_secondary_uw_full_name
        .Add Me.txt_analyst_id
        .Add Me.txt_analyst_full_name
        .Add Me.txt_internal_approver_quote_id
        .Add Me.txt_internal_approver_quote_full_name
        .Add Me.txt_internal_approver_binding_id
        .Add Me.txt_internal_approver_binding_full_name
        .Add Me.txt_risk_feel_id
        .Add Me.txt_primary_or_xs_id
        .Add Me.txt_budget_home_id
        .Add Me.txt_program_limit
        .Add Me.txt_primary_insurer_id
        .Add Me.txt_retention
        .Add Me.txt_retention_dropped
        .Add Me.txt_drop_period
        .Add Me.txt_inception_date
        .Add Me.txt_policy_period
    End With
    
    Set Me.col_txt_fields_for_buyer_and_seller_section = Nothing
    With Me.col_txt_fields_for_buyer_and_seller_section
        .Add Me.txt_ultimate_seller
        .Add Me.txt_seller_business_name
        .Add Me.txt_ultimate_buyer
        .Add Me.txt_buyer_business_name
        .Add Me.txt_insured_legal_name
        .Add Me.txt_insured_registered_country_id
        .Add Me.txt_insured_main_region_id
    End With
    
    Set Me.col_txt_fields_for_submission_notes_section = Nothing
    With Me.col_txt_fields_for_submission_notes_section
        .Add Me.txt_reviewer_notes
        .Add Me.txt_deal_info
        .Add Me.txt_submission_limits
    End With
    
    Set Me.col_txt_fields_for_financial_info_section = Nothing
    With Me.col_txt_fields_for_financial_info_section
        .Add Me.txt_uw_fee
        .Add Me.txt_counsel_fee
        .Add Me.txt_net_uw_fee
        .Add Me.txt_uw_fee_due_date
        .Add Me.txt_uw_fee_received_date
        .Add Me.txt_counsel_fee_due_date
        .Add Me.txt_counsel_fee_received_date
        .Add Me.txt_signing_premium_amount
        .Add Me.txt_signing_premium_due_date
        .Add Me.txt_signing_premium_received_date
        .Add Me.txt_closing_premium_amount
        .Add Me.txt_closing_premium_due_date
        .Add Me.txt_closing_premium_received_date
        .Add Me.txt_total_premium
    End With
    
    Set Me.col_txt_fields_for_closing_info = Nothing
    With Me.col_txt_fields_for_closing_info
        .Add Me.txt_sanction_checks
        .Add Me.txt_esg_country
        .Add Me.txt_esg_company
        .Add Me.txt_is_counsel_invoice_received_id
        .Add Me.txt_is_vdr_received_id
        .Add Me.txt_vdr_password
        .Add Me.txt_closing_set_received
        .Add Me.txt_is_rr_completed_id
        .Add Me.txt_are_emails_filed_id
        .Add Me.txt_is_closing_booked_id
    End With
    
    Set Me.col_txt_fields_for_admin_section = Nothing
    With Me.col_txt_fields_for_admin_section
        .Add Me.txt_create_date
        .Add Me.txt_premium_eur
        .Add Me.txt_premium_local
        .Add Me.txt_deal_fx_to_eur
        .Add Me.txt_local_fx_to_eur
        .Add Me.txt_currency_date
        .Add Me.txt_vat_home
        .Add Me.txt_creating_uw
    End With
    
    Set Me.col_txt_fields_for_misc_section = Nothing
    With Me.col_txt_fields_for_misc_section
        .Add Me.txt_surplus_broker_firm
        .Add Me.txt_surplus_broker_person
        .Add Me.txt_surplus_broker_license_no
        .Add Me.txt_us_producer_no
        .Add Me.txt_is_underwritten_id
        .Add Me.txt_comments
        .Add Me.txt_stella_message_board
    End With
    
    'add all text fields to collection for all text fields
    Set Me.col_all_text_fields = Nothing
    With Me.col_all_text_fields
        For Each form_field In Me.col_txt_fields_for_nda_section
            .Add form_field
        Next form_field
        For Each form_field In Me.col_txt_fields_for_deal_and_target_info
            .Add form_field
        Next form_field
        For Each form_field In Me.col_txt_fields_for_advisors
            .Add form_field
        Next form_field
        For Each form_field In Me.col_txt_fields_for_insurance_info
            .Add form_field
        Next form_field
        For Each form_field In Me.col_txt_fields_for_buyer_and_seller_section
            .Add form_field
        Next form_field
        For Each form_field In Me.col_txt_fields_for_submission_notes_section
            .Add form_field
        Next form_field
        For Each form_field In Me.col_txt_fields_for_financial_info_section
            .Add form_field
        Next form_field
        For Each form_field In Me.col_txt_fields_for_closing_info
            .Add form_field
        Next form_field
        For Each form_field In Me.col_txt_fields_for_admin_section
            .Add form_field
        Next form_field
        For Each form_field In Me.col_txt_fields_for_misc_section
            .Add form_field
        Next form_field
    End With

    'continent specific collections
    Set Me.col_special_america_fields = Nothing
    With Me.col_special_america_fields
        .Add Me.lbl_stage_id
        .Add Me.txt_stage_id
        .Add Me.lbl_max_limit_quoted
        .Add Me.txt_max_limit_quoted
        .Add Me.lbl_internal_approver_quote_id
        .Add Me.txt_internal_approver_quote_id
        .Add Me.txt_internal_approver_quote_full_name
        .Add Me.lbl_internal_approver_binding_id
        .Add Me.txt_internal_approver_binding_id
        .Add Me.txt_internal_approver_binding_full_name
        .Add Me.lbl_target_sic_code
        .Add Me.txt_target_sic_code
        .Add Me.lbl_target_naic_code
        .Add Me.txt_target_naic_code
        .Add Me.lbl_retention_dropped
        .Add Me.txt_retention_dropped
        .Add Me.lbl_drop_period
        .Add Me.txt_drop_period
        .Add Me.lbl_esg_country
        .Add Me.txt_esg_country
        .Add Me.lbl_esg_company
        .Add Me.txt_esg_company
        .Add Me.lbl_closing_set_received
        .Add Me.txt_closing_set_received
    End With

outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init_notification_area(ByVal deal_id As Long)
    With Me.notifiation_area
        If Me.are_there_claims_on_deal(deal_id) = True Then
            .field_top = Me.rows_first_top
            .field_height = utilities.twips_converter(2, "cm")
            .field_width = utilities.twips_converter(8, "cm")
            .field_value = "There are claim(s) on this deal. Check with Stella Claims or the claims team for more info."
            .field_visible = True
            .field_bg_color = load.colors.light_red
        Else
            .field_visible = False
            .field_value = ""
            .field_height = 0
        End If
    End With
End Sub
Public Sub refresh_first_column()
    Me.refresh_admin_section
End Sub
Public Sub refresh_second_column()
    Me.refresh_deal_and_target_info_section
    Me.refresh_buyer_and_seller
End Sub
Public Sub refresh_third_column()
    Me.refresh_advisors_section
    Me.refresh_submission_notes_section
    Me.refresh_financial_info_section
    Me.refresh_misc_section
End Sub
Public Sub refresh_deal_details(ByVal deal_id As Long)
    Const proc_name As String = "cls_form_deal_details.refresh_deal_details"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim above_field As cls_field
    Dim field_counter As Integer
    Dim field_left As Long
    
    Me.init_notification_area deal_id
    
    '''''''''
    'headers'
    '''''''''
        Me.refresh_nda_section
        
        With Me.header_deal_and_target_info
            'no settings to refresh
        End With
        
        Me.refresh_advisors_section
               
        Me.refresh_insurance_info_section
        
        Me.refresh_buyer_and_seller
        
        Me.refresh_submission_notes_section
        
        Me.refresh_financial_info_section
        
        Me.refresh_admin_section
        
        Me.refresh_check_points_section
        
        Me.refresh_misc_section
    
        
    ''''''''
    'labels'
    ''''''''
        'set individual properties for labels
        
        'deal and target info section
        With Me.lbl_risk_currency_id
            .field_top = Me.header_deal_and_target_info.field_top + Me.header_deal_and_target_info.field_height
            Set above_field = Me.lbl_risk_currency_id
        End With
        With Me.lbl_ev
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.lbl_ev
        End With
        With Me.lbl_target_legal_name
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.lbl_target_legal_name
        End With
        With Me.lbl_target_business_name
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.lbl_target_business_name
        End With
        With Me.lbl_target_registered_country_id
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.lbl_target_registered_country_id
        End With
        With Me.lbl_target_main_country_id
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.lbl_target_main_country_id
        End With
        With Me.lbl_target_super_sector_id
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.lbl_target_super_sector_id
        End With
        With Me.lbl_target_sub_sector_id
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.lbl_target_sub_sector_id
        End With
        With Me.lbl_target_sic_code
            .field_top = above_field.field_top + above_field.field_height
            If load.system_info.app_continent = load.system_info.continents.americas Then
                Set above_field = Me.lbl_target_sic_code
            End If
        End With
        With Me.lbl_target_naic_code
            .field_top = above_field.field_top + above_field.field_height
            If load.system_info.app_continent = load.system_info.continents.americas Then
                Set above_field = Me.lbl_target_naic_code
            End If
        End With
        With Me.lbl_risk_signing_date
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.lbl_risk_signing_date
        End With
        With Me.lbl_risk_closing_date
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.lbl_risk_closing_date
        End With
        
    '''''''''''''
    'text fields'
    '''''''''''''
       
        'deal and target info section
        field_left = 6700
        With Me.txt_risk_currency_id
            .field_top = Me.header_deal_and_target_info.field_top + Me.header_deal_and_target_info.field_height
            Set above_field = Me.txt_risk_currency_id
        End With
        With Me.txt_ev
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.txt_ev
        End With
        With Me.txt_target_legal_name
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.txt_target_legal_name
        End With
        With Me.txt_target_business_name
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.txt_target_business_name
        End With
        With Me.txt_target_registered_country_id
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.txt_target_registered_country_id
        End With
        With Me.txt_target_main_country_id
            'need to cater for extra height on lavel on previous row
            .field_top = above_field.field_top + Me.lbl_target_registered_country_id.field_height
            Set above_field = Me.txt_target_main_country_id
        End With
        With Me.txt_target_super_sector_id
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.txt_target_super_sector_id
        End With
        With Me.txt_target_sub_sector_id
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.txt_target_sub_sector_id
        End With
        With Me.txt_target_sic_code
            .field_top = above_field.field_top + above_field.field_height
            If load.system_info.app_continent = load.system_info.continents.americas Then
                Set above_field = Me.txt_target_sic_code
            End If
        End With
        With Me.txt_target_naic_code
            .field_top = above_field.field_top + above_field.field_height
            If load.system_info.app_continent = load.system_info.continents.americas Then
                Set above_field = Me.txt_target_naic_code
            End If
        End With
        With Me.txt_risk_signing_date
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.txt_risk_signing_date
        End With
        With Me.txt_risk_closing_date
            .field_top = above_field.field_top + above_field.field_height
            Set above_field = Me.txt_risk_closing_date
        End With
    
outro:
    Set above_field = Nothing
    Set form_field = Nothing
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub

Public Sub init(Optional ByVal deal_id As Long)
    Const proc_name As String = "cls_form_deal_details.init"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim above_field As cls_field
    Dim field_left As Long
    Dim section_label_width As Long
    Dim section_field_width As Long
    
    form_name = "deal_details_f"
    
    Me.init_collections
    
    'set default properties for all labels
    For Each form_field In Me.col_all_labels
        With form_field
            .field_height = utilities.twips_converter(0.56, "cm")
            .font_size = 10
            .field_visible = True
            .field_width = 1800
        End With
    Next form_field
    
    'set common properties for all text fields
    For Each form_field In Me.col_all_text_fields
        With form_field
            .field_height = utilities.twips_converter(0.56, "cm")
            .font_size = 10
            .field_visible = True
            .field_width = 1800
            .has_help_text = False
            .field_bg_color = load.colors.white
        End With
    Next form_field
    
    Me.columns_first_left = 50
    Me.columns_second_left = 4350
    Me.columns_third_left = 10600
    Me.rows_first_top = 1000
    
    
    '''''''''''''
    'misc fields'
    '''''''''''''
        
        For Each form_field In Me.col_misc_txt_fields
            With form_field
                .field_height = 0
                .field_value = ""
                .field_visible = False
                .field_bg_color = load.colors.white
                .font_color = load.colors.black
            End With
        Next form_field
        
        With Me.notifiation_area
            .field_top = Me.rows_first_top
            .field_left = Me.columns_third_left
            .field_name = "notification_area"
        End With
        If deal_id <> 0 Then
            Me.init_notification_area deal_id
        End If
        
        With Me.help_text
            .field_name = "help_text"
        End With
        
        
    ''''''
    'cmds'
    ''''''
        
        'set default values for each collection
        For Each form_field In Me.col_cmd_expanders
            With form_field
                .font_color = load.colors.white
                .font_color_pressed = load.colors.white
                .font_color_hover = load.colors.white
                .field_caption = "collapse"
            End With
        Next form_field

        For Each form_field In Me.col_cmd_generals
            With form_field
                .font_color = load.colors.white
                .font_color_pressed = load.colors.white
                .font_color_hover = load.colors.white
                .field_bg_color = load.colors.deal_details_command_button
                .field_top = 100
                .field_width = utilities.twips_converter(2.4, "cm")
                .field_visible = True
            End With
        Next form_field
        
        'add specific values to each command button
        
        'command buttons at top of form
        Dim space_between_cmds As Long
        space_between_cmds = 250
        
        With Me.cmd_general_search
            .field_name = "cmdSearchNewDeal"
            .field_caption = "search for other deal"
            .field_left = 200
            .field_width = utilities.twips_converter(2, "cm")
            Set above_field = Me.cmd_general_search
        End With
        With Me.cmd_general_refresh_nbi_templates
            .field_name = "cmd_refresh_nbi_templates"
            .field_caption = "refresh nbi templates"
            .field_left = above_field.field_left + above_field.field_width + space_between_cmds
            Set above_field = Me.cmd_general_refresh_nbi_templates
        End With
        With Me.cmd_general_open_folder
            .field_name = "cmdOpenFolders"
            .field_caption = "open folder"
            .field_left = above_field.field_left + above_field.field_width + space_between_cmds
            .field_height = utilities.twips_converter(0.6, "cm")
            .field_width = utilities.twips_converter(3, "cm")
            Set above_field = Me.cmd_general_open_folder
        End With
        With Me.cmd_general_open_imanage
            .field_name = "cmd_open_workspace"
            .field_caption = "open iManage"
            .field_left = Me.cmd_general_open_folder.field_left
            .field_top = Me.cmd_general_open_folder.field_bottom + 100
            .field_width = utilities.twips_converter(3, "cm")
        End With
        With Me.cmd_general_emails
            .field_left = Me.cmd_general_open_imanage.field_right + space_between_cmds
            .field_name = "emails"
            .field_type = "combo_box"
            Set above_field = Me.cmd_general_emails
        End With
        With Me.cmd_general_create_folder
            .field_caption = "create folder"
            .field_name = "cmdCreateFolder"
            .field_left = above_field.field_right + space_between_cmds
            Set above_field = Me.cmd_general_create_folder
        End With
        With Me.cmd_general_show_policies
            .field_name = "cmd_policies"
            .field_caption = "policies"
            .field_left = above_field.field_left + above_field.field_width + space_between_cmds
            Set above_field = Me.cmd_general_show_policies
        End With
        With Me.cmd_general_open_cm
            .field_name = "cmd_open_cm"
            .field_caption = "compliance module"
            .field_left = above_field.field_left + above_field.field_width + space_between_cmds
            Set above_field = Me.cmd_general_open_cm
        End With
        With Me.cmd_generate_rr
            .field_name = "cmd_generate_rr"
            .field_caption = "generate rr"
            .field_left = above_field.field_left + above_field.field_width + space_between_cmds
            Set above_field = Me.cmd_generate_rr
        End With
        With Me.cmd_general_tags
            .field_name = "cmd_general_tags"
            .field_caption = "deal tags"
            .field_left = above_field.field_left + above_field.field_width + space_between_cmds
            .field_width = utilities.twips_converter(1.5, "cm")
            Set above_field = Me.cmd_general_tags
        End With
        With Me.cmd_general_uw_positions
            .field_name = "cmd_general_uw_positions"
            .field_caption = "UW positions"
            .field_left = Me.cmd_general_tags.field_right + space_between_cmds
            Set above_field = Me.cmd_general_uw_positions
        End With
        With Me.cmd_general_delete
            .field_name = "cmdDelete"
            .field_caption = "delete"
            .field_left = above_field.field_left + above_field.field_width + space_between_cmds
            .field_bg_color = load.colors.delete_button
            .field_width = utilities.twips_converter(1.6, "cm")
            Set above_field = Me.cmd_general_delete
        End With
        
    
    '''''''''
    'headers'
    '''''''''
        For Each form_field In Me.col_headers
            With form_field
                .field_bg_color = load.colors.label_bg
                .border_color = load.colors.white
                .margin_internal_left = 50
                .top_margin = 0
                .font_is_bold = True
                .font_color = load.colors.white
                .font_size = 14
                .field_height = utilities.twips_converter(0.7, "cm")
                .field_width = utilities.twips_converter(6.5, "cm")
            End With
        Next form_field
        
        Me.init_nda_section
        
        Me.init_deal_and_target_info_section
        
        Me.init_advisors_section
        
        Me.init_submisson_notes_section
        
        Me.init_financial_info_section
        
        Me.init_misc_section
        
        Me.init_insurance_info_section
        
        Me.init_buyer_and_seller_section
        
        Me.init_admin_section
        
        Me.init_check_points_section
    
    'hide America specific fields and labels unless in America
    If load.system_info.app_continent <> load.system_info.continents.americas Then
        For Each form_field In Me.col_special_america_fields
            With form_field
                If Me.make_all_fields_visible_for_debugging = False Then
                    .field_visible = False
                End If
            End With
        Next form_field
    End If
    
    is_init = True

outro:
    Set above_field = Nothing
    Set form_field = Nothing
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init_nda_section()
    Const proc_name As String = "cls_form_deal_details.init_nda_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim above_field As cls_field
    Dim field_left As Long
    
    'header
    With Me.header_nda_stage
        .field_name = "header_nda_stage"
        .field_top = Me.rows_first_top
        .field_left = Me.columns_first_left
        If .user_has_expanded_section = False Then
            .is_collapsed = False
        End If
        If load.system_info.app_continent = load.system_info.continents.americas Then
            .section_height = utilities.twips_converter(8, "cm")
        Else
            .section_height = utilities.twips_converter(8, "cm")
        End If
    End With
    
    'labels
    For Each form_field In Me.col_labels_for_nda_section
        With form_field
            .field_left = Me.header_nda_stage.field_left + 50
        End With
    Next form_field
    
    With Me.lbl_deal_id
        .field_name = "lbl_deal_id"
        .field_top = Me.header_nda_stage.field_top + Me.header_nda_stage.field_height
        Set above_field = Me.lbl_deal_id
    End With
    With Me.lbl_deal_status_id
        .field_name = "lbl_deal_status_id"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_deal_status_id
    End With
    With Me.lbl_risk_name
        .field_name = "lbl_risk_name"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_risk_name
    End With
    With Me.lbl_broker_firm_id
        .field_name = "lbl_broker_firm_id"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_broker_firm_id
    End With
    With Me.lbl_broker_person_id
        .field_name = "lbl_broker_person_id"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_broker_person_id
    End With
    With Me.lbl_spa_law_id
        .field_name = "lbl_spa_law_id"
        .field_top = above_field.field_top + above_field.field_height
        .field_height = utilities.twips_converter(1.2, "cm")
        Set above_field = Me.lbl_spa_law_id
    End With
    With Me.lbl_stage_id
        .field_name = "lbl_stage_id"
        .field_top = above_field.field_top + above_field.field_height
        If load.system_info.app_continent = load.system_info.continents.americas Or Me.make_all_fields_visible_for_debugging = True Then
            Set above_field = Me.lbl_stage_id
        End If
    End With
    With Me.lbl_max_limit_quoted
        .field_name = "lbl_max_limit_quoted"
        .field_top = above_field.field_top + above_field.field_height
        If load.system_info.app_continent = load.system_info.continents.americas Or Me.make_all_fields_visible_for_debugging = True Then
            Set above_field = Me.lbl_max_limit_quoted
        End If
    End With
    With Me.lbl_quote_due_date
        .field_name = "lbl_quote_due_date"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_quote_due_date
    End With
    With Me.lbl_was_quoted_id
        .field_name = "lbl_was_quoted_id"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_was_quoted_id
    End With
    With Me.lbl_is_test_deal_id
        .field_name = "lbl_is_test_deal_id"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_is_test_deal_id
    End With
    
    'text fields
    With Me.txt_deal_id
        .field_name = "deal_id"
        .field_name_in_recordset = "deal_id"
        .field_top = Me.header_nda_stage.field_top + Me.header_nda_stage.field_height
        Set above_field = Me.txt_deal_id
    End With
    With Me.txt_deal_status_id
        .field_name = "deal_status_id"
        .field_name_in_recordset = "deal_status_id"
        .field_top = above_field.field_top + above_field.field_height
        .has_help_text = True
        Set above_field = Me.txt_deal_status_id
    End With
    With Me.txt_risk_name
        .field_name = "deal_name"
        .field_name_in_recordset = "deal_name"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.txt_risk_name
    End With
    With Me.txt_broker_firm_id
        .field_name = "broker_firm_id"
        .field_name_in_recordset = "broker_firm_id"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.txt_broker_firm_id
    End With
    With Me.txt_broker_person
        .field_name = "broker_person"
        .field_name_in_recordset = "broker_person"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.txt_broker_person
    End With
    With Me.txt_spa_law
        .field_name = "spa_law"
        .field_name_in_recordset = "spa_law"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.txt_spa_law
    End With
    With Me.txt_stage_id
        .field_name = "stage_id"
        .field_name_in_recordset = "deal_stage_id"
        .field_top = above_field.field_top + utilities.twips_converter(1.2, "cm")
        If load.system_info.app_continent = load.system_info.continents.americas Or Me.make_all_fields_visible_for_debugging = True Then
            Set above_field = Me.txt_stage_id
        End If
    End With
    With Me.txt_max_limit_quoted
        .field_name = "max_limit_quoted"
        .field_name_in_recordset = "max_limit_quoted"
        .field_top = above_field.field_bottom
        If load.system_info.app_continent = load.system_info.continents.americas Or Me.make_all_fields_visible_for_debugging = True Then
            Set above_field = Me.txt_max_limit_quoted
        End If
    End With
    With Me.txt_quote_due_date
        .field_name = "quote_due_date"
        .field_name_in_recordset = "quote_due_date"
        If load.system_info.app_continent = load.system_info.continents.americas Or Me.make_all_fields_visible_for_debugging = True Then
            .field_top = above_field.field_top + above_field.field_height
        Else
            'spa label is 1.2 cm high
            .field_top = above_field.field_top + utilities.twips_converter(1.2, "cm")
        End If
        Set above_field = Me.txt_quote_due_date
    End With
    With Me.txt_was_quoted_id
        .field_name = "was_quoted_id"
        .field_name_in_recordset = "was_quoted_id"
        .field_top = above_field.field_top + above_field.field_height
        .has_help_text = True
        Set above_field = Me.txt_was_quoted_id
    End With
    With Me.txt_is_test_deal_id
        .field_name = "is_test_deal_id"
        .field_name_in_recordset = "is_test_deal_id"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.txt_is_test_deal_id
    End With
    
    For Each form_field In Me.col_txt_fields_for_nda_section
    With form_field
        .field_left = Me.lbl_deal_status_id.field_left + 50 + Me.lbl_deal_status_id.field_width
        .field_width = 1700
        .has_line = False
    End With
    Next form_field
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub refresh_nda_section()
    Dim form_field As cls_field, field_counter As Integer
    
End Sub
Public Sub init_admin_section()
    Const proc_name As String = "cls_form_deal_deatals.init_admin_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim field_left As Long
    
    'header
    With Me.header_admin
        .field_name = "header_admin"
        .field_top = Me.header_insurance_info.field_top + Me.header_insurance_info.section_height
        .field_left = Me.columns_first_left
        If Me.header_admin.user_has_expanded_section = False Then
            .is_collapsed = True
        End If
    End With
        
    'expander
    With Me.cmd_expand_admin_section
        .field_name = "cmd_expand_admin_section"
        .field_caption = "expand"
        .field_top = Me.header_admin.field_top
        .field_left = Me.header_admin.field_left + 2500
    End With
    
    'labels
    With Me.lbl_create_date
        .field_name = "lbl_create_date"
        Set .above_field = Me.header_admin
    End With
    With Me.lbl_premium_eur
        .field_name = "lbl_premium_eur"
        Set .above_field = Me.lbl_create_date
    End With
    With Me.lbl_premium_local
        .field_name = "lbl_premium_local"
        Set .above_field = Me.lbl_premium_eur
    End With
    With Me.lbl_local_fx_to_eur
        .field_name = "lbl_local_fx_to_eur"
        Set .above_field = Me.lbl_premium_local
    End With
    With Me.lbl_deal_fx_to_eur
        .field_name = "lbl_deal_fx_to_eur"
        Set .above_field = Me.lbl_local_fx_to_eur
    End With
    With Me.lbl_currency_date
        .field_name = "lbl_currency_date"
        Set .above_field = Me.lbl_deal_fx_to_eur
    End With
    With Me.lbl_vat_home
        .field_name = "lbl_vat_home"
        Set .above_field = Me.lbl_currency_date
    End With
    With Me.lbl_creating_uw
        .field_name = "lbl_creating_uw"
        Set .above_field = Me.lbl_vat_home
    End With
    
    For Each form_field In Me.col_labels_for_admin_section
        With form_field
            .field_left = Me.header_admin.field_left + 50
            If Me.header_admin.user_has_expanded_section = False Then
                .field_visible = False
            End If
            .field_top = .above_field.field_top + .above_field.field_height
        End With
    Next form_field
    
    'text fields
    field_left = Me.lbl_create_date.field_left + Me.lbl_create_date.field_width
    With Me.txt_create_date
        .field_name = "create_date"
        .field_name_in_recordset = "create_date"
        Set .above_field = Me.header_admin
    End With
    With Me.txt_premium_eur
        .field_name = "premium_eur"
        .field_name_in_recordset = "total_rp_premium_on_deal_eur"
        Set .above_field = Me.txt_create_date
    End With
    With Me.txt_premium_local
        .field_name = "premium_local"
        .field_name_in_recordset = "total_rp_premium_on_deal_local"
        Set .above_field = Me.txt_premium_eur
    End With
    With Me.txt_local_fx_to_eur
        .field_name = "currency_rate_local"
        .field_name_in_recordset = "currency_rate_local"
        Set .above_field = Me.txt_premium_local
    End With
    With Me.txt_deal_fx_to_eur
        .field_name = "currency_rate_deal"
        .field_name_in_recordset = "currency_rate_deal"
        Set .above_field = Me.txt_local_fx_to_eur
    End With
    With Me.txt_currency_date
        .field_name = "currency_date"
        .field_name_in_recordset = "currency_date"
        Set .above_field = Me.txt_deal_fx_to_eur
    End With
    With Me.txt_vat_home
        .field_name = "vat_home"
        .field_name_in_recordset = "vat_home"
        Set .above_field = Me.txt_currency_date
    End With
    With Me.txt_creating_uw
        .field_name = "creating_uw"
        .field_name_in_recordset = "creating_uw"
        Set .above_field = Me.txt_vat_home
    End With
    
    For Each form_field In Me.col_txt_fields_for_admin_section
    With form_field
        .field_left = field_left
        .field_top = .above_field.field_top + .above_field.field_height
        .field_width = 1700
        If Me.header_admin.user_has_expanded_section = False Then
            .field_visible = False
        End If
        .has_line = False
    End With
    Next form_field
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub

Public Sub refresh_admin_section()
    Const proc_name As String = "cls_form_deal_details.refresh_admin_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim field_counter As Integer
    
    'header
    With Me.header_admin
        .field_top = Me.header_insurance_info.field_top + Me.header_insurance_info.section_height
    End With
        
    'expander
    With Me.cmd_expand_admin_section
        .field_top = Me.header_admin.field_top
        If Me.header_admin.user_has_expanded_section = True Then
            .field_caption = "collapse"
        Else
            .field_caption = "expand"
        End If
    End With
    
    'labels
    For Each form_field In Me.col_labels_for_admin_section
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
    
    'text fields
    For Each form_field In Me.col_txt_fields_for_admin_section
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init_advisors_section()
    Const proc_name As String = "cls_form_deal_details.init_advisors_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim field_left As Long
    
    'header
    With Me.header_advisors
        .field_name = "header_advisors"
        .field_top = Me.rows_first_top + Me.notifiation_area.field_height * 1.1
        .field_left = Me.columns_third_left
        If .user_has_expanded_section = False Then
            .is_collapsed = False
        End If
        .section_height = utilities.twips_converter(6, "cm")
    End With
        
    'labels
    With Me.lbl_seller_law_firm_id
        .field_name = "lbl_seller_law_firm_id"
        Set .above_field = Me.header_advisors
    End With
    With Me.lbl_buyer_law_firm_1_id
        .field_name = "lbl_buyer_law_firm_1_id"
        Set .above_field = Me.lbl_seller_law_firm_id
    End With
    With Me.lbl_buyer_law_firm_2_id
        .field_name = "lbl_buyer_law_firm_2_id"
        Set .above_field = Me.lbl_buyer_law_firm_1_id
    End With
    With Me.lbl_buyer_financial_advisor_id
        .field_name = "lbl_buyer_financial_advisor_id"
        Set .above_field = Me.lbl_buyer_law_firm_2_id
    End With
    With Me.lbl_legal_uw_counsel_firm_id
        .field_name = "lbl_legal_uw_counsel_firm_id"
        Set .above_field = Me.lbl_buyer_financial_advisor_id
    End With
    With Me.lbl_uw_counsel_1_id
        .field_name = "lbl_uw_counsel_1_id"
        Set .above_field = Me.lbl_legal_uw_counsel_firm_id
    End With
    With Me.lbl_financial_uw_counsel_firm_id
        .field_name = "lbl_financial_uw_counsel_firm_id"
        Set .above_field = Me.lbl_uw_counsel_1_id
    End With
    With Me.lbl_uw_counsel_2_id
        .field_name = "lbl_uw_counsel_2_id"
        Set .above_field = Me.lbl_financial_uw_counsel_firm_id
    End With
    
    For Each form_field In Me.col_labels_for_advisors
        With form_field
            .field_left = Me.header_advisors.field_left + 50
            .field_width = 2400
            .field_top = .above_field.field_top + .above_field.field_height
        End With
    Next form_field
    
    'text fields
    With Me.txt_seller_law_firm_id
        .field_name = "SellerLegalFirm"
        .field_name_in_recordset = "SellerLegalFirm"
        Set .above_field = Me.header_advisors
    End With
    With Me.txt_buyer_law_firm_1_id
        .field_name = "buyer_law_firm_1_id"
        .field_name_in_recordset = "buyer_law_firm_1_id"
        Set .above_field = Me.txt_seller_law_firm_id
    End With
    With Me.txt_buyer_law_firm_2_id
        .field_name = "buyer_law_firm_2_id"
        .field_name_in_recordset = "buyer_law_firm_2_id"
        Set .above_field = Me.txt_buyer_law_firm_1_id
    End With
    With Me.txt_buyer_financial_advisor_id
        .field_name = "BuyerFinancialCounsel"
        .field_name_in_recordset = "buyer_financial_firm_id"
        Set .above_field = Me.txt_buyer_law_firm_2_id
    End With
    With Me.txt_uw_law_firm_id
        .field_name = "uw_law_firm_id"
        .field_name_in_recordset = "uw_law_firm_id"
        Set .above_field = Me.lbl_buyer_financial_advisor_id
    End With
    With Me.txt_uw_counsel_1_id
        .field_name = "UwCounselPerson1"
        .field_name_in_recordset = "UwCounselPerson1"
        Set .above_field = Me.txt_uw_law_firm_id
    End With
    With Me.txt_uw_financial_advisor_id
        .field_name = "uw_financial_advisor_id"
        .field_name_in_recordset = "uw_financial_advisor_id"
        Set .above_field = Me.txt_uw_counsel_1_id
    End With
    With Me.txt_uw_counsel_2_id
        .field_name = "UwCounselPerson2"
        .field_name_in_recordset = "UwCounselPerson2"
        Set .above_field = Me.txt_uw_financial_advisor_id
    End With
    
    For Each form_field In Me.col_txt_fields_for_advisors
        With form_field
            .field_left = Me.lbl_seller_law_firm_id.field_left + 50 + Me.lbl_seller_law_firm_id.field_width
            .field_width = 3200
            .field_top = .above_field.field_top + .above_field.field_height
            .has_line = False
        End With
    Next form_field
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub refresh_advisors_section()
    Dim form_field As cls_field, field_counter As Integer
        
    'header
    With Me.header_advisors
        .field_top = Me.rows_first_top + Me.notifiation_area.field_height * 1.1
    End With
    
    'labels
    For Each form_field In Me.col_labels_for_advisors
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
    
    'text fields
    For Each form_field In Me.col_txt_fields_for_advisors
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
    
End Sub
Public Sub init_buyer_and_seller_section()
    Const proc_name As String = "cls_form_deal_details.init_buyer_and_seller_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim field_left As Long
    
    'header
    With Me.header_buyer_and_seller
        .field_name = "header_buyer_and_seller"
        .field_top = Me.header_deal_and_target_info.field_top + Me.header_deal_and_target_info.section_height
        .field_left = Me.columns_second_left
        .is_collapsed = False
        .section_height = utilities.twips_converter(5.5, "cm")
    End With
    
    'labels
    With Me.lbl_ultimate_seller
        .field_name = "lbl_ultimate_seller"
        Set .above_field = Me.header_buyer_and_seller
    End With
    With Me.lbl_seller_business_name
        .field_name = "lbl_seller_business_name"
        Set .above_field = Me.lbl_ultimate_seller
    End With
    With Me.lbl_ultimate_buyer
        .field_name = "lbl_ultimate_buyer"
        Set .above_field = Me.lbl_seller_business_name
    End With
    With Me.lbl_buyer_business_name
        .field_name = "lbl_buyer_business_name"
        Set .above_field = Me.lbl_ultimate_buyer
    End With
    With Me.lbl_insured_legal_name
        .field_name = "lbl_insured_legal_name"
        Set .above_field = Me.lbl_buyer_business_name
    End With
    With Me.lbl_insured_registered_country_id
        .field_name = "lbl_insured_registered_country_id"
        .field_height = utilities.twips_converter(1, "cm")
        Set .above_field = Me.lbl_insured_legal_name
    End With
    With Me.lbl_insured_main_region_id
        .field_name = "lbl_insured_main_region_id"
        Set .above_field = Me.lbl_insured_registered_country_id
    End With
    
    For Each form_field In Me.col_labels_for_buyer_and_seller_section
        With form_field
            .field_top = .above_field.field_top + .above_field.field_height
            .field_left = Me.header_buyer_and_seller.field_left + 50
            .field_width = 2500
        End With
    Next form_field
    
    'text fields
    With Me.txt_ultimate_seller
        .field_name = "UltimateSeller"
        .field_name_in_recordset = "UltimateSeller"
        Set .above_field = Me.header_buyer_and_seller
    End With
    With Me.txt_seller_business_name
        .field_name = "seller_business_name"
        .field_name_in_recordset = "seller_business_name"
        Set .above_field = Me.txt_ultimate_seller
    End With
    With Me.txt_ultimate_buyer
        .field_name = "ultimate_buyer"
        .field_name_in_recordset = "ultimate_buyer"
        Set .above_field = Me.txt_seller_business_name
    End With
    With Me.txt_buyer_business_name
        .field_name = "buyer_business_name"
        .field_name_in_recordset = "buyer_business_name"
        Set .above_field = Me.txt_ultimate_buyer
    End With
    With Me.txt_insured_legal_name
        .field_name = "insured_legal_name"
        .field_name_in_recordset = "insured_legal_name"
        Set .above_field = Me.txt_buyer_business_name
    End With
    With Me.txt_insured_registered_country_id
        .field_name = "insured_registered_country_id"
        .field_name_in_recordset = "insured_registered_country_id"
        Set .above_field = Me.txt_insured_legal_name
    End With
    With Me.txt_insured_main_region_id
        .field_name = "insured_main_region_id"
        .field_name_in_recordset = "insured_main_region_id"
        Set .above_field = Me.lbl_insured_registered_country_id
    End With
    
    For Each form_field In Me.col_txt_fields_for_buyer_and_seller_section
        With form_field
            .field_left = Me.lbl_ultimate_seller.field_right + 50
            .field_top = .above_field.field_bottom
        End With
    Next form_field
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub refresh_buyer_and_seller()
    Const proc_name As String = "cls_form_deal_details.refresh_buyer_and_seller"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim field_counter As Integer
        
    'header
    With Me.header_buyer_and_seller
        .field_top = Me.header_deal_and_target_info.field_top + Me.header_deal_and_target_info.section_height
    End With
    
    'labels
    For Each form_field In Me.col_labels_for_buyer_and_seller_section
        With form_field
            .field_top = .above_field.field_top + .above_field.field_height
        End With
    Next form_field
    
    'text fields
    For Each form_field In Me.col_txt_fields_for_buyer_and_seller_section
        With form_field
            .field_top = .above_field.field_top + .above_field.field_height
        End With
    Next form_field
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init_check_points_section()
    Const proc_name As String = "cls_form_deal_details.init_check_points_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    
    'header
    With Me.header_closing_info
        .field_name = "header_closing_info"
        .field_top = Me.header_buyer_and_seller.field_top + Me.header_buyer_and_seller.section_height
        .field_left = Me.columns_second_left
        If Me.header_closing_info.user_has_expanded_section = False Then
            .is_collapsed = True
        End If
    End With
    
    'command button for collapsing/expanding
    With Me.cmd_expand_closing_info_section
        .field_name = "cmd_expand_closing_info_section"
        If Me.header_closing_info.user_has_expanded_section = False Then
            .field_caption = "expand"
        Else
            .field_caption = "collapse"
        End If
        .field_top = Me.header_closing_info.field_top
        .field_left = Me.header_closing_info.field_left + 2500
    End With
    
    'labels
    'closing info section
    With Me.lbl_sanction_checks
        .field_name = "lbl_sanction_checks"
        Set .above_field = Me.header_closing_info
    End With
    With Me.lbl_esg_country
        .field_name = "lbl_esg_country"
        .visible_in_europe = False
        Set .above_field = Me.lbl_sanction_checks
    End With
    With Me.lbl_esg_company
        .field_name = "lbl_esg_company"
        .visible_in_europe = False
        Set .above_field = Me.lbl_esg_country
    End With
    With Me.lbl_is_counsel_invoice_received_id
        .field_name = "lbl_is_counsel_invoice_received_id"
        If load.system_info.app_continent = load.system_info.continents.americas Or load.deal_details.make_all_fields_visible_for_debugging Then
            Set .above_field = Me.lbl_esg_company
        Else
            Set .above_field = Me.lbl_sanction_checks
        End If
    End With
    With Me.lbl_is_vdr_received_id
        .field_name = "lbl_is_vdr_received_id"
        Set .above_field = Me.lbl_is_counsel_invoice_received_id
    End With
    With Me.lbl_vdr_password
        .field_name = "lbl_vdr_password"
        Set .above_field = Me.lbl_is_vdr_received_id
    End With
    With Me.lbl_closing_set_received
        .field_name = "lbl_closing_set_received"
        .visible_in_europe = False
        Set .above_field = Me.lbl_vdr_password
    End With
    With Me.lbl_is_rr_completed_id
        .field_name = "lbl_is_rr_completed_id"
        If load.system_info.app_continent = load.system_info.continents.americas Or load.deal_details.make_all_fields_visible_for_debugging Then
            Set .above_field = Me.lbl_closing_set_received
        Else
            Set .above_field = Me.lbl_vdr_password
        End If
    End With
    With Me.lbl_are_emails_filed_id
        .field_name = "lbl_are_emails_filed_id"
        Set .above_field = Me.lbl_is_rr_completed_id
    End With
    With Me.lbl_is_closing_booked_id
        .field_name = "lbl_is_closing_booked_id"
        Set .above_field = Me.lbl_are_emails_filed_id
    End With
    
    For Each form_field In Me.col_labels_for_closing_info
        With form_field
            .field_left = Me.header_closing_info.field_left + 50
            .field_width = utilities.twips_converter(4.6, "cm")
            If Me.header_closing_info.user_has_expanded_section = False Then
                .field_visible = False
            End If
            .field_top = .above_field.field_bottom
        End With
    Next form_field
        
    'text fields
    With Me.txt_sanction_checks
        .field_name = "sanction_checks"
        .field_name_in_recordset = "sanction_checks_done"
        Set .above_field = Me.header_closing_info
    End With
    With Me.txt_esg_country
        .field_name = "esg_country"
        .field_name_in_recordset = "esg_country"
                .visible_in_europe = False
        Set .above_field = Me.lbl_sanction_checks
    End With
    With Me.txt_esg_company
        .field_name = "esg_company"
        .field_name_in_recordset = "esg_company"
                .visible_in_europe = False
        Set .above_field = Me.lbl_esg_country
    End With
    With Me.txt_is_counsel_invoice_received_id
        .field_name = "CounselInvoiceReceived"
        .field_name_in_recordset = "CounselInvoiceReceived"
        If load.system_info.app_continent = load.system_info.continents.americas Or load.deal_details.make_all_fields_visible_for_debugging Then
            Set .above_field = Me.lbl_esg_company
        Else
            Set .above_field = Me.lbl_sanction_checks
        End If
    End With
    With Me.txt_is_vdr_received_id
        .field_name = "VDRReceived"
        .field_name_in_recordset = "VDRReceived"
        Set .above_field = Me.txt_is_counsel_invoice_received_id
    End With
    With Me.txt_vdr_password
        .field_name = "VDRPassword"
        .field_name_in_recordset = "VDRPassword"
        Set .above_field = Me.txt_is_vdr_received_id
    End With
    With Me.txt_closing_set_received
        .field_name = "closing_set_received_id"
        .field_name_in_recordset = "closing_set_received_id"
        .visible_in_europe = False
        Set .above_field = Me.txt_vdr_password
    End With
    With Me.txt_is_rr_completed_id
        .field_name = "rr_done"
        .field_name_in_recordset = "rr_done"
        If load.system_info.app_continent = load.system_info.continents.americas Or load.deal_details.make_all_fields_visible_for_debugging Then
            Set .above_field = Me.lbl_closing_set_received
        Else
            Set .above_field = Me.lbl_vdr_password
        End If
        
    End With
    With Me.txt_are_emails_filed_id
        .field_name = "c5_emails_filed"
        .field_name_in_recordset = "are_emails_filed_id"
        Set .above_field = Me.txt_is_rr_completed_id
    End With
    With Me.txt_is_closing_booked_id
        .field_name = "ClosingBooked"
        .field_name_in_recordset = "ClosingBooked"
        Set .above_field = Me.txt_are_emails_filed_id
    End With
    
    Dim field_left As Long
    field_left = 7200
    For Each form_field In Me.col_txt_fields_for_closing_info
        With form_field
            .field_left = field_left
            If Me.header_closing_info.user_has_expanded_section = False Then
                .field_visible = False
            End If
            .field_top = .above_field.field_bottom
        End With
    Next form_field
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub

Public Sub refresh_check_points_section()
    Const proc_name As String = "cls_form_deal_details.refresh_check_points_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim field_counter As Integer
    
    'header
    With Me.header_closing_info
        .field_top = Me.header_buyer_and_seller.field_top + Me.header_buyer_and_seller.section_height
    End With
    
    'command button for collapsing/expanding
    With Me.cmd_expand_closing_info_section
        .field_top = Me.header_closing_info.field_top
        If Me.header_closing_info.user_has_expanded_section = True Then
            .field_caption = "collapse"
        Else
            .field_caption = "expand"
        End If
    End With
    
    For Each form_field In Me.col_labels_for_closing_info
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
    
    For Each form_field In Me.col_txt_fields_for_closing_info
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init_deal_and_target_info_section()
    Const proc_name As String = "cls_form_deal_details.init_deal_and_target_info_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim above_field As cls_field
    
    With Me.header_deal_and_target_info
        .field_name = "header_deal_and_target_info"
        .field_top = Me.rows_first_top
        .field_left = Me.columns_second_left
        If .user_has_expanded_section = False Then
            .is_collapsed = False
        End If
        .section_height = utilities.twips_converter(9, "cm")
        If load.system_info.app_continent = load.system_info.continents.americas Then
            .section_height = utilities.twips_converter(10, "cm")
        End If
    End With
        
    For Each form_field In Me.col_labels_for_deal_and_target_info
        With form_field
            .field_left = Me.header_deal_and_target_info.field_left + 50
            .field_width = 2600
        End With
    Next form_field
    
    'labels
    With Me.lbl_risk_currency_id
        .field_name = "lbl_risk_currency_id"
        .field_top = Me.header_deal_and_target_info.field_top + Me.header_deal_and_target_info.field_height
        Set above_field = Me.lbl_risk_currency_id
    End With
    With Me.lbl_ev
        .field_name = "lbl_ev"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_ev
    End With
    With Me.lbl_target_legal_name
        .field_name = "lbl_target_legal_name"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_target_legal_name
    End With
    With Me.lbl_target_business_name
        .field_name = "lbl_target_business_name"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_target_business_name
    End With
    With Me.lbl_target_registered_country_id
        .field_name = "lbl_target_registered_country_id"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_target_registered_country_id
    End With
    With Me.lbl_target_main_country_id
        .field_name = "lbl_target_main_country_id"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_target_main_country_id
    End With
    With Me.lbl_target_super_sector_id
        .field_name = "lbl_target_super_sector_id"
        .field_top = above_field.field_top + above_field.field_height
        Set above_field = Me.lbl_target_super_sector_id
    End With
    With Me.lbl_target_sub_sector_id
        .field_name = "lbl_target_sub_sector_id"
        .field_top = above_field.field_bottom
        Set above_field = Me.lbl_target_sub_sector_id
    End With
    With Me.lbl_target_sic_code
        .field_name = "lbl_target_sic_code"
        .field_top = above_field.field_top + above_field.field_height
        If load.system_info.app_continent = load.system_info.continents.americas Or load.deal_details.make_all_fields_visible_for_debugging Then
            Set above_field = Me.lbl_target_sic_code
        End If
    End With
    With Me.lbl_target_naic_code
        .field_name = "lbl_target_naic_code"
        .field_top = above_field.field_bottom
        If load.system_info.app_continent = load.system_info.continents.americas Or load.deal_details.make_all_fields_visible_for_debugging Then
            Set above_field = Me.lbl_target_naic_code
        End If
    End With
    With Me.lbl_risk_signing_date
        .field_name = "lbl_risk_signing_date"
        .field_top = above_field.field_bottom
        Set above_field = Me.lbl_risk_signing_date
    End With
    With Me.lbl_risk_closing_date
        .field_name = "lbl_risk_closing_date"
        .field_top = above_field.field_bottom
        .field_caption = "Closing Date"
        Set above_field = Me.lbl_risk_closing_date
    End With
    With Me.lbl_target_desc
        .field_name = "lbl_target_desc"
        .field_top = above_field.field_bottom
        Set above_field = Me.lbl_target_desc
    End With
    
    'txts
    Dim field_left As Integer
    field_left = 6700
    With Me.txt_risk_currency_id
        .field_name = "deal_currency"
        .field_name_in_recordset = "deal_currency"
        .field_top = Me.header_deal_and_target_info.field_bottom
        .field_left = field_left
        Set above_field = Me.txt_risk_currency_id
    End With
    With Me.txt_ev
        .field_name = "ev"
        .field_name_in_recordset = "ev"
        .field_top = above_field.field_bottom
        .field_left = field_left
        Set above_field = Me.txt_ev
    End With
    With Me.txt_target_legal_name
        .field_name = "target_legal_name"
        .field_name_in_recordset = "target_legal_name"
        .field_top = above_field.field_top + above_field.field_height
        .field_left = field_left
        .field_width = utilities.twips_converter(6, "cm")
        Set above_field = Me.txt_target_legal_name
    End With
    With Me.txt_target_business_name
        .field_name = "target_business_name"
        .field_name_in_recordset = "target_business_name"
        .field_top = above_field.field_top + above_field.field_height
        .field_left = field_left
        .field_width = utilities.twips_converter(6, "cm")
        Set above_field = Me.txt_target_business_name
    End With
    With Me.txt_target_registered_country_id
        .field_name = "target_country"
        .field_name_in_recordset = "target_country"
        .field_top = above_field.field_top + above_field.field_height
        .field_left = field_left
        Set above_field = Me.txt_target_registered_country_id
    End With
    With Me.txt_target_main_country_id
        .field_name = "target_main_jurisdiction_id"
        .field_name_in_recordset = "target_main_jurisdiction_id"
        .field_top = above_field.field_top + above_field.field_height
        .field_left = field_left
        Set above_field = Me.txt_target_main_country_id
    End With
    With Me.txt_target_super_sector_id
        .field_name = "target_super_sector_id"
        .field_name_add_many_deals_f = "target_super_sector_id"
        .field_name_in_recordset = "target_super_sector_id"
        .field_name_in_table = "target_super_sector_id"
        .field_top = above_field.field_top + above_field.field_height
        .field_left = field_left
        Set above_field = Me.txt_target_super_sector_id
    End With
    With Me.txt_target_sub_sector_id
        .field_name = "target_sub_sector_id"
        .field_name_add_many_deals_f = "target_sub_sector_id"
        .field_name_in_recordset = "target_sub_sector_id"
        .field_name_in_table = "target_sub_sector_id"
        .field_top = above_field.field_top + above_field.field_height
        .field_left = field_left
        Set above_field = Me.txt_target_sub_sector_id
    End With
    With Me.txt_target_sic_code
        .field_name = "target_sic_code"
        .field_name_in_recordset = "target_sic_code"
        .field_top = above_field.field_top + above_field.field_height
        .field_left = field_left
        If load.system_info.app_continent = load.system_info.continents.americas Or load.deal_details.make_all_fields_visible_for_debugging Then
            Set above_field = Me.txt_target_sic_code
        End If
    End With
    With Me.txt_target_naic_code
        .field_name = "target_naic_code"
        .field_name_in_recordset = "target_naic_code"
        .field_top = above_field.field_top + above_field.field_height
        .field_left = field_left
        If load.system_info.app_continent = load.system_info.continents.americas Or load.deal_details.make_all_fields_visible_for_debugging Then
            Set above_field = Me.txt_target_naic_code
        End If
    End With
    With Me.txt_risk_signing_date
        .field_name = "spa_signing_date"
        .field_name_in_recordset = "spa_signing_date"
        .field_top = above_field.field_top + above_field.field_height
        .field_left = field_left
        Set above_field = Me.txt_risk_signing_date
    End With
    With Me.txt_risk_closing_date
        .field_name = "closing_date"
        .field_name_in_recordset = "closing_date"
        .field_top = above_field.field_bottom
        .field_left = field_left
        Set above_field = Me.txt_risk_closing_date
    End With
    With Me.txt_target_desc
        .field_name = "target_desc"
        .field_name_in_recordset = "target_desc"
        .field_top = above_field.field_bottom
        .field_left = field_left
        .field_width = 3500
        .field_height = 1200
        Set above_field = Me.txt_target_desc
    End With
    
    Set form_field = Nothing
    Set above_field = Nothing
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub refresh_deal_and_target_info_section()
    Dim form_field As cls_field
    
End Sub
Public Sub init_financial_info_section()
    Const proc_name As String = "cls_form_deal_details.init_financial_info_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    
    'header
    With Me.header_financial_info
        .field_name = "header_financial_info"
        .field_top = Me.header_submission_notes.field_top + Me.header_submission_notes.section_height
        .field_left = Me.columns_third_left
        .is_collapsed = False
        .section_height = utilities.twips_converter(6, "cm")
    End With
    
    'labels
    With Me.lbl_amounts
        .field_name = "lbl_amounts"
        .field_left = Me.header_financial_info.field_left + 50 + utilities.twips_converter(3.5, "cm")
        .field_width = utilities.twips_converter(2.6, "cm")
        Set .above_field = Me.header_financial_info
    End With
    With Me.lbl_due_dates
        .field_name = "lbl_due_dates"
        .field_left = Me.lbl_amounts.field_left + Me.lbl_amounts.field_width
        .field_width = utilities.twips_converter(2.6, "cm")
        Set .above_field = Me.header_financial_info
    End With
    With Me.lbl_received_dates
        .field_name = "lbl_received_dates"
        .field_left = Me.lbl_due_dates.field_left + Me.lbl_due_dates.field_width
        .field_width = utilities.twips_converter(2.6, "cm")
        Set .above_field = Me.header_financial_info
    End With
    With Me.lbl_uw_fee
        .field_name = "lbl_uw_fee"
        .field_left = Me.header_financial_info.field_left + 50
        .field_width = utilities.twips_converter(3.6, "cm")
        Set .above_field = Me.lbl_received_dates
    End With
    With Me.lbl_counsel_fee
        .field_name = "lbl_counsel_fee"
        .field_left = Me.header_financial_info.field_left + 50
        .field_width = utilities.twips_converter(3.6, "cm")
        Set .above_field = Me.lbl_uw_fee
    End With
    With Me.lbl_net_uw_fee
        .field_name = "lbl_net_uw_fee"
        .field_left = Me.header_financial_info.field_left + 50
        .field_width = utilities.twips_converter(3.6, "cm")
        Set .above_field = Me.lbl_counsel_fee
    End With
    With Me.lbl_signing_premium
        .field_name = "lbl_signing_premium"
        .field_left = Me.header_financial_info.field_left + 50
        .field_width = utilities.twips_converter(3.6, "cm")
        .field_top = Me.header_financial_info.field_top + utilities.twips_converter(3.5, "cm")
        Set .above_field = Me.lbl_net_uw_fee
    End With
    With Me.lbl_closing_premium
        .field_name = "lbl_closing_premium"
        .field_left = Me.header_financial_info.field_left + 50
        .field_width = utilities.twips_converter(3.6, "cm")
        Set .above_field = Me.lbl_signing_premium
    End With
    With Me.lbl_total_premium
        .field_name = "lbl_total_premium"
        .field_left = Me.header_financial_info.field_left + 50
        .field_width = utilities.twips_converter(3.6, "cm")
        Set .above_field = Me.lbl_closing_premium
    End With

    For Each form_field In Me.col_labels_for_financial_info_section
        With form_field
            .field_visible = True
            If .field_top = 0 Then .field_top = .above_field.field_top + .above_field.field_height
        End With
    Next form_field
    
    'text fields
    'reset field_top
    For Each form_field In Me.col_txt_fields_for_financial_info_section
        With form_field
            .field_top = 0
            .field_width = utilities.twips_converter(2.6, "cm")
        End With
    Next form_field
    
    With Me.txt_uw_fee
        .field_name = "uw_fee"
        .field_name_in_recordset = "uw_fee_amount"
        .field_left = Me.lbl_amounts.field_left
        Set .above_field = Me.lbl_amounts
    End With
    With Me.txt_uw_fee_due_date
        .field_name = "uw_fee_due_date"
        .field_name_in_recordset = "uw_fee_due_date"
        .field_left = Me.lbl_due_dates.field_left
        Set .above_field = Me.lbl_due_dates
    End With
    With Me.txt_uw_fee_received_date
        .field_name = "uw_fee_received_date"
        .field_name_in_recordset = "uw_fee_received_date"
        .field_left = Me.lbl_received_dates.field_left
        Set .above_field = Me.lbl_received_dates
    End With
    With Me.txt_counsel_fee
        .field_name = "counsel_fee"
        .field_name_in_recordset = "counsel_fee_amount"
        .field_left = Me.lbl_amounts.field_left
        Set .above_field = Me.txt_uw_fee
    End With
    With Me.txt_counsel_fee_due_date
        .field_name = "counsel_fee_due_date"
        .field_name_in_recordset = "counsel_fee_due_date"
        .field_left = Me.lbl_due_dates.field_left
        Set .above_field = Me.txt_uw_fee
    End With
    With Me.txt_counsel_fee_received_date
        .field_name = "counsel_fee_received_date"
        .field_name_in_recordset = "counsel_fee_received_date"
        .field_left = Me.lbl_received_dates.field_left
        Set .above_field = Me.txt_uw_fee
    End With
    With Me.txt_net_uw_fee
        .field_name = "net_uw_fee"
        .field_name_in_recordset = "-1"
        .field_left = Me.lbl_amounts.field_left
        Set .above_field = Me.lbl_counsel_fee
    End With
    With Me.txt_signing_premium_amount
        .field_name = "signing_premium_amount"
        .field_name_in_recordset = "signing_invoice_amount"
        .field_left = Me.lbl_amounts.field_left
        .field_top = Me.header_financial_info.field_top + utilities.twips_converter(3.5, "cm")
        Set .above_field = Me.lbl_net_uw_fee
    End With
    With Me.txt_signing_premium_due_date
        .field_name = "signing_premium_due_date"
        .field_name_in_recordset = "signing_premium_due_date"
        .field_left = Me.lbl_due_dates.field_left
        .field_top = Me.header_financial_info.field_top + utilities.twips_converter(3.5, "cm")
        Set .above_field = Me.lbl_net_uw_fee
    End With
    With Me.txt_signing_premium_received_date
        .field_name = "signing_premium_received_date"
        .field_name_in_recordset = "signing_premium_received_date"
        .field_left = Me.lbl_received_dates.field_left
        .field_top = Me.header_financial_info.field_top + utilities.twips_converter(3.5, "cm")
        Set .above_field = Me.lbl_net_uw_fee
    End With
    With Me.txt_closing_premium_amount
        .field_name = "closing_premium_amount"
        .field_name_in_recordset = "-1"
        .field_left = Me.lbl_amounts.field_left
        Set .above_field = Me.txt_signing_premium_amount
    End With
    With Me.txt_closing_premium_due_date
        .field_name = "closing_premium_due_date"
        .field_name_in_recordset = "closing_premium_due_date"
        .field_left = Me.lbl_due_dates.field_left
        Set .above_field = Me.txt_signing_premium_amount
    End With
    With Me.txt_closing_premium_received_date
        .field_name = "closing_premium_received_date"
        .field_name_in_recordset = "closing_premium_received_date"
        .field_left = Me.lbl_received_dates.field_left
        Set .above_field = Me.txt_signing_premium_amount
    End With
    With Me.txt_total_premium
        .field_name = "total_premium"
        .field_name_in_recordset = "-1"
        .field_left = Me.lbl_amounts.field_left
        Set .above_field = Me.txt_closing_premium_amount
    End With
    
    For Each form_field In Me.col_txt_fields_for_financial_info_section
        With form_field
            .field_visible = True
            If .field_top = 0 Then .field_top = .above_field.field_top + .above_field.field_height
        End With
    Next form_field
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub refresh_financial_info_section()
    Const proc_name As String = "cls_form_deal_deatils.refresh_financial_info_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim field_counter As Integer
    
    'header
    With Me.header_financial_info
        .field_top = Me.header_submission_notes.field_top + Me.header_submission_notes.section_height
    End With
    
    'labels
    For Each form_field In Me.col_labels_for_financial_info_section
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
    
    With Me.lbl_signing_premium
        .field_top = Me.header_financial_info.field_top + utilities.twips_converter(3.5, "cm")
    End With
    With Me.lbl_closing_premium
        .field_top = Me.lbl_signing_premium.field_bottom
    End With
    With Me.lbl_total_premium
        .field_top = Me.lbl_closing_premium.field_bottom
    End With
    
    'text fields
    For Each form_field In Me.col_txt_fields_for_financial_info_section
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
    
    With Me.txt_signing_premium_amount
        .field_top = Me.header_financial_info.field_top + utilities.twips_converter(3.5, "cm")
    End With
    With Me.txt_signing_premium_due_date
        .field_top = Me.header_financial_info.field_top + utilities.twips_converter(3.5, "cm")
    End With
    With Me.txt_signing_premium_received_date
        .field_top = Me.header_financial_info.field_top + utilities.twips_converter(3.5, "cm")
    End With
    With Me.txt_closing_premium_amount
        .field_top = Me.lbl_signing_premium.field_bottom
    End With
    With Me.txt_closing_premium_due_date
        .field_top = Me.lbl_signing_premium.field_bottom
    End With
    With Me.txt_closing_premium_received_date
        .field_top = Me.lbl_signing_premium.field_bottom
    End With
    With Me.txt_total_premium
        .field_top = Me.lbl_closing_premium.field_bottom
    End With
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init_insurance_info_section()
    Const proc_name As String = "cls_form_deal_details.init_insurance_info_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    
    'header
    With Me.header_insurance_info
        .field_name = "header_insurance_info"
        .field_top = Me.rows_first_top + Me.header_nda_stage.section_height
        .field_left = Me.columns_first_left
        .is_collapsed = False
        If load.system_info.app_continent = load.system_info.continents.americas Then
            .section_height = utilities.twips_converter(11.2, "cm")
        Else
            .section_height = utilities.twips_converter(9, "cm")
        End If
    End With
    
    'labels
    For Each form_field In Me.col_labels_for_insurance_info
        With form_field
            .field_left = Me.header_insurance_info.field_left + 50
            .field_width = 2100
        End With
    Next form_field
    
    With Me.lbl_risk_type_id
        .field_name = "lbl_risk_type_id"
        Set .above_field = Me.header_insurance_info
    End With
    With Me.lbl_nbi_prepper
        .field_name = "lbl_nbi_prepper"
        Set .above_field = Me.lbl_risk_type_id
    End With
    With Me.lbl_primary_uw_id
        .field_name = "lbl_primary_uw_id"
        Set .above_field = Me.lbl_nbi_prepper
    End With
    With Me.lbl_secondary_uw_id
        .field_name = "lbl_secondary_uw_id"
        Set .above_field = Me.lbl_primary_uw_id
    End With
    With Me.lbl_analyst_id
        .field_name = "lbl_analyst_id"
        Set .above_field = Me.lbl_secondary_uw_id
    End With
    With Me.lbl_internal_approver_quote_id
        .field_name = "lbl_internal_approver_quote_id"
        Set .above_field = Me.lbl_analyst_id
    End With
    With Me.lbl_internal_approver_binding_id
        .field_name = "lbl_internal_approver_binding_id"
        Set .above_field = Me.lbl_internal_approver_quote_id
    End With
    With Me.lbl_risk_feel_id
        .field_name = "lbl_risk_feel_id"
        Set .above_field = Me.lbl_analyst_id
        If load.system_info.app_continent = load.system_info.continents.americas Or Me.make_all_fields_visible_for_debugging = True Then
            Set .above_field = Me.lbl_internal_approver_binding_id
        End If
    End With
    With Me.lbl_primary_or_xs_id
        .field_name = "lbl_primary_or_xs_id"
        Set .above_field = Me.lbl_risk_feel_id
    End With
    With Me.lbl_budget_home_id
        .field_name = "lbl_budget_home_id"
        Set .above_field = Me.lbl_primary_or_xs_id
    End With
    With Me.lbl_program_limit
        .field_name = "lbl_program_limit"
        If load.system_info.app_continent = load.system_info.continents.americas Then
            .field_caption = "Total tower size"
        End If
        Set .above_field = Me.lbl_budget_home_id
    End With
    With Me.lbl_primary_insurer_id
        .field_name = "lbl_primary_insurer_id"
        Set .above_field = Me.lbl_program_limit
    End With
    With Me.lbl_retention
        .field_name = "lbl_retention"
        Set .above_field = Me.lbl_primary_insurer_id
    End With
    With Me.lbl_retention_dropped
        .field_name = "lbl_retention_dropped"
        Set .above_field = Me.lbl_retention
    End With
    With Me.lbl_drop_period
        .field_name = "lbl_drop_period"
        Set .above_field = Me.lbl_retention_dropped
    End With
    With Me.lbl_inception_date
        .field_name = "lbl_inception_date"
        If load.system_info.app_continent = load.system_info.continents.americas Or Me.make_all_fields_visible_for_debugging = True Then
            Set .above_field = Me.lbl_drop_period
        Else
            Set .above_field = Me.lbl_retention
        End If
    End With
    With Me.lbl_policy_period
        .field_name = "lbl_policy_period"
        .field_top = Me.lbl_retention.field_bottom
        .field_height = utilities.twips_converter(0.8, "cm")
        Set .above_field = Me.lbl_inception_date
    End With
    
    For Each form_field In Me.col_labels_for_insurance_info
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
        
    'text fields
    For Each form_field In Me.col_txt_fields_for_insurance_info
        With form_field
            .field_left = Me.header_insurance_info.field_left + Me.lbl_risk_type_id.field_width
        End With
    Next form_field
    
    With Me.txt_risk_type_id
        .field_name = "risk_type_id"
        .field_name_in_recordset = "risk_type_id"
        .field_left = 1200
        .field_width = 2700
        Set .above_field = Me.header_insurance_info
    End With
    With Me.txt_nbi_prepper_id
        .field_name = "nbi_prepper_id"
        .field_name_add_many_deals_f = "nbi_prepper_id"
        .field_name_in_recordset = "nbi_prepper_id"
        .field_name_in_table = "nbi_prepper"
        .field_left = 1200
        .field_width = 900
        Set .above_field = Me.txt_risk_type_id
    End With
    With Me.txt_nbi_prepper_full_name
        .field_name = "nbi_prepper_full_name"
        .field_name_in_recordset = "nbi_prepper"
        Set .above_field = Me.txt_risk_type_id
    End With
    With Me.txt_primary_uw_id
        .field_name = "primary_uw"
        .field_name_in_recordset = "primary_uw_id"
        .field_left = 1200
        .field_width = Me.txt_nbi_prepper_id.field_width
         Set .above_field = Me.txt_nbi_prepper_id
    End With
    With Me.txt_primary_uw_full_name
        .field_name = "primary_uw_full_name"
        .field_name_in_recordset = "primary_uw_full_name"
        Set .above_field = Me.txt_nbi_prepper_id
    End With
    With Me.txt_secondary_uw_id
        .field_name = "secondary_uw"
        .field_name_in_recordset = "secondary_uw_id"
        .field_left = 1200
        .field_width = Me.txt_nbi_prepper_id.field_width
        Set .above_field = Me.lbl_primary_uw_id
    End With
    With Me.txt_secondary_uw_full_name
        .field_name = "secondary_uw_full_name"
        .field_name_in_recordset = "primary_uw_full_name"
        Set .above_field = Me.lbl_primary_uw_id
    End With
    With Me.txt_analyst_id
        .field_name = "analyst_id"
        .field_name_in_recordset = "analyst_id"
        .field_name_in_table = "analyst_id"
        .field_left = 1200
        .field_width = Me.txt_nbi_prepper_id.field_width
        Set .above_field = Me.lbl_secondary_uw_id
    End With
    With Me.txt_analyst_full_name
        .field_name = "analyst_full_name"
        .field_name_in_recordset = "analyst_full_name"
        Set .above_field = Me.lbl_secondary_uw_id
    End With
    With Me.txt_internal_approver_quote_id
        .field_name = "internal_approver_quote_id"
        .field_name_in_recordset = "internal_approver_quote_id"
        .field_left = 1200
        .field_width = Me.txt_nbi_prepper_id.field_width
        Set .above_field = Me.lbl_analyst_id
    End With
    With Me.txt_internal_approver_quote_full_name
        .field_name = "internal_approver_quote_full_name"
        .field_name_in_recordset = "internal_approver_quote_full_name"
        Set .above_field = Me.lbl_analyst_id
    End With
    With Me.txt_internal_approver_binding_id
        .field_name = "internal_approver_binding_id"
        .field_name_in_recordset = "internal_approver_binding_id"
        .field_left = 1200
        .field_width = Me.txt_nbi_prepper_id.field_width
        Set .above_field = Me.lbl_internal_approver_quote_id
    End With
    With Me.txt_internal_approver_binding_full_name
        .field_name = "internal_approver_binding_full_name"
        .field_name_in_recordset = "internal_approver_binding_full_name"
        Set .above_field = Me.lbl_internal_approver_quote_id
    End With
    With Me.txt_risk_feel_id
        .field_name = "risk_feel_id"
        .field_name_in_recordset = "risk_feel_id"
        Set .above_field = Me.lbl_analyst_id
        If load.system_info.app_continent = load.system_info.continents.americas Or Me.make_all_fields_visible_for_debugging = True Then
            Set .above_field = Me.lbl_internal_approver_binding_id
        End If
    End With
    With Me.txt_primary_or_xs_id
        .field_name = "primary_or_xs_id"
        .field_name_in_recordset = "primary_or_xs_id"
        Set .above_field = Me.txt_risk_feel_id
    End With
    With Me.txt_budget_home_id
        .field_name = "budget_home_id"
        .field_name_add_many_deals_f = "budget_home"
        .field_name_in_table = "budget_home_id"
        .field_name_in_recordset = "budget_home_id"
        
        Set .above_field = Me.txt_primary_or_xs_id
    End With
    With Me.txt_program_limit
        .field_name = "program_limit"
        .field_name_in_recordset = "program_limit"
        Set .above_field = Me.txt_budget_home_id
    End With
    With Me.txt_primary_insurer_id
        .field_name = "primary_insurer"
        .field_name_in_recordset = "primary_insurer"
        Set .above_field = Me.txt_program_limit
    End With
    With Me.txt_retention
        .field_name = "retention"
        .field_name_in_recordset = "retention"
        Set .above_field = Me.txt_primary_insurer_id
    End With
    With Me.txt_retention_dropped
        .field_name = "retention_dropped"
        .field_name_in_recordset = "drop_end"
        Set .above_field = Me.lbl_retention
    End With
    With Me.txt_drop_period
        .field_name = "drop_period"
        .field_name_in_recordset = "drop_period"
        Set .above_field = Me.lbl_retention_dropped
    End With
    With Me.txt_inception_date
        .field_name = "inception_date"
        .field_name_in_recordset = "inception_date"
        If load.system_info.app_continent = load.system_info.continents.americas Or Me.make_all_fields_visible_for_debugging = True Then
            Set .above_field = Me.lbl_drop_period
        Else
            Set .above_field = Me.lbl_retention
        End If
    End With
    With Me.txt_policy_period
        .field_name = "policy_period_in_months"
        .field_name_in_recordset = "policy_period_in_months"
        Set .above_field = Me.lbl_inception_date
    End With
    
    For Each form_field In Me.col_txt_fields_for_insurance_info
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub refresh_insurance_info_section()
    Dim form_field As cls_field
    
    'header
    With Me.header_insurance_info
        .field_top = Me.rows_first_top + Me.header_nda_stage.section_height
    End With
    
    'labels
    For Each form_field In Me.col_labels_for_insurance_info
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
    
    'text fields
    For Each form_field In Me.col_txt_fields_for_insurance_info
        With form_field
            .field_top = .above_field.field_bottom
        End With
    Next form_field
    
End Sub
Public Sub init_submisson_notes_section()
    Const proc_name As String = "cls_form_deal_details.init_submisson_notes_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    
    'header
    With Me.header_submission_notes
        .field_name = "header_submission_notes"
        .field_top = Me.header_advisors.field_top + Me.header_advisors.section_height
        .field_left = Me.columns_third_left
        .is_collapsed = True
        If Me.header_submission_notes.user_has_expanded_section = False Then
            .section_height = utilities.twips_converter(1.5, "cm")
        Else
            .section_height = utilities.twips_converter(12, "cm")
        End If
    End With
    
    'expanders
    With Me.cmd_expand_submission_notes_section
        .field_name = "cmd_expand_submission_notes_section"
        .field_caption = "expand"
        .field_top = Me.header_submission_notes.field_top
        .field_left = Me.header_submission_notes.field_left + 2500
    End With
    
    'labels and texts
    For Each form_field In Me.col_txt_fields_for_submission_notes_section
        With form_field
            .field_height = utilities.twips_converter(2.6, "cm")
            .field_left = Me.header_submission_notes.field_left + 50
            .field_width = utilities.twips_converter(10, "cm")
            If Me.header_submission_notes.user_has_expanded_section = False Then
                .field_visible = False
            End If
        End With
    Next form_field
    
    For Each form_field In Me.col_labels_for_submission_notes_section
        With form_field
            .field_left = Me.header_submission_notes.field_left + 50
            .field_height = utilities.twips_converter(0.6, "cm")
            If Me.header_submission_notes.user_has_expanded_section = False Then
                .field_visible = False
            End If
        End With
    Next form_field
    
    With Me.lbl_reviewer_notes
        .field_name = "lbl_reviewer_notes"
        .field_top = Me.header_submission_notes.field_bottom
    End With
    With Me.txt_reviewer_notes
        .field_name = "reviewer_notes"
        .field_name_in_recordset = "reviewer_notes"
        .field_top = Me.lbl_reviewer_notes.field_bottom
    End With
    With Me.lbl_deal_info
        .field_name = "lbl_deal_info"
        .field_top = Me.txt_reviewer_notes.field_bottom + utilities.twips_converter(0.5, "cm")
    End With
    With Me.txt_deal_info
        .field_name = "deal_info"
        .field_name_in_recordset = "deal_info"
        .field_top = Me.lbl_deal_info.field_bottom
    End With
    With Me.lbl_submission_limits
        .field_name = "lbl_submission_limits"
        .field_top = Me.txt_deal_info.field_bottom + utilities.twips_converter(0.5, "cm")
    End With
    With Me.txt_submission_limits
        .field_name = "submission_limits"
        .field_name_in_recordset = "submission_limits"
        .field_top = Me.lbl_submission_limits.field_bottom
    End With
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub refresh_submission_notes_section()
    Const proc_name As String = "cls_form_deal_details.refresh_submission_notes_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim field_counter As Integer
    
    'header
    With Me.header_submission_notes
        .field_top = Me.header_advisors.field_top + Me.header_advisors.section_height
    End With
    
    'expander
    With Me.cmd_expand_submission_notes_section
        .field_top = Me.header_submission_notes.field_top
        If Me.header_submission_notes.user_has_expanded_section = True Then
            .field_caption = "collapse"
            Me.header_submission_notes.section_height = utilities.twips_converter(12, "cm")
        Else
            .field_caption = "expand"
            Me.header_submission_notes.section_height = utilities.twips_converter(1.5, "cm")
        End If
    End With
    
    'labels
    For Each form_field In Me.col_labels_for_submission_notes_section
        With form_field
            .field_visible = False
            If Me.header_submission_notes.user_has_expanded_section = True Then .field_visible = True
        End With
    Next form_field
    
    'text fields
    For Each form_field In Me.col_txt_fields_for_submission_notes_section
        With form_field
            .field_visible = False
            If Me.header_submission_notes.user_has_expanded_section = True Then .field_visible = True
        End With
    Next form_field
    
    With Me.lbl_reviewer_notes
        .field_top = Me.header_submission_notes.field_bottom
    End With
    With Me.txt_reviewer_notes
        .field_top = Me.lbl_reviewer_notes.field_bottom
    End With
    With Me.lbl_deal_info
        .field_top = Me.txt_reviewer_notes.field_bottom + utilities.twips_converter(0.5, "cm")
    End With
    With Me.txt_deal_info
        .field_top = Me.lbl_deal_info.field_bottom
    End With
    With Me.lbl_submission_limits
        .field_top = Me.txt_deal_info.field_bottom + utilities.twips_converter(0.5, "cm")
    End With
    With Me.txt_submission_limits
        .field_top = Me.lbl_submission_limits.field_bottom
    End With
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init_misc_section()
    Const proc_name As String = "cls_form_deal_details.init_misc_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim field_left As Long
    Dim section_label_width As Long
        
    'header
    With Me.header_misc
        .field_name = "header_misc"
        .field_top = Me.header_financial_info.section_height + Me.header_financial_info.field_top
        .field_left = Me.columns_third_left
        .section_height = utilities.twips_converter(12.8, "cm")
        .is_collapsed = False
    End With
        
    field_left = Me.header_misc.field_left + 50
    section_label_width = utilities.twips_converter(5.5, "cm")
    
    'surplus broker firm
    With Me.lbl_surplus_broker_firm
        .field_name = "lbl_surplus_broker_firm"
        .field_top = Me.header_misc.field_bottom
        .field_left = field_left
        .field_width = section_label_width
    End With
    With Me.txt_surplus_broker_firm
        .field_name = "surplus_broker_firm"
        .field_name_in_recordset = "surplus_broker_firm"
        .field_top = Me.header_misc.field_bottom
        .field_left = Me.lbl_surplus_broker_firm.field_right
        .field_width = section_label_width
    End With
    
    'surplus broker person
    With Me.lbl_surplus_broker_person
        .field_name = "lbl_surplus_broker_person"
        .field_top = Me.lbl_surplus_broker_firm.field_bottom
        .field_left = field_left
        .field_width = section_label_width
    End With
    With Me.txt_surplus_broker_person
        .field_name = "surplus_broker_person"
        .field_name_in_recordset = "surplus_broker_person"
        .field_top = Me.lbl_surplus_broker_firm.field_bottom
        .field_left = Me.lbl_surplus_broker_person.field_right
        .field_width = section_label_width
    End With
    
    'surplus broker license no
    With Me.lbl_surplus_broker_license_no
        .field_name = "lbl_surplus_broker_license_no"
        .field_top = Me.lbl_surplus_broker_person.field_bottom
        .field_left = field_left
        .field_width = section_label_width
    End With
    With Me.txt_surplus_broker_license_no
        .field_name = "surplus_broker_license_no"
        .field_name_in_recordset = "surplus_broker_license_no"
        .field_top = Me.lbl_surplus_broker_person.field_bottom
        .field_left = Me.lbl_surplus_broker_license_no.field_right
        .field_width = section_label_width
    End With
    
    'us producer number aka National Producer Number (NPN)
    With Me.lbl_us_producer_no
        .field_name = "lbl_us_producer_no"
        .field_top = Me.lbl_surplus_broker_license_no.field_bottom
        .field_left = field_left
        .field_width = section_label_width
    End With
    With Me.txt_us_producer_no
        .field_name = "us_producer_no"
        .field_name_in_recordset = "us_producer_no"
        .field_top = Me.lbl_us_producer_no.field_top
        .field_left = Me.lbl_surplus_broker_license_no.field_right
        .field_width = section_label_width
    End With
    
    'is underwritten
    With Me.lbl_is_underwritten_id
        .field_name = "lbl_is_underwritten_id"
        .field_top = Me.lbl_us_producer_no.field_bottom
        .field_left = field_left
        .field_width = section_label_width
    End With
    With Me.txt_is_underwritten_id
        .field_name = "is_underwritten"
        .field_name_in_recordset = "is_underwritten"
        .field_top = Me.lbl_is_underwritten_id.field_top
        .field_left = Me.lbl_is_underwritten_id.field_right
    End With
    With Me.lbl_comments
        .field_name = "lbl_comments"
        .field_top = Me.lbl_is_underwritten_id.field_bottom + utilities.twips_converter(0.5, "cm")
        .field_left = field_left
        .field_width = section_label_width
    End With
    With Me.txt_comments
        .field_name = "comments"
        .field_name_in_recordset = "comments"
        .field_top = Me.lbl_comments.field_bottom
        .field_left = Me.lbl_comments.field_left
        .field_height = utilities.twips_converter(3, "cm")
        .field_width = utilities.twips_converter(10, "cm")
    End With
    With Me.lbl_stella_message_board
        .field_name = "lbl_stella_message_board"
        .field_top = Me.txt_comments.field_bottom + utilities.twips_converter(0.5, "cm")
        .field_left = field_left
        .field_width = section_label_width
    End With
    With Me.txt_stella_message_board
        .field_name = "message_board"
        .field_name_in_recordset = "message_board"
        .field_top = Me.lbl_stella_message_board.field_bottom
        .field_left = Me.lbl_stella_message_board.field_left
        .field_height = utilities.twips_converter(3, "cm")
        .field_width = Me.txt_comments.field_width
    End With
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub refresh_misc_section()
    Const proc_name As String = "cls_form_deal_details.refresh_misc_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field, field_counter As Integer

    With Me.header_misc
        .field_top = Me.header_financial_info.field_top + Me.header_financial_info.section_height
    End With
    
    'labels and text fields
    With Me.lbl_surplus_broker_firm
        .field_top = Me.header_misc.field_bottom
    End With
    With Me.txt_surplus_broker_firm
        .field_top = Me.header_misc.field_bottom
    End With
    With Me.lbl_surplus_broker_person
        .field_top = Me.lbl_surplus_broker_firm.field_bottom
    End With
    With Me.txt_surplus_broker_person
        .field_top = Me.lbl_surplus_broker_firm.field_bottom
    End With
    With Me.lbl_surplus_broker_license_no
        .field_top = Me.lbl_surplus_broker_person.field_bottom
    End With
    With Me.txt_surplus_broker_license_no
        .field_top = Me.lbl_surplus_broker_person.field_bottom
    End With
    With Me.lbl_us_producer_no
        .field_top = Me.lbl_surplus_broker_license_no.field_bottom
    End With
    With Me.txt_us_producer_no
        .field_top = Me.lbl_us_producer_no.field_top
    End With
    With Me.lbl_is_underwritten_id
        .field_top = Me.lbl_us_producer_no.field_bottom
    End With
    With Me.txt_is_underwritten_id
        .field_top = Me.lbl_surplus_broker_license_no.field_bottom
    End With
    With Me.lbl_comments
        .field_top = Me.lbl_is_underwritten_id.field_bottom + utilities.twips_converter(0.5, "cm")
    End With
    With Me.txt_comments
        .field_top = Me.lbl_comments.field_bottom
    End With
    With Me.lbl_stella_message_board
        .field_top = Me.txt_comments.field_bottom + utilities.twips_converter(0.5, "cm")
    End With
    With Me.txt_stella_message_board
        .field_top = Me.lbl_stella_message_board.field_bottom
    End With
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub

Public Sub set_bg_of_txt_fields_white()
    Const proc_name As String = "cls_form_deal_details.set_bg_of_text_fields_white"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String
    Dim form_field As cls_field
    str_form = load.deal_details.form_name
    
    For Each form_field In Me.col_all_text_fields
        form_field.field_bg_color = load.colors.white
    Next form_field
    
outro:
    If Not form_field Is Nothing Then Set form_field = Nothing
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub paint_individual_fields(ByVal form_field As cls_field)
    Const proc_name As String = "cls_form_deal_details.paint_individual_fields"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    If CurrentProject.AllForms(Me.form_name).IsLoaded = False Then
        GoTo outro
    End If
    
    With Forms(Me.form_name).Controls(form_field.field_name)
        .BackColor = form_field.field_bg_color
        If form_field.field_caption <> -1 Then
            .Caption = form_field.field_caption
        End If
        .ForeColor = form_field.font_color
        .Height = form_field.field_height
        .Left = form_field.field_left
        .Top = form_field.field_top
        .Width = form_field.field_width
    End With

outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub paint_deal_details()
    Const proc_name As String = "cls_form_deal_details.paint_deal_details"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim str_form As String
    Dim str_milestone As String
    str_form = load.deal_details.form_name
    
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    
    For Each form_field In Me.col_headers
        With Forms(str_form).Controls(form_field.field_name)
            str_milestone = form_field.field_name
            .BackColor = form_field.field_bg_color
            .BorderColor = form_field.border_color
            .ForeColor = form_field.font_color
            .FontBold = form_field.font_is_bold
            .FontSize = form_field.font_size
            .LeftMargin = form_field.margin_internal_left
            .TopMargin = form_field.top_margin
            .Top = form_field.field_top
            .Left = form_field.field_left
            .Height = form_field.field_height
            .Width = form_field.field_width
        End With
    Next form_field
    
    For Each form_field In Me.col_cmd_expanders
        With Forms(str_form).Controls(form_field.field_name)
            str_milestone = form_field.field_name
            .ForeColor = form_field.font_color
            .PressedForeColor = form_field.font_color_pressed
            .HoverForeColor = form_field.font_color_hover
            .Left = form_field.field_left
            .Top = form_field.field_top
            .Caption = form_field.field_caption
        End With
    Next form_field
    
    For Each form_field In Me.col_cmd_generals
        With Forms(str_form).Controls(form_field.field_name)
            str_milestone = form_field.field_name
            If form_field.field_type <> "combo_box" Then
                .Caption = form_field.field_caption
                .ForeColor = form_field.font_color
                .PressedForeColor = form_field.font_color_pressed
                .HoverForeColor = form_field.font_color_hover
                .BackColor = form_field.field_bg_color
            End If
            .Left = form_field.field_left
            .Top = form_field.field_top
            .Width = form_field.field_width
            If Me.make_all_fields_visible_for_debugging = False Then
                .Visible = form_field.field_visible
            End If
        End With
    Next form_field
    
    For Each form_field In Me.col_all_labels
        With Forms(str_form).Controls(form_field.field_name)
            str_milestone = form_field.field_name
            .Top = form_field.field_top
            .Left = form_field.field_left
            If Me.make_all_fields_visible_for_debugging = False Then
                .Visible = form_field.field_visible
            End If
            If form_field.field_caption <> "-1" Then
                .Caption = form_field.field_caption
            End If
            .Width = form_field.field_width
        End With
    Next form_field
    
    For Each form_field In Me.col_all_text_fields
        With Forms(str_form).Controls(form_field.field_name)
            str_milestone = form_field.field_name
            .Top = form_field.field_top
            .Height = form_field.field_height
            .Left = form_field.field_left
            If Me.make_all_fields_visible_for_debugging = False Then
                .Visible = form_field.field_visible
            End If
            .Width = form_field.field_width
            If form_field.has_line = True Then
                With Forms(str_form).Controls(form_field.line_name)
                    .Top = form_field.field_height + form_field.field_top
                    .Left = form_field.field_left
                End With
            End If
            .BackColor = form_field.field_bg_color
        End With
    Next form_field
    
    For Each form_field In Me.col_misc_txt_fields
        With Forms(str_form).Controls(form_field.field_name)
            str_milestone = form_field.field_name
            If Me.make_all_fields_visible_for_debugging = False Then
                .Visible = form_field.field_visible
                .value = form_field.field_value
            End If
            If form_field.field_top <> -1 Then
                .Top = form_field.field_top
            End If
            If form_field.field_left <> -1 Then
                .Left = form_field.field_left
            End If
            .Height = form_field.field_height
            .Width = form_field.field_width
            .BackColor = form_field.field_bg_color
            .ForeColor = form_field.font_color
        End With
    Next form_field
    
    For Each form_field In Me.col_special_america_fields
        With Forms(str_form).Controls(form_field.field_name)
            str_milestone = form_field.field_name
            If Me.make_all_fields_visible_for_debugging = False Then
                .Visible = form_field.field_visible
            End If
        End With
    Next form_field

outro:
    If Not form_field Is Nothing Then Set form_field = Nothing
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, str_milestone, "", "", True
    Resume outro
End Sub
Public Sub paint_financial_info_section(ByVal deal_id As Long)
    Const proc_name As String = "cls_form_deal_details.paint_financial_info_section"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String
    Dim form_field As cls_field
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    
    str_form = "deal_details_f"
    
    If CurrentProject.AllForms(str_form).IsLoaded = False Then GoTo outro
    If load.deal_details.is_init = False Then load.deal_details.init deal_id
    
    str_sql = "SELECT deal_id, uw_fee_amount, uw_fee_due_date, uw_fee_received_date" _
        & ", counsel_fee_amount, counsel_fee_due_date, counsel_fee_received_date" _
        & ", signing_invoice_amount, signing_premium_due_date, signing_premium_received_date" _
        & ", closing_premium_due_date, closing_premium_received_date" _
        & ", total_rp_premium_on_deal" _
        & " FROM " & load.sources.risk_details_f_view _
        & " WHERE deal_id = " & deal_id
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    rs.Open
        With Forms(str_form)
            For Each form_field In load.deal_details.col_txt_fields_for_financial_info_section
                If form_field.field_name_in_recordset <> "-1" Then
                    .Controls(form_field.field_name) = rs.Fields(form_field.field_name_in_recordset)
                End If
            Next form_field
                
            !net_uw_fee = !uw_fee - !counsel_fee
            !closing_premium_amount = !total_premium - !signing_premium_amount
        End With
    rs.Close
    
outro:
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then rs.Close
        Set rs = Nothing
    End If
    If Not form_field Is Nothing Then Set form_field = Nothing
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "str_sql = " & str_sql, "", "", True
    Resume outro
End Sub

Public Sub put_data_into_form(ByVal section_collection As Collection, ByVal deal_id As Long)
    Const proc_name As String = "cls_form_deal_details.put_data_into_form"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    Dim form_field As cls_field
    
    str_form = Me.form_name
    'generate sql string
    For Each form_field In section_collection
        str_sql = str_sql & form_field.field_name_in_recordset & ", "
    Next form_field
    str_sql = "SELECT " & str_sql & " deal_id last_field" & " FROM " & load.sources.risk_details_f_view & " WHERE deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        For Each form_field In section_collection
            With Forms(str_form).Controls(form_field.field_name)
                .value = rs.Fields(form_field.field_name_in_recordset).value
            End With
        Next form_field
        .Close
    End With
    
    Set rs = Nothing
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub change_section_visible_status(ByRef section_collection As Collection)
    Const proc_name As String = "cls_form_deal_details.change_section_visible_status"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    Dim new_visible_status As Boolean
    new_visible_status = True
    If section_collection.Count = 0 Then GoTo outro
    If section_collection(1).field_visible = True Then new_visible_status = False
    
    For Each form_field In section_collection
        If new_visible_status = False Then
            form_field.field_visible = new_visible_status
        ElseIf load.system_info.app_continent = load.system_info.continents.americas Then
            If form_field.visible_in_america = True Then
                form_field.field_visible = new_visible_status
            End If
        ElseIf load.system_info.app_continent = load.system_info.continents.eurasia Then
            If form_field.visible_in_europe = True Then
                form_field.field_visible = new_visible_status
            End If
        End If
    Next form_field
outro:
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Function are_there_claims_on_deal(ByVal deal_id As Long) As Boolean
    Const proc_name As String = "cls_form_deal_deatils.are_there_claims_on_deal"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    
    str_sql = "SELECT claim_id FROM " & load.sources.claims_table & " WHERE is_deleted = 0 AND deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        If .RecordCount > 0 Then
            are_there_claims_on_deal = True
        Else
            are_there_claims_on_deal = False
        End If
        .Close
    End With
    Set rs = Nothing
outro:
    Exit Function
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Function
Public Sub print_rr(ByVal deal_id As Long)
    Const proc_name As String = "cls_form_deal_details.print_rr"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim deal_folder As String
    Dim retention_text_virtual As String
    Dim rs As ADODB.Recordset
    Dim str_currency As String
    Dim str_sql As String
    Dim str_target_path As String
    
    If load.system_info.app_continent <> load.system_info.continents.americas Then
        MsgBox "This only works for US deals.", , "Can't generate RR for non-US deals"
        GoTo outro
    End If
    
    'check that deal folder exists
    deal_folder = Paths.folder_path_from_scripts(deal_id, "find")
    If deal_folder = "-1" Then
        MsgBox "I could not find the deal folder, and can therefore not generate the risk review. Maybe it is not created?", , "Deal folder missing"
        GoTo outro
    End If
    
    'get deal name
    str_sql = "SELECT deal_id, deal_name FROM " & load.sources.deals_table & " WHERE deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        .Open
        str_target_path = deal_folder & "\4. Underwriting\" & !deal_name & " - Risk Review.docx"
        .Close
    End With
    
    'check if risk review already exists. Open if yes.
    If Not Dir(str_target_path, vbDirectory) = "" Then
        FollowHyperlink str_target_path
        GoTo outro
    End If
    
    'Fetch data
    str_sql = "SELECT * FROM us_rr_v WHERE deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql): rs.Open
        rs.MoveFirst
        
        'prepare retention text. Most texts are in the recordset, but this has to be calculated.
        str_currency = rs!deal_currency
        If str_currency = "USD" Then
            str_currency = "$"
        End If
        If rs!retention = rs!drop_end Or rs!drop_period = 0 Then
            retention_text_virtual = str_currency & " " & rs!retention & " No drop."
        Else
            retention_text_virtual = str_currency & " " & rs!retention & " dropping to " & str_currency & " " & rs!drop_end_hr & " after " & rs!drop_period & " months."
        End If
        
        'prepare program summary text. Word can replace strings with max 255 chars.
        Dim program_summary_1 As String
        Dim program_summary_2 As String
        Dim program_summary_3 As String
        Dim program_summary_4 As String
        If IsNull(rs!program_summary) = False Then
            program_summary_1 = Left(rs!program_summary, 255)
            program_summary_2 = Mid(rs!program_summary, 256, 255)
            program_summary_3 = Mid(rs!program_summary, 511, 255)
            program_summary_4 = Mid(rs!program_summary, 766, 255)
        End If
        
        Dim input_data As Scripting.Dictionary
        Set input_data = New Scripting.Dictionary
        With input_data
            .Add "deal_name", rs!deal_name.value
            .Add "deal_currency", rs!deal_currency.value
            .Add "primary_uw_hr", rs!primary_uw_hr.value
            .Add "secondary_uw_hr", Nz(rs!secondary_uw_hr.value, "**secondary_uw_hr**")
            .Add "ev_hr", str_currency & " " & rs!ev_hr.value
            .Add "target_business_name", rs!target_business_name.value
            .Add "target_desc", rs!target_desc.value
            .Add "buyer_business_name", rs!buyer_business_name.value
            .Add "total_rp_limit_on_deal_hr", str_currency & " " & rs!total_rp_limit_on_deal_hr.value
            .Add "retention_text_virtual", retention_text_virtual
            .Add "total_rp_premium_on_deal_hr", str_currency & " " & rs!total_rp_premium_on_deal_hr.value
            .Add "rol_hr", Nz(rs!rol_hr.value, "")
            .Add "inception_date", rs!inception_date.value
            .Add "closing_date", Nz(rs!closing_date, "**closing_date**")
            .Add "uw_law_firm_hr", Nz(rs!uw_law_firm_hr.value, "")
            .Add "uw_financial_advisor_hr", Nz(rs!uw_financial_advisor_hr.value, "")
            .Add "primary_or_xs_hr", Nz(rs!primary_or_xs_hr.value, "primary_or_xs_hr")
            .Add "program_summary_1", Replace(program_summary_1, "<br>", vbNewLine)
            .Add "program_summary_2", Replace(program_summary_2, "<br>", vbNewLine)
            .Add "program_summary_3", Replace(program_summary_3, "<br>", vbNewLine)
            .Add "program_summary_4", Replace(program_summary_4, "<br>", vbNewLine)
            .Add "comments", Nz(rs!Comments.value, " ")
            .Add "buyer_law_firm_1", rs!buyer_law_firm_1.value
            .Add "broker_firm", rs!broker_firm.value
            .Add "risk_type", rs!risk_type.value
        End With
    rs.Close
    Set rs = Nothing
        
    'find template file
    Dim item As Variant
    Dim str_template_path As String
    Dim word_document As Word.Document
    Dim WordApp As Word.Application
    
    str_template_path = load.system_info.system_paths.template_path & "rr_us_with_codes.docx"
    
    'Copy template to deal folder
    FileCopy str_template_path, str_target_path

    Set WordApp = New Word.Application
    WordApp.Visible = False
    Set word_document = WordApp.Documents.Open(str_target_path)
    
    'insert data into risk review
    With word_document.Content.Find
        For Each item In input_data
            .Text = "**" & item & "**"
            .Replacement.Text = input_data(item)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        Next item
    End With
    Set input_data = Nothing
    WordApp.Visible = True
    word_document.Save
    WordApp.Activate
    Set word_document = Nothing
    Set WordApp = Nothing
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "deal_id = " & deal_id, "", True
    Resume outro
End Sub

Public Sub sort_all_fields_in_form_for_debugging()
    'to sort the fields, set the form to design mode, and run this module
    
    Dim form_field As cls_field
    load.check_conn_and_variables
    Me.make_all_fields_visible_for_debugging = True
    load.is_debugging = True
    load.deal_details.init 20663
    For Each form_field In Me.col_headers
        With form_field
            .user_has_expanded_section = True
        End With
    Next form_field
    Me.refresh_admin_section
    Me.refresh_submission_notes_section
    Me.refresh_financial_info_section
    Me.refresh_misc_section
    load.deal_details.paint_deal_details
    Me.make_all_fields_visible_for_debugging = False
End Sub

Public Sub add_regions_to_insured_main_region_id()
    load.call_stack = load.call_stack & vbNewLine & "cls_form_deal_details.add_regions_to_insured_main_region_id"
    'add menu items to insured_main_region_id if insured_registered_country_id is USA
    Dim str_sql As String, rs As ADODB.Recordset
    With Forms(load.deal_details.form_name)!insured_main_region_id
        Do While .ListCount > 0
            .RemoveItem (0)
        Loop
        If Forms(load.deal_details.form_name).Controls(load.deal_details.txt_insured_registered_country_id.field_name) = 13 Then '13 = USA
            Set rs = utilities.create_adodb_rs(conn, load.sources.menu_lists.us_states): rs.Open
                Do While rs.EOF = False
                    .AddItem rs!id & ";" & rs!menu_item
                    rs.MoveNext
                Loop
            rs.Close
            Set rs = Nothing
        End If
    End With
    
End Sub
Public Function are_required_field_to_start_cm_populated(ByVal deal_id As Long) As Boolean
    Const proc_name As String = "cls_form_deal_details.are_required_field_to_start_cm_populated"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim col_missing_fields As Collection
    Dim form_field As cls_field
    
    Set col_missing_fields = mandatory_data.check_data(deal_id, load.mandatory_tests.open_cm)
    If col_missing_fields.Count > 0 Then
        For Each form_field In col_missing_fields
            With form_field
                .field_bg_color = load.colors.yellow
            End With
        Next form_field

        are_required_field_to_start_cm_populated = False
    Else
        are_required_field_to_start_cm_populated = True
    End If
    
outro:
    Exit Function
err_handler:
    are_required_field_to_start_cm_populated = "-1"
    Central.err_handler proc_name, Err.Number, Err.Description, str_sql, "deal_id = " & deal_id, "", True
    Resume outro
End Function
Public Function folders_create_folder_us(ByVal obj_deal As cls_deal) As Integer
    Const proc_name As String = "cls_form_deal_details.folders_create_folder_us"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    Dim deal_name As String
    Dim rs As ADODB.Recordset
    Dim str_path As String
    Dim str_sql As String
    
    If obj_deal.is_folder_created = "yes" Then
        str_path = Paths.folder_path_from_scripts(obj_deal.deal_id, "find")
        If str_path = "-1" Then
            str_path = Paths.folder_path_from_scripts(obj_deal.deal_id, "normative")
        End If
    Else
        str_path = Paths.folder_path_from_scripts(obj_deal.deal_id, "normative")
    End If
    
    If Dir(str_path, vbDirectory) = "" Then
        MkDir str_path
        MkDir str_path & "\" & "1. Submission"
        MkDir str_path & "\" & "2. UW Assessment"
        MkDir str_path & "\" & "3. Mandate, Setup"
        MkDir str_path & "\" & "3. Mandate, Setup\NDAs, NRLs"
        MkDir str_path & "\" & "3. Mandate, Setup\Referrals, Approvals"
        MkDir str_path & "\" & "4. Underwriting"
        MkDir str_path & "\" & "4. Underwriting\Buyer DD Reports"
        MkDir str_path & "\" & "4. Underwriting\Purchase Agreements"
        MkDir str_path & "\" & "4. Underwriting\Disclosure Schedules"
        MkDir str_path & "\" & "4. Underwriting\UW Advisor Reports"
        MkDir str_path & "\" & "4. Underwriting\UW Questions"
        MkDir str_path & "\" & "4. Underwriting\UW Questions\Call Agenda & Notes"
        MkDir str_path & "\" & "4. Underwriting\UW Questions\Follow-Ups"
        MkDir str_path & "\" & "4. Underwriting\Emails"
        MkDir str_path & "\" & "5. Binders"
        MkDir str_path & "\" & "5. Binders\Final Binders"
        MkDir str_path & "\" & "5. Binders\Binder Turns"
        MkDir str_path & "\" & "5. Binders\Item Bill"
        MkDir str_path & "\" & "5. Binders\Underlying Binders"
        MkDir str_path & "\" & "6. Policies"
        MkDir str_path & "\" & "6. Policies\NCDs"
        MkDir str_path & "\" & "6. Policies\Policy Turns"
        MkDir str_path & "\" & "6. Policies\Issued Policy"
        MkDir str_path & "\" & "6. Policies\Underlying Policies"
        MkDir str_path & "\" & "7. Miscellaneous"
    Else
        Application.FollowHyperlink str_path
        GoTo outro
    End If
    
    str_sql = "UPDATE deals_t SET is_folder_created = 'yes' WHERE deal_id = " & obj_deal.deal_id
    conn.Execute str_sql
    
    ' Copy files to new structure
    Dim strPathNBI As String
    Dim strPathMisc As String
    Dim str_template_path As String
    
    strPathNBI = str_path & "\2. UW Assessment\"
    strPathMisc = str_path & "\7.1 Miscellaneous\"
    
    deal_name = obj_deal.deal_name
    str_template_path = load.system_info.system_paths.template_path
    On Error Resume Next
        FileCopy str_template_path & "deal_overview.docx", strPathNBI & deal_name & " - overview.docx"
        FileCopy str_template_path & "nbi_template.docx", strPathNBI & deal_name & " - NBI.docx"
        FileCopy str_template_path & "rating_tool.xlsm", strPathNBI & deal_name & " - rating tool.xlsm"
    On Error GoTo 0
    
    'log folder creation
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = sources.deals_table
        .field_name = ""
        .comment = "Folder was created for " & obj_deal.deal_id
        .record_id = obj_deal.deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
    folders_create_folder_us = 1
    
    Application.FollowHyperlink str_path
    
outro:
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then
            rs.Close
        End If
        Set rs = Nothing
    End If
    Exit Function
err_handler:
    folders_create_folder_us = -1
    Central.err_handler proc_name, Err.Number, Err.Description, "", "obj_deal.deal_id = " & obj_deal.deal_id, "", True
    Resume outro
End Function

Public Function folders_create_folder_eurasia(ByVal obj_deal As cls_deal) As Integer
    Const proc_name As String = "cls_form_deal_details.create_folder_eurasia"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    'check if deal folder exists
    Dim folder_feedback As String
    Dim str_path As String
    Dim str_sql As String
    
    If obj_deal.is_folder_created = "yes" Then
        str_path = Paths.folder_path_from_scripts(obj_deal.deal_id, "find")
        If str_path = "-1" Then
            str_path = Paths.folder_path_from_scripts(obj_deal.deal_id, "normative")
        End If
    Else
        str_path = Paths.folder_path_from_scripts(obj_deal.deal_id, "normative")
    End If
    
    ' Create new folder structure
    MkDir str_path
    MkDir str_path & "\" & "Miscellaneous"
    MkDir str_path & "\" & "NBI etc"
    MkDir str_path & "\" & "NDA, Release letters etc"
    MkDir str_path & "\" & "Policy"
    MkDir str_path & "\" & "Policy\executed policy"
    MkDir str_path & "\" & "Policy\ncd"
    MkDir str_path & "\" & "SPA etc"
    MkDir str_path & "\" & "SPA etc\executed spa and schedules"
    MkDir str_path & "\" & "Underwriting"
    MkDir str_path & "\" & "UW Questions"
    MkDir str_path & "\" & "Accounting"
    
    str_sql = "UPDATE deals_t SET is_folder_created = 'yes' WHERE deal_id = " & obj_deal.deal_id
    conn.Execute str_sql
    
    ' Copy files to new structure
    Dim str_path_nbi As String
    Dim path_invoice As String
    Dim path_source As String
    Dim path_target As String
    Dim target_rating_tool_path As String
    Dim source_rating_tool_file As String
    
    str_path_nbi = str_path & "\NBI etc\"
    path_invoice = str_path & "\Accounting\"
         
    'Copy templates to deal folder
    Dim strTemplatePath As String
    strTemplatePath = load.system_info.system_paths.template_path
    On Error Resume Next
        FileCopy Central.get_nbi_template_path(, obj_deal.spa_law), str_path_nbi & obj_deal.deal_name & " - RiskPoint NBI" & ".docx"
        
        source_rating_tool_file = load.system_info.system_paths.rating_tool_template
        target_rating_tool_path = str_path_nbi & obj_deal.deal_name & " - Rating Tool" & ".xlsm"
        FileCopy source_rating_tool_file, target_rating_tool_path
        
        path_source = load.system_info.system_paths.base_path & "Templates\Invoices\Underwriting Fee Invoice.docx"
        path_target = path_invoice & obj_deal.deal_name & " - UW Fee invoice template.docx"
        FileCopy path_source, path_target
        FileCopy strTemplatePath & "Emails\NBI_approval_request.msg", str_path_nbi & obj_deal.deal_name & " - NBI_approval_request.msg"
        FileCopy strTemplatePath & "Empty document templates\Key issues.docx", str_path & "\" & "Underwriting\" & obj_deal.deal_name & " - Key issues.docx"
        
        path_source = strTemplatePath & "sanction_searches\sanction_screening_stella.xlsx"
        path_target = str_path & "\" & "Miscellaneous\" & obj_deal.deal_name & " (" & obj_deal.deal_id & ") " & " - Sanction Screening.xlsx"
        FileCopy path_source, path_target
    On Error GoTo err_handler
    If load.is_debugging = True Then On Error GoTo 0
    
    'put deal_id into rating tool
    Dim appExcel As Excel.Application
    Dim wb As workbook
    Dim ws_input As Worksheet
    
    Set appExcel = New Excel.Application
    appExcel.Visible = False
    Set wb = appExcel.Workbooks.Open(target_rating_tool_path, ReadOnly:=False)
        Set ws_input = wb.Worksheets("input")
            ws_input.Range("c3") = obj_deal.deal_id
            wb.Save
            If load.is_debugging = True Then appExcel.Visible = True
        Set ws_input = Nothing
        wb.Close
        appExcel.Quit
    Set wb = Nothing
    Set appExcel = Nothing
    
    'open deal folder
    Application.FollowHyperlink str_path
    
    'log that folder was created
    Dim log_object As cls_log_object
    Set log_object = New cls_log_object
    With log_object
        .changer_id = Environ("username")
        .data_set = sources.deals_table
        .field_name = ""
        .comment = "Deal folder was created"
        .record_id = obj_deal.deal_id
    End With
    Central.data_logger log_object
    Set log_object = Nothing
    
    'update the message board
    Dim str_message_board As String
    str_message_board = Date & ": Folder was created by " & current_uw.initials & "." & vbNewLine & str_message_board
    conn.Execute "UPDATE " & sources.deals_table & " SET MessageBoard = '" & str_message_board & "' WHERE deal_id = " & obj_deal.deal_id
    
outro:
    Exit Function
err_handler:
    folders_create_folder_eurasia = -1
    Central.err_handler proc_name, Err.Number, Err.Description, "", "obj_deal.deal_id = " & obj_deal.deal_id, "", True
    Resume outro
End Function


