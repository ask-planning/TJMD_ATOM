Attribute VB_Name = "Form_F_M_項目マスタ_SUB_追加用"
Attribute VB_Base = "0{29BD4AED-6CF2-48FE-BC3A-2B5CAF9253B1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub contents_num_AfterUpdate()

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_M_項目マスタ_MAIN]![login_code]
    
End Sub

Private Sub contents1_AfterUpdate()

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_M_項目マスタ_MAIN]![login_code]
    
End Sub

Private Sub contents2_AfterUpdate()

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_M_項目マスタ_MAIN]![login_code]
    
End Sub

Private Sub contents3_AfterUpdate()

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_M_項目マスタ_MAIN]![login_code]
    
End Sub

Private Sub contents4_AfterUpdate()

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_M_項目マスタ_MAIN]![login_code]
    
End Sub

Private Sub pk_seq_no_AfterUpdate()

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_M_項目マスタ_MAIN]![login_code]
    
End Sub
