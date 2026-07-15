Attribute VB_Name = "Form_F_M_項目マスタ_SUB"
Attribute VB_Base = "0{F0780184-44D3-434A-9E52-96BA1B472DC8}"
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
