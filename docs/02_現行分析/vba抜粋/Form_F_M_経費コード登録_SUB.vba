Attribute VB_Name = "Form_F_M_経費コード登録_SUB"
Attribute VB_Base = "0{CE81BED0-90ED-44E8-BF82-1B243C861B65}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub keihi_code_text_AfterUpdate()

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_M_経費コード登録_MAIN]![login_code]
    
End Sub

Private Sub keihi_kbn_AfterUpdate()

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_M_経費コード登録_MAIN]![login_code]
    
End Sub

Private Sub pk_keihi_code_AfterUpdate()

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_M_経費コード登録_MAIN]![login_code]
    
End Sub
