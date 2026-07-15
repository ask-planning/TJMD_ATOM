Attribute VB_Name = "Form_F_M_経費コード登録_MAIN"
Attribute VB_Base = "0{41E9EADD-DF52-4672-B007-07873BC80A5C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_close_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![return_form_name]) Then
        strFormName = "F_3_輸入諸経費入力_MAIN"
        intMenuNo = 0
        intMenuNoSeq = 0
    Else
        strFormName = Me![return_form_name]
        intMenuNo = 301
        intMenuNoSeq = 0
    End If
    strOpenMode = "OC"
    Call form_open_close("F_M_経費コード登録_MAIN", strFormName, strOpenMode, "NULL")

End Sub

Private Sub F_M_経費コード登録_SUB_Enter()

    DoCmd.GoToRecord , , acFirst
    
End Sub

Private Sub btn_setsuzoku_Click()
    
    '再接続
    Call Login_ReInput("F_M_経費コード登録_MAIN")

End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    
End Sub
