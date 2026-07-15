Attribute VB_Name = "Form_F_9_メール送信確認_包装材_明細"
Attribute VB_Base = "0{52231719-03A1-4FAA-B45B-9C67CC6088EC}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_9_メール送信確認_包装材_明細")

End Sub

Private Sub btn_tojiru_Click()
    
    MsgBox ("送信フォームに戻り、再度【送信ボタン】をクリックして送信してください。")
    
    DoCmd.Close
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    
End Sub
