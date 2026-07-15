Attribute VB_Name = "Form_F_9_メール送信確認_経費入力_保険用"
Attribute VB_Base = "0{858D5CF0-685A-4D08-AC7E-92D54EE0378E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_9_メール送信確認_経費入力_保険用")

End Sub

Private Sub btn_tojiru_Click()
    
    MsgBox ("送信フォームに戻り、再度【送信ボタン】をクリックして送信してください。")
    
    DoCmd.Close
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    
End Sub
