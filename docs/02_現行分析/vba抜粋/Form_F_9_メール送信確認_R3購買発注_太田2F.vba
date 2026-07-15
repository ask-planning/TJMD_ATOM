Attribute VB_Name = "Form_F_9_メール送信確認_R3購買発注_太田2F"
Attribute VB_Base = "0{CD2C10A9-E6D7-4789-BF4A-B2A1782BC932}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_9_メール送信確認_太田2F発注")

End Sub

Private Sub btn_tojiru_Click()
    
    MsgBox ("送信フォームに戻り、再度【送信ボタン】をクリックして送信してください。")
    
    DoCmd.Close
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    
End Sub
