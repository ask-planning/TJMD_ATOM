Attribute VB_Name = "Form_F_9_メール送信確認_R3購買発注_出荷伝票番号"
Attribute VB_Base = "0{C9D04835-6351-45FF-8249-40631DB7D236}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_setsuzoku_LostFocus()

    Call Login_ReInput("F_9_メール送信確認_R3購買発注_出荷伝票番号")

End Sub

Private Sub btn_tojiru_Click()
    
    MsgBox ("送信フォームに戻り、再度【送信ボタン】をクリックして送信してください。")
    
    DoCmd.Close
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    
End Sub
