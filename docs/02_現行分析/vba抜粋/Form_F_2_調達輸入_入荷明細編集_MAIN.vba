Attribute VB_Name = "Form_F_2_調達輸入_入荷明細編集_MAIN"
Attribute VB_Base = "0{B4D34B40-6147-4F49-936C-E2FA661FEE08}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub bnt_menu_Click()
   
    strFormName = "F_2_調達輸入_INVOICE_HEADER"
    If IsNull(Me![pk_invoice_no_main]) Then
        strInvoiceNo = Me![pk_invoice_no_main]
    Else
        strInvoiceNo = "NULL"
    End If
    strOpenMode = "C"
    
    Call form_open_close("F_2_調達輸入_入荷明細編集_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_2_調達輸入_入荷明細編集_MAIN")
    intMenuNo = 206
    intMenuNoSeq = 0

End Sub

Private Sub btn_touroku_Click()

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE NOを入力して下さい")
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_chotatsu_nyuka_meisai_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    If IsNull(cmd.Parameters(3)) Or cmd.Parameters(3) = 0 Then
        MsgBox ("データは更新されていません、明細をご確認ください")
    Else
        MsgBox ("更新完了しました。")
    End If
    
    Set cmd = Nothing

End Sub
