Attribute VB_Name = "Form_F_1_SAP_AMOUNT_CHECK_MAIN"
Attribute VB_Base = "0{3B690A88-EB14-441B-B5C4-930CA0818E9D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_menu_Click()

    'メインメニューへ戻る
    strFormName = "F_1_SAP_AMOUNT_CHECK_MAIN"
    intMenuNo = 107
    intMenuNoSeq = 1
    strOpenMode = "C"
    
    Call form_open_close("F_1_SAP_AMOUNT_CHECK_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub


Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_1_SAP_AMOUNT_CHECK_MAIN")
    If intShanghaiChotatsu = 2 Then
        intMenuNo = 205
    Else
        intMenuNo = 107
    End If
    intMenuNoSeq = 0

End Sub


Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    If intShanghaiChotatsu = 2 Then
        intMenuNo = 205
    Else
        intMenuNo = 107
    End If
    intMenuNoSeq = 0
    
    'Call chk_display_kbn_Click
    Me![F_1_SAP_AMOUNT_CHECK_SUB].Requery
    Me.Requery
    
End Sub

Private Sub shanghai_chotatsu_AfterUpdate()

    Select Case Me![shanghai_chotatsu]
        Case 1
            intShanghaiChotatsu = 1
            intMenuNo = 107
        Case 2
            intShanghaiChotatsu = 2
            intMenuNo = 205
        Case 3
            intShanghaiChotatsu = 3
            intMenuNo = 107
        Case Else
            intShanghaiChotatsu = 1
            intMenuNo = 107
    End Select
    
    Me![INVOICE_NO_MAIN] = Null
    
    strFormName = "F_1_SAP_AMOUNT_CHECK_MAIN"
    Call next_form_open

End Sub

Private Sub next_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    Select Case Me![shanghai_chotatsu]
        Case 1
            intShanghaiChotatsu = 1
        Case 2
            intShanghaiChotatsu = 2
        Case 3
            intShanghaiChotatsu = 3
        Case Else
            intShanghaiChotatsu = 1
    End Select
    
    strInvoiceNo = "NULL"
    strOpenMode = "O"
    
    Call form_open_close("F_1_SAP_AMOUNT_CHECK_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub

