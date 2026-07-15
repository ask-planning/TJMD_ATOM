Attribute VB_Name = "Form_F_1_取込内容確認_MAIN"
Attribute VB_Base = "0{6A18BAEB-D02A-4712-83D4-C1760574B8AB}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_excel_Click()

    If intShanghaiChotatsu = 2 Then
        intMenuNo = 201
    Else
        intMenuNo = 104
    End If
    strInputCode = ""
    intMenuNoSeq = 1
    
    strFormName = "F_9_メール送信_MAIN"
    
    strOpenMode = "CO"
    Call form_open_close("F_1_取込内容確認_MAIN", strFormName, strOpenMode, "null")

End Sub

Private Sub btn_packing_Click()

    If Me![btn_packing].Caption = "PACKING" Then
        Me![F_1_取込内容確認_SUB].SourceObject = "F_1_取込内容確認_SUB_PACKING"
        Me![btn_packing].Caption = "INVOICE"
    Else
        Me![F_1_取込内容確認_SUB].SourceObject = "F_1_取込内容確認_SUB_INVOICE"
        Me![btn_packing].Caption = "PACKING"
    End If
        
    Me![F_1_取込内容確認_SUB].Requery
    
End Sub

Private Sub btn_setsuzoku_Click()
    
    '再接続
    Call Login_ReInput("F_1_取込内容確認_MAIN")

End Sub


Private Sub btn_tojiru_Click()

    Dim strRefNo As String
    Dim intCurrencyType As Integer
    
    'フォームを閉じる
    If IsNull(Me![return_form_name]) Then
        strFormName = "F_1_上海輸入_INVOICE取込"
    Else
        If Me![return_form_name] = "F_9_メール送信_MAIN" Then
            strFormName = "F_1_R3購買発注取込"
        Else
            strFormName = Me![return_form_name]
        End If
    End If
    strOpenMode = "OC"
    If IsNull(Me![pk_invoice_no_main]) Then
        strInvoiceNo = "NULL"
    Else
        strInvoiceNo = Me![pk_invoice_no_main]
    End If
    
    strRefNo = ""
    strTextName = ""
    intCurrencyType = 0
    If Not IsNull(Me![REF_NO]) Then
        strRefNo = Me![REF_NO]
    End If
    If Not IsNull(Me![text_name]) Then
        strTextName = Me![text_name]
    End If
    If Not IsNull(Me![currency_type]) Then
        intCurrencyType = Me![currency_type]
    End If
    
    Call form_open_close("F_1_取込内容確認_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
    If strFormName = "F_1_上海輸入_INVOICE取込" Then
        Forms(strFormName)![REF_NO] = strRefNo
        Forms(strFormName)![text_name] = strTextName
        Forms(strFormName)![currency_type] = intCurrencyType
    End If
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    '最大化
    DoCmd.Maximize
    If strSubFormName = "F_1_取込内容確認_SUB_INVOICE" Or strSubFormName = "F_1_取込内容確認_SUB_PACKING" Then
        Me![btn_packing].Visible = True
        Me![box_packing].Visible = True
    Else
        Me![btn_packing].Visible = False
        Me![box_packing].Visible = False
    End If
    
End Sub
