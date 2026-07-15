Attribute VB_Name = "Form_F_0_START_MENU"
Attribute VB_Base = "0{2A7C4D41-AB23-4B4B-8D58-62A3953F842E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_calender_Click()

    'カレンダー、日新配送スケジュール処理へ
    strFormName = "F_M_CALENDER"
    intMenuNo = 904
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_chotatsu_header_Click()

    '配送手配書作成（調達分＝上海以外）
    strFormName = "F_2_調達輸入_INVOICE_HEADER"
    intShanghaiChotatsu = 2
    intMenuNo = 203
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_chotatsu_R3torikomi_Click()

    'R3購買発注、在庫転送、出庫伝票データの取込み処理へ（調達分＝上海以外）
    strFormName = "F_1_R3購買発注取込"
    intShanghaiChotatsu = 2
    intMenuNo = 201
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_chotatu_R3hacchu_Click()

    'R3から取り込んだ購買発注の照会処理へ（調達分＝上海以外）
    strFormName = "F_1_R3購買発注_MAIN"
    intShanghaiChotatsu = 2
    intMenuNo = 202
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_fax_cover_sheet_Click()

    'ファックスカバーシート印刷
    strFormName = "F_4_FAXカバーシート"
    intMenuNo = 404
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_hatsuban_Click()

    '社内参照番号（KH、HSZ）の発番処理へ
    strFormName = "F_M_社内参照番号_発番_MAIN"
    intMenuNo = 902
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_help_Click()

    MsgBox ("各ページではこのカップボタンでマニュアルへのリンクが表示されます。")
    
End Sub

Private Sub btn_INVOICE_Click()

    '上海輸入INVOICE明細照会、修正処理へ
    strFormName = "F_1_上海輸入_INVOICE_MAIN"
    'intShanghaiChotatsu = 1
    intMenuNo = 102
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_invoice_header_Click()

    '上海輸入INVOICEのヘッダ情報入力処理へ
    strFormName = "F_1_上海輸入_INVOICE_HEADER"
    'intShanghaiChotatsu = 1
    intMenuNo = 103
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_keihi_Click()

    '輸入諸経費入力処理へ
    strFormName = "F_3_輸入諸経費入力_MAIN"
    intMenuNo = 301
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_keihi_code_Click()

    '輸入諸経費入力処理へ
    strFormName = "F_M_経費コード登録_MAIN"
    intMenuNo = 907
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_keihiritsu_Click()

    '輸入諸経費入力処理へ
    strFormName = "F_3_上海輸入諸経費_運賃率_MAIN"
    intMenuNo = 303
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_koumoku_Click()

    'システム内のフラグなどの設定マスタ照会、入力処理へ
    strFormName = "F_M_項目マスタ_MAIN"
    intMenuNo = 903
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub


Private Sub btn_ohta_2Fhacchu_Click()

    'システム内のフラグなどの設定マスタ照会、入力処理へ
    strFormName = "F_1_太田2F発注_台帳照会_MAIN"
    intMenuNo = 109
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_quit_Click()

    'システムの終了、ACCESSの終了
    DoCmd.Close acForm, "F_0_START_MENU"
    DoCmd.Quit
    
End Sub

Private Sub btn_R3_soshin_Click()

    'R3への在庫転送用テキスト作成処理へ
    strFormName = "F_1_R3在庫転送_MAIN"
    'intShanghaiChotatsu = 1
    intMenuNo = 106
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_R3hacchu_Click()

    'R3から取り込んだ購買発注の照会処理へ
    strFormName = "F_1_R3購買発注_MAIN"
    'intShanghaiChotatsu = 1
    intMenuNo = 105
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_sap_amount_Click()

    'INVOICEとR3購買発注との金額チェック処理へ
    strFormName = "F_1_SAP_AMOUNT_MAIN"
    'intShanghaiChotatsu = 1
    intMenuNo = 107
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_shinchoku_Click()
    
    '輸入業務の処理の状況入力、照会処理へ
    strFormName = "F_4_輸入全般_進行状況確認_MAIN"
    intShanghaiChotatsu = 1
    intMenuNo = 403
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_0_START_MENU")
    intMenuNo = 0
    intMenuNoSeq = 0

End Sub

Private Sub btn_shiiresaki_master_Click()

    '仕入先マスタ
    strFormName = "F_M_仕入先マスタメンテ"
    intShanghaiChotatsu = 2
    intMenuNo = 905
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_shikyu_hosozai_Click()

    '支給包装材の入力、R3発注処理へ
    strFormName = "F_1_支給包装材発注_MAIN"
    intShanghaiChotatsu = 1
    intMenuNo = 108
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_syohin_meisyo_Click()

    '調達輸入関連マスタ（商品群別名称登録）
    strFormName = "F_M_調達輸入_商品名称入力_MAIN"
    intShanghaiChotatsu = 2
    intMenuNo = 906
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_tatekaekin_Click()

    '立替金の送金明細、振替伝票出力処理へ
    strFormName = "F_3_立替金送金明細_MAIN"
    intMenuNo = 302
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_torikomi_Click()

    'INVOICE/PACKINGの取込み処理へ
    strFormName = "F_1_上海輸入_INVOICE取込"
    'intShanghaiChotatsu = 1
    intMenuNo = 101
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_R3torikomi_Click()

    'R3購買発注、在庫転送、出庫伝票データの取込み処理へ
    strFormName = "F_1_R3購買発注取込"
    'intShanghaiChotatsu = 1
    intMenuNo = 104
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_makishin_Click()

    '上海からの巻き芯の入荷入力、配送手配処理へ
    strFormName = "F_4_上海輸入_巻芯入力_MAIN"
    intShanghaiChotatsu = 1
    intMenuNo = 401
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_yunyu_schedule_Click()

    '輸入INVOICEの処理状況一覧
    strFormName = "F_4_スケジュール検索_MAIN"
    intMenuNo = 402
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    intMenuNo = 0
    intMenuNoSeq = 0
    
End Sub


Private Sub next_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If Len(strFormName) <= 0 Then
        Exit Sub
    End If

    '帳票印刷フォームへ
    strCurrentForm = "F_0_START_MENU"
    strOpenMode = "OC"
    Call form_open_close("F_0_START_MENU", strFormName, strOpenMode, "NULL")
    
End Sub
