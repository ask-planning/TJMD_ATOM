Attribute VB_Name = "Form_F_1_太田2F発注_台帳照会_MAIN"
Attribute VB_Base = "0{E72C8E92-A701-4D8E-B0D8-ADC7CC827D45}"
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
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub


Private Sub btn_pickup_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    Dim dblYyyyMmFrom As Double
    Dim dblYyyyMmTo As Double
    Dim datKakakuTekiyoKaishiYmd As Date
    Dim datKakakuTekiyoSyuryoYmd As Date
    
    Set cmd.ActiveConnection = conn
    
    If Not IsNull(Me![pk_koubai_hacchu_no]) Then
        strKoubaiHacchuNo = Me![pk_koubai_hacchu_no]
    Else
        strKoubaiHacchuNo = "%%"
    End If
    If Not IsNull(Me![syanai_sansyo_no]) Then
        strSyanaiSansyoNo = Me![syanai_sansyo_no]
    Else
        strSyanaiSansyoNo = "%%"
    End If
    
    datKakakuTekiyoKaishiYmd = #1/1/1900#
    datKakakuTekiyoSyuryoYmd = Date
    
    '上海価格の適用開始日と終了日
    If Not IsNull(Me![kakaku_yyyymm_from]) Then
        With cmd
            .CommandText = "usp_TJM_master_kansan_rate_get"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![kakaku_yyyymm_from]
            .Execute
        End With
        If Not IsNull(cmd.Parameters(2)) Then
            datKakakuTekiyoKaishiYmd = cmd.Parameters(2)
        End If
        If Not IsNull(cmd.Parameters(3)) Then
            datKakakuTekiyoSyuryoYmd = cmd.Parameters(3)
        End If
    
        dblYyyyMmFrom = Me![kakaku_yyyymm_from]
        dblYyyyMmTo = Me![kakaku_yyyymm_to]
    Else
        dblYyyyMmFrom = 0
    End If
    
    If Not IsNull(Me![kakaku_yyyymm_to]) Then
        With cmd
            .CommandText = "usp_TJM_master_kansan_rate_get"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![kakaku_yyyymm_from]
            .Execute
        End With
        If Not IsNull(cmd.Parameters(3)) Then
            datKakakuTekiyoSyuryoYmd = cmd.Parameters(3)
        End If
        dblYyyyMmTo = Me![kakaku_yyyymm_to]
    Else
        dblYyyyMmTo = 0
    End If
    
    With cmd
        .CommandText = "usp_R3_koubai_hacchu_temp_create_ohta2F"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = strKoubaiHacchuNo
        .Parameters(3) = strSyanaiSansyoNo
        .Parameters(4) = datKakakuTekiyoKaishiYmd
        .Parameters(5) = datKakakuTekiyoSyuryoYmd
        .Execute
    End With
    
    If IsNull(cmd.Parameters(6)) Or cmd.Parameters(6) = 0 Then
        MsgBox ("対象データはありませんでした。")
    Else
        MsgBox ("確認してください。")
    End If
    
    Me![F_1_太田2F発注_台帳照会_SUB].Requery
        
    Set cmd = Nothing

End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_1_太田2F発注_台帳照会_MAIN")
    intMenuNo = 109
    intMenuNoSeq = 0

End Sub

Private Sub btn_soshin_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    intMenuNo = 109
    intMenuNoSeq = 1
    strInputCode = ""
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open
    
End Sub

Private Sub btn_torikomi_Click()

    '購買発注取込処理へ
    strFormName = "F_1_R3購買発注取込"
    intMenuNo = 104
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    DoCmd.Maximize
    intMenuNo = 109
    intMenuNoSeq = 1
    
End Sub

Private Sub next_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    intShanghaiChotatsu = 1
    strInvoiceNo = ""
    strOpenMode = "OC"
    
    Call form_open_close("F_1_太田2F発注_台帳照会_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub
