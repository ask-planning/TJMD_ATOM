Attribute VB_Name = "Form_F_3_上海輸入諸経費_運賃率_MAIN"
Attribute VB_Base = "0{DA382738-C5A5-4B1E-A9DD-E8840E897B3C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_close_Click()
    
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_create_Click()

    If IsNull(Me![pk_yyyy]) Then
        MsgBox ("対象年を入力してください。")
        Exit Sub
    End If
    
    Dim datTaisyoYmdFrom As Date
    Dim datTaisyoYmdTo As Date
    
    datTaisyoYmdFrom = DateSerial(Me![pk_yyyy] - 1, 10, 1)
    datTaisyoYmdTo = DateSerial(Me![pk_yyyy], 9, 30)
    
    Set cmd.ActiveConnection = conn
    
    '抽出期間テーブル作成
    With cmd
        .CommandText = "usp_shanghai_keihiritsu_taisyo_kikan"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_yyyy]
        .Parameters(2) = datTaisyoYmdFrom
        .Parameters(3) = datTaisyoYmdTo
        .Execute
    End With
    
    '経費の合計テーブル作成
    With cmd
        .CommandText = "usp_shanghai_keihiritsu_create_keihi_gokei"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_yyyy]
        .Execute
    End With
    
    'INVOICEの合計金額、M3合計作成
    With cmd
        .CommandText = "usp_shanghai_keihiritsu_create_good"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 12000
        .Parameters.Refresh
        .Parameters(1) = Me![pk_yyyy]
        .Execute
    End With
    
    '品目グループ別の経理率更新
    With cmd
        .CommandText = "usp_shanghai_keihiritsu_create_keihi_ritsu"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_yyyy]
        .Execute
    End With
    
    Set cmd = Nothing
    
    MsgBox ("作成完了しました。")
    
    Me![F_3_上海輸入諸経費_運賃率_SUB].Requery
        
        
End Sub

Private Sub btn_excel_Click()

    If IsNull(Me![pk_yyyy]) Then
        MsgBox ("対象年を入力してください。")
        Exit Sub
    End If
    
    'EXCELへ
    intMenuNo = 303
    intMenuNoSeq = 1
    If Not IsNull(Me![pk_yyyy]) Then
        strInputCode = Me![pk_yyyy]
    Else
        strInputCode = ""
        
    End If
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_3_上海輸入諸経費_運賃率_MAIN")
    intMenuNo = 303
    intMenuNoSeq = 0
    If intShanghaiChotatsu = 0 Then
        intShanghaiChotatsu = 1
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    intMenuNo = 303
    intMenuNoSeq = 0

    Me![pk_yyyy] = Year(Date)
    
End Sub

Private Sub pk_yyyy_LostFocus()

    Me![F_3_上海輸入諸経費_運賃率_SUB].Requery

End Sub

Private Sub next_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    strInvoiceNo = "NULL"
    strOpenMode = "OC"
    
    Call form_open_close("F_3_上海輸入諸経費_運賃率_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub

