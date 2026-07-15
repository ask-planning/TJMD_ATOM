Attribute VB_Name = "Form_F_M_調達輸入_商品名称入力_MAIN"
Attribute VB_Base = "0{F5C65E1E-117D-4D18-B30D-7432714F7FB5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_close_Click()
    
    If IsNull(Me![return_form_name]) Then
        strFormName = "F_0_START_MENU"
        intMenuNo = 0
        intMenuNoSeq = 0
    Else
        strFormName = Me![return_form_name]
        intMenuNo = 203
        intMenuNoSeq = 0
    End If
    If Len(Me![INVOICE_NO_MAIN]) > 0 Then
        strInvoiceNo = Me![INVOICE_NO_MAIN]
    Else
        strInvoiceNo = "NULL"
    End If
    strOpenMode = "OC"
    
    Call form_open_close("F_M_調達輸入_商品名称入力_MAIN", strFormName, strOpenMode, strInvoiceNo)

End Sub

Private Sub btn_koushin_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    DoCmd.GoToRecord , , acFirst
    
    'マスタテーブル更新
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_chotatsu_syohingun_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    
    Set cmd = Nothing
    
    MsgBox ("更新完了しました。")
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_M_調達輸入_商品名称入力_MAIN")
    intMenuNo = 906
    intMenuNoSeq = 0

End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    intMenuNo = 906
    intMenuNoSeq = 0
    
End Sub

Private Sub pk_shiiresaki_code_AfterUpdate()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![pk_shiiresaki_code]) Then
        MsgBox ("仕入先コードを入力してください。")
        Exit Sub
    End If
    
    If Len(Me![pk_shiiresaki_code]) < 10 Then
        Me![pk_shiiresaki_code] = Right("0000000000" & Me![pk_shiiresaki_code], 10)
    End If

    Set cmd.ActiveConnection = conn
    
    '作業用テンポラリ作成
    With cmd
        .CommandText = "usp_chotatsu_syohingun_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_shiiresaki_code]
        .Execute
    End With
    
    Set cmd = Nothing
    
    Me![F_M_調達輸入_商品名称入力_SUB].Requery
    
    
End Sub
