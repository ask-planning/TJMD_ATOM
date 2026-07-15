Attribute VB_Name = "Form_F_M_項目マスタ_MAIN"
Attribute VB_Base = "0{D5C58151-8427-4BFA-8CC1-61CBA537455F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_menu_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    'スタートメニューへ戻る
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    strOpenMode = "OC"
    Call form_open_close("F_M_項目マスタ_MAIN", strFormName, strOpenMode, "NULL")

End Sub

Private Sub btn_setsuzoku_Click()
    
    '再接続
    Call Login_ReInput("F_M_項目マスタ_MAIN")
    intMenuNo = 903
    intMenuNoSeq = 0

End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    intMenuNo = 903
    intMenuNoSeq = 0
    
End Sub

Private Sub pk_koumoku_code_AfterUpdate()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    Dim intHenkoFuka As Integer
    Dim strMidashi1 As String
    Dim strMidashi2 As String
    Dim strMidashi3 As String
    Dim strMidashi4 As String

    Set cmd.ActiveConnection = conn
    
    '更新可かどうか
    With cmd
        .CommandText = "usp_koumoku_koushin_flg_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_koumoku_code]
        .Execute
    End With
    
    If IsNull(cmd.Parameters(2)) Then
        intHenkoFuka = 1
    Else
        intHenkoFuka = cmd.Parameters(2)
    End If
    
    If Me![bumon_code] = "HN12501" Then
        intHenkoFuka = 0
    End If
    
    strMidashi1 = ""
    strMidashi2 = ""
    strMidashi3 = ""
    strMidashi4 = ""
    With cmd
        .CommandText = "usp_koumoku_midashi_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_koumoku_code]
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(2)) Then
        strMidashi1 = cmd.Parameters(2)
    End If
    If Not IsNull(cmd.Parameters(3)) Then
        strMidashi2 = cmd.Parameters(3)
    End If
    If Not IsNull(cmd.Parameters(4)) Then
        strMidashi3 = cmd.Parameters(4)
    End If
    If Not IsNull(cmd.Parameters(5)) Then
        strMidashi4 = cmd.Parameters(5)
    End If
    
    With cmd
        .CommandText = "usp_koumoku_contents_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "KMK"
        .Parameters(2) = Me![koumoku_seq_no]
        .Execute
    End With
    If Not IsNull(cmd.Parameters(6)) Then
        strSubFormName = cmd.Parameters(6)
    Else
        strSubFormName = "F_M_項目マスタ_SUB"
    End If
    
    Set cmd = Nothing
    
    Me![remarks].Requery
    
    Select Case intHenkoFuka
        Case 0
            Me![F_M_項目マスタ_SUB].SourceObject = strSubFormName
        Case 1
            Me![F_M_項目マスタ_SUB].SourceObject = "F_M_項目マスタ_SUB_変更不可"
        Case 2
            Me![F_M_項目マスタ_SUB].SourceObject = "F_M_項目マスタ_SUB_追加用"
        Case Else
            Me![F_M_項目マスタ_SUB].SourceObject = "F_M_項目マスタ_SUB_変更不可"
    End Select
    Me![F_M_項目マスタ_SUB].Requery
    
    If Len(strMidashi1) > 0 Then
        Me![F_M_項目マスタ_SUB]![lbl_contents1].Caption = strMidashi1
    End If
    If Len(strMidashi2) > 0 Then
        Me![F_M_項目マスタ_SUB]![lbl_contents2].Caption = strMidashi2
    End If
    If Len(strMidashi3) > 0 Then
        Me![F_M_項目マスタ_SUB]![lbl_contents3].Caption = strMidashi3
    End If
    On Error Resume Next
    If Len(strMidashi4) > 0 Then
        Me![F_M_項目マスタ_SUB]![lbl_contents4].Caption = strMidashi4
    End If
    On Error GoTo 0

End Sub
