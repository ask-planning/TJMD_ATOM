Attribute VB_Name = "Form_F_2_調達輸入_出荷伝票登録_MAIN"
Attribute VB_Base = "0{37BA0C84-26F9-451B-ABB8-0E6E5F7D10A9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub bnt_menu_Click()

    Dim strSyukkaDenpyoNo As String

    If Not IsNull(Me![pk_invoice_no_main]) Then
    
        On Error Resume Next
        Set cmd.ActiveConnection = conn
    
        strSyukkaDenpyoNo = ""
        strSql = "SELECT [pk_syukka_denpyo_no] FROM [tbl_t_chotatsu_syukka_denpyo] WHERE [pk_invoice_no_main] = '" & Me![pk_invoice_no_main] & "' "
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
        End With
        Set rs1 = cmd.Execute(intRcnt)
        If intRcnt <> 0 Then
            rs1.MoveFirst
            Do Until rs1.EOF
                If Len(strSyukkaDenpyoNo) > 0 Then
                    strSyukkaDenpyoNo = strSyukkaDenpyoNo & "/"
                End If
                strSyukkaDenpyoNo = strSyukkaDenpyoNo & rs1![pk_syukka_denpyo_no].Value
                rs1.MoveNext
            Loop
        End If
        rs1.Close
        
        Set rs1 = Nothing
        Set cmd = Nothing
        
        Forms![F_2_調達輸入_INVOICE_HEADER]![syukka_denpyo_no] = strSyukkaDenpyoNo
        On Error GoTo 0
    End If
    
    strFormName = "F_2_調達輸入_INVOICE_HEADER"
    If IsNull(Me![pk_invoice_no_main]) Then
        strInvoiceNo = Me![pk_invoice_no_main]
    Else
        strInvoiceNo = "NULL"
    End If
    strOpenMode = "C"
    
    Call form_open_close("F_2_調達輸入_出荷伝票登録_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_2_調達輸入_出荷伝票登録_MAIN")
    intMenuNo = 206
    intMenuNoSeq = 0

End Sub
