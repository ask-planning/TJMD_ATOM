Attribute VB_Name = "Form_F_4_上海輸入_巻芯入力_SUB"
Attribute VB_Base = "0{5AF263C2-3010-470A-9412-9045F9ED372C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub makishin_code_AfterUpdate()

    Set cmd.ActiveConnection = conn
    
    '巻き芯のコードから配送先GET
    With cmd
        .CommandText = "usp_koumoku_contents_get_text"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "MKS"
        .Parameters(2) = Me![makishin_code]
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(4)) Then
        Me![haiso_type] = cmd.Parameters(4)
        '配送先＝93は鳥居なので、鳥居フラグをON
        If cmd.Parameters(4) = "93" Then
            Me![torii_flg] = 1
        Else
            Me![torii_flg] = 0
        End If
    Else
        Me![haiso_type] = 91
        Me![torii_flg] = 0
    End If
    
    Set cmd = Nothing
    
End Sub
