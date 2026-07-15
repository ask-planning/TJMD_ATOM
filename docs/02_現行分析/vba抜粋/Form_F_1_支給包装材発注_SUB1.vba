Attribute VB_Name = "Form_F_1_支給包装材発注_SUB1"
Attribute VB_Base = "0{72E67553-18FB-47BF-AFAA-3ACB303D554D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub hinmoku_code_AfterUpdate()

    '品目コードが入力されたら、プラント、保管場所のデフォルトセット
    If Not IsNull(Me![hinmoku_code]) And IsNull(Me![pk_shiiresaki_code]) Then
        Me![pk_shiiresaki_code] = Forms![F_1_支給包装材発注_MAIN]![pk_shiiresaki_code]
    End If
    
    If Not IsNull(Me![hinmoku_code]) And IsNull(Me![plant]) Then
        Me![plant] = Forms![F_1_支給包装材発注_MAIN]![plant]
    End If

    If Not IsNull(Me![hinmoku_code]) And IsNull(Me![hokan_basyo]) Then
        Me![hokan_basyo] = Forms![F_1_支給包装材発注_MAIN]![hokan_basyo]
    End If
    
    If Not IsNull(Me![hinmoku_code]) And IsNull(Me![BL_DATE]) Then
        Me![BL_DATE] = Forms![F_1_支給包装材発注_MAIN]![BL_DATE]
    End If
    
    If Not IsNull(Me![hinmoku_code]) And IsNull(Me![touroku_ymd]) Then
        Me![touroku_ymd] = Now()
    End If

    If Not IsNull(Me![hinmoku_code]) And IsNull(Me![touroku_tanto]) Then
        Me![touroku_tanto] = Forms![F_1_支給包装材発注_MAIN]![touroku_tanto]
    End If
    
End Sub
