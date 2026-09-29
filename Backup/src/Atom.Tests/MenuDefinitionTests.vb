Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Atom.Web.Common

''' <summary>
''' メニュー定義の単体テストです。
''' </summary>
<TestClass()>
Public Class MenuDefinitionTests

#Region "テストメソッド"
    ''' <summary>
    ''' メニューが 4 ブロックで構成されていることを確認します。
    ''' </summary>
    <TestMethod()>
    Public Sub CreateBlocks_ブロック数が4件である()
        Dim blocks As IList(Of MenuBlock) = MenuDefinition.CreateBlocks()
        Assert.AreEqual(4, blocks.Count)
    End Sub

    ''' <summary>
    ''' メニュー項目の合計が 17 件であることを確認します。
    ''' </summary>
    <TestMethod()>
    Public Sub CreateBlocks_メニュー項目の合計が17件である()
        Dim blocks As IList(Of MenuBlock) = MenuDefinition.CreateBlocks()

        Dim total As Integer = 0
        For Each block As MenuBlock In blocks
            total += block.Entries.Count
        Next

        Assert.AreEqual(17, total)
    End Sub

    ''' <summary>
    ''' メニュー番号が重複していないことを確認します。
    ''' </summary>
    <TestMethod()>
    Public Sub CreateBlocks_メニュー番号が重複しない()
        Dim blocks As IList(Of MenuBlock) = MenuDefinition.CreateBlocks()
        Dim menuNos As New HashSet(Of Integer)()

        For Each block As MenuBlock In blocks
            For Each entry As MenuEntry In block.Entries
                Assert.IsTrue(menuNos.Add(entry.MenuNo),
                              "メニュー番号 " & entry.MenuNo.ToString() & " が重複しています。")
            Next
        Next
    End Sub

    ''' <summary>
    ''' 実装済みの画面には URL が登録されていることを確認します。
    ''' </summary>
    <TestMethod()>
    Public Sub GetUrl_実装済み画面はURLを返す()
        Assert.IsTrue(PageUrlMap.IsImplemented(907))
        Assert.AreNotEqual(String.Empty, PageUrlMap.GetUrl(907))
    End Sub

    ''' <summary>
    ''' 未登録のメニュー番号は空文字を返すことを確認します。
    ''' </summary>
    <TestMethod()>
    Public Sub GetUrl_未実装画面は空文字を返す()
        Assert.IsFalse(PageUrlMap.IsImplemented(999))
        Assert.AreEqual(String.Empty, PageUrlMap.GetUrl(999))
    End Sub
#End Region

End Class
