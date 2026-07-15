Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Atom.Core.Common

<TestClass()>
Public Class NumberHelperTests

    <TestMethod()>
    Public Sub RoundHalfUp_四捨五入()
        Assert.AreEqual(1.24D, NumberHelper.RoundHalfUp(1.235D, 2))
    End Sub

    <TestMethod()>
    Public Sub Truncate_切り捨て()
        Assert.AreEqual(1.23D, NumberHelper.Truncate(1.239D, 2))
    End Sub

End Class
