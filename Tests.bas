Attribute VB_Name = "Tests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestMethod("HowMuchILoveYou")
Private Sub HowMuchILoveYou()
    Assert.AreEqual Kata.HowMuchILoveYou(7), "I love you"
    Assert.AreEqual Kata.HowMuchILoveYou(3), "a lot"
    Assert.AreEqual Kata.HowMuchILoveYou(6), "not at all"
End Sub


'@TestMethod("PascalsTriangle")
Private Sub PascalsTriangle()
    Assert.SequenceEquals Array(1^), Kata.PascalsTriangle(0)
    Assert.SequenceEquals Array(1^, 1^), Kata.PascalsTriangle(1)
    Assert.SequenceEquals Array(1^, 2^, 1^), Kata.PascalsTriangle(2)
    Assert.SequenceEquals Array(1^, 3^, 3^, 1^), Kata.PascalsTriangle(3)
    Assert.SequenceEquals Array(1^, 4^, 6^, 4^, 1^), Kata.PascalsTriangle(4)
    Assert.SequenceEquals Array(1^, 5^, 10^, 10^, 5^, 1^), Kata.PascalsTriangle(5)
End Sub

'@TestMethod("EasyLineTests")
Private Sub EasyLineTests()
    Assert.AreEqual 3432^, Kata.EasyLine(7)
    Assert.AreEqual 10400600^, Kata.EasyLine(13)
    Assert.AreEqual 2333606220^, Kata.EasyLine(17)
    Assert.AreEqual 35345263800^, Kata.EasyLine(19)
End Sub


'@TestMethod("MultiplyTests")
Private Sub MultiplyTests()
    Assert.AreEqual 1, Kata.Multiply(1, 1), "Given: 1, 1"
    Assert.AreEqual 15, Kata.Multiply(3, 5), "Given: 3, 5"
    Assert.AreEqual 49, Kata.Multiply(7, 7), "Given: 7, 7"
    Assert.AreEqual 121, Kata.Multiply(11, 11), "Given: 11, 11"
End Sub


'@TestMethod("CenturyFromYear")
Private Sub CenturyFromYearTest()
    Assert.AreEqual 18, Kata.Century(1705)
    Assert.AreEqual 19, Kata.Century(1900)
    Assert.AreEqual 17, Kata.Century(1601)
    Assert.AreEqual 20, Kata.Century(2000)
    Assert.AreEqual 1, Kata.Century(89)
End Sub


'@TestMethod("AddTest")
Private Sub AddTest()
    On Error GoTo TestFail
    
    Assert.AreEqual Kata.Add(5, 6), 11
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


