Attribute VB_Name = "Kata"
Option Explicit



Public Function HowMuchILoveYou(ByVal nb_petals As Integer) As String
  HowMuchILoveYou = Array("not at all", "I love you", "a little", "a lot", "passionately", "madly")(nb_petals Mod 6)
End Function

Sub Test()
  
  Debug.Print PascalsTriangle(3)(1) = 3^
  
End Sub

Public Function PascalsTriangle(ByVal n As Integer) As Variant
  If n = 0 Then PascalsTriangle = Array(1^): Exit Function
  If n = 1 Then PascalsTriangle = Array(1^, 1^): Exit Function
  
  ReDim PrevArr(0 To n) As LongLong
  ReDim CurrArr(0 To n) As LongLong
  
  PrevArr(0) = 1
  PrevArr(1) = 1
  
  Dim i, j As Integer
  For i = 2 To n
    For j = 0 To i
      If j = 0 Or j = i Then
        CurrArr(j) = 1
      Else
        CurrArr(j) = PrevArr(j - 1) + PrevArr(j)
      End If
    Next
    PrevArr = CurrArr
  Next
  
  PascalsTriangle = CurrArr
  
End Function

Public Function EasyLine(ByVal n As Integer) As LongLong
    Dim sum As LongLong
    Dim arr As Variant
    
    arr = PascalsTriangle(n)
    
    Dim item
    For Each item In arr
      sum = sum + item * item
    Next

    EasyLine = sum
End Function


Public Function Add(x As Integer, y As Integer) As Integer
  Add = x + y
End Function

Public Function Multiply(ByVal a As Integer, ByVal b As Integer) As Integer
  Multiply = a * b
End Function

Public Function Century(ByVal year As Integer) As Integer
    Century = (year - 1) \ 100 + 1
End Function
