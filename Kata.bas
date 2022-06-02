Attribute VB_Name = "Kata"
Option Explicit

Public Function CountBy(x As Integer, n As Integer)
  ReDim arr(0 To n - 1) As Variant
  
  Dim i As Integer
  For i = 0 To n - 1
    arr(i) = (i + 1) * x
  Next
  
  CountBy = arr
End Function

Public Function BmiAsString(bmi As Double) As String
  Select Case bmi
    Case Is <= 18.5
      BmiAsString = "Underweight"
    Case 18.5 To 25
      BmiAsString = "Normal"
    Case 25 To 30
      BmiAsString = "Overweight"
    Case Is >= 30
      BmiAsString = "Obese"
  End Select
End Function

Public Function bmi(weight As Double, height As Double) As Double
  If height = 0 Then Err.Raise Number:=11, Description:="BMI Calculation tried to divide by zero"
  If weight < 0 Or height < 0 Then Err.Raise Number:=1004, Description:="During BMI calculation, either height or weight was negative, which doesn't make sense"
  bmi = weight / height / height
  If bmi > 40 Then MsgBox "BMI Calculation resulted in BMI > 40, which is morbidly obese.  Does this make sense?", vbExclamation
  If bmi < 15 Then MsgBox "BMI Calculation resulted in BMI < 15, which is morbidly obese.  Does this make sense?", vbExclamation
  
End Function

Public Sub Test()
  bmi 14, 1
End Sub

Public Function HowMuchILoveYou(ByVal nb_petals As Integer) As String
  HowMuchILoveYou = Array("not at all", "I love you", "a little", "a lot", "passionately", "madly")(nb_petals Mod 6)
End Function

Public Function StringToNumber(s As String) As Integer
  StringToNumber = CInt(s)
End Function

Public Function AreYouPlayingBanjo(name As String) As String
  Dim regex As New regExp
  regex.Pattern = "^r"
  regex.IgnoreCase = True
  
  If regex.Test(name) Then AreYouPlayingBanjo = name + " plays banjo": Exit Function
  AreYouPlayingBanjo = name + " does not play banjo"
End Function

'Public Function AreYouPlayingBanjo(name As String) As String
'  If Left(name, 1) = "R" Or Left(name, 1) = "r" Then AreYouPlayingBanjo = name + " plays banjo": Exit Function
'  AreYouPlayingBanjo = name + " does not play banjo"
'End Function

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
