# Практика 2
Задание 1
```vb
Sub ex_1()
    Dim a As Double, b As Double
    a = CDbl(InputBox("A="))
    b = CDbl(InputBox("B="))

    If a > b Then
        MsgBox "A больше B"
    End If
End Sub
```

Задание 2
```vb
    Sub ex_2()
        Dim a As Double, b As Double, c As Double, d As Double
        Dim result As Double, prod As Double
        a = CDbl(InputBox("A=")): b = CDbl(InputBox("B="))
        c = CDbl(InputBox("C=")): d = CDbl(InputBox("D="))
        
        If a * b > c * d Then
            result = (a + b + c + d) / 4
            MsgBox "Среднее арифметическое = " & result
        Else
            prod = a * b * c * d
            If prod < 0 Then
                MsgBox "Геометрическое среднее четырёх действительных чисел не является действительным."
            Else
                result = prod ^ 0.25
                MsgBox "Среднее геометрическое = " & result
            End If
        End If
        
    End Sub
```
Задание 3
```vb
    Sub ex_3()
        Dim a As Double, b As Double, c As Double
        a = CDbl(InputBox("Сторона a="))
        b = CDbl(InputBox("Сторона b="))
        c = CDbl(InputBox("Сторона c="))
        
        If a > 0 And b > 0 And c > 0 And _
           a + b > c And a + c > b And b + c > a Then
            MsgBox "Треугольник построить можно."
        Else
            MsgBox "Треугольник построить нельзя."
        End If
        
    End Sub
```   
Задание 4
```vb
    Sub ex_4()
    
      Dim n As Long, i As Long, mode As Long
      Dim x As Double, mx As Double, s As String
      n = CLng(InputBox("Кол-во чисел для сравнения(от 3 до 5)")
      
      If n < 3 Or n > 5 Then
          MsgBox "Допустимо ввести 3, 4 или 5."
          Exit Sub
      End If
      mode = CLng(InputBox("1 — ввод вручную; 2 — случайные числа"))
      Randomize
      For i = 1 To n
          If mode = 1 Then
              x = CDbl(InputBox("Введите число " & i))
          Else
              x = Int(Rnd * 200) - 100
          End If
          s = s & x & " "
          If i = 1 Or x > mx Then mx = x
      Next i
      MsgBox "Числа: " & s & vbCrLf & "Максимум = " & mx
      
    End Sub
``` 
Задание 5
```vb
    Sub ex_5()
        Dim a As Double, b As Double, c As Double, d1 As Double, d2 As Double, d3 As Double, mn As Double
        a = CDbl(InputBox("A=")): b = CDbl(InputBox("B=")): c = CDbl(InputBox("C="))
        d1 = Abs(a - b): d2 = Abs(a - c): d3 = Abs(b - c)
        mn = d1
        If d2 < mn Then mn = d2
        If d3 < mn Then mn = d3
        MsgBox "Минимальная абсолютная разность = " & mn
    End Sub
```
