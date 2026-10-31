# Практика 2
Задание 1
```vb
Sub ex_1()
    Dim a As Double, b As Double, t As Double
    Dim oldA As Double, oldB As Double
    a = CDbl(InputBox("A="))
    b = CDbl(InputBox("B="))
    oldA = a: oldB = b
    If a > b Then
        t = a
        a = b
        b = t
    End If
    MsgBox "Было: A=" & oldA & ", B=" & oldB & vbCrLf & _
           "Стало: A=" & a & ", B=" & b
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
Задание 6
```vb
Sub ex_6()
    Dim f As Long, a As Double, b As Double, c As Double, d As Double
    Dim h As Double, p As Double, s As Double, q As Double
    f = CLng(InputBox("1-треугольник, 2-прямоугольник, 3-квадрат, 4-круг, 5-трапеция, 6-ромб"))

    If f = 1 Then
        a = CDbl(InputBox("Сторона a=")): b = CDbl(InputBox("Сторона b="))
        c = CDbl(InputBox("Сторона c="))
        If a <= 0 Or b <= 0 Or c <= 0 Or a+b<=c Or a+c<=b Or b+c<=a Then
            MsgBox "Неверные стороны треугольника.": Exit Sub
        End If
        p = a + b + c: q = p / 2: s = Sqr(q*(q-a)*(q-b)*(q-c))
    ElseIf f = 2 Then
        a = CDbl(InputBox("Длина=")): b = CDbl(InputBox("Ширина="))
        p = 2*(a+b): s = a*b
    ElseIf f = 3 Then
        a = CDbl(InputBox("Сторона=")): p = 4*a: s = a*a
    ElseIf f = 4 Then
        a = CDbl(InputBox("Радиус=")): p = 2*3.14159265358979*a: s = 3.14159265358979*a*a
    ElseIf f = 5 Then
        a = CDbl(InputBox("Нижнее основание=")): b = CDbl(InputBox("Верхнее основание="))
        c = CDbl(InputBox("Боковая сторона 1=")): d = CDbl(InputBox("Боковая сторона 2="))
        h = CDbl(InputBox("Высота=")): p = a+b+c+d: s = (a+b)*h/2
    ElseIf f = 6 Then
        a = CDbl(InputBox("Сторона ромба=")): h = CDbl(InputBox("Высота="))
        p = 4*a: s = a*h
    Else
        MsgBox "Неверный номер фигуры.": Exit Sub
    End If
    MsgBox "Периметр = " & p & vbCrLf & "Площадь = " & s
End Sub
        
```
Задание 7
```vb
Sub ex_7()
    Dim m As Long, season As String, days As Long
    m = CLng(InputBox("Номер месяца (1–12)="))
    If m < 1 Or m > 12 Then
        MsgBox "Неверный номер месяца.": Exit Sub
    End If
    If m = 12 Or m = 1 Or m = 2 Then
        season = "Зима"
    ElseIf m >= 3 And m <= 5 Then
        season = "Весна"
    ElseIf m >= 6 And m <= 8 Then
        season = "Лето"
    Else
        season = "Осень"
    End If
    If m = 2 Then
        days = 28
    ElseIf m = 4 Or m = 6 Or m = 9 Or m = 11 Then
        days = 30
    Else
        days = 31
    End If
    MsgBox season & ", дней: " & days
End Sub
```
Задание 8
```vb
Sub ex8()

    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim q1 As Integer
    Dim q2 As Integer

    x1 = InputBox("Точка 1. Введите X:")
    y1 = InputBox("Точка 1. Введите Y:")

    x2 = InputBox("Точка 2. Введите X:")
    y2 = InputBox("Точка 2. Введите Y:")

    If x1 > 0 And y1 > 0 Then
        q1 = 1
    ElseIf x1 < 0 And y1 > 0 Then
        q1 = 2
    ElseIf x1 < 0 And y1 < 0 Then
        q1 = 3
    ElseIf x1 > 0 And y1 < 0 Then
        q1 = 4
    End If

    If x2 > 0 And y2 > 0 Then
        q2 = 1
    ElseIf x2 < 0 And y2 > 0 Then
        q2 = 2
    ElseIf x2 < 0 And y2 < 0 Then
        q2 = 3
    ElseIf x2 > 0 And y2 < 0 Then
        q2 = 4
    End If

    MsgBox "Точка 1 находится в " & q1 & "-й четверти" & vbCrLf & _
           "Точка 2 находится в " & q2 & "-й четверти"

End Sub
```
