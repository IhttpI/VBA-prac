# VBA-prac

Задание 1
```vb
Option Explicit
	Sub ex1()
	    Dim a As Double, V As Double, S As Double
	    a = CDbl(InputBox("Введите длину ребра куба"))
	    V = a ^ 3
	    S = 6 * a ^ 2
	    MsgBox "Объём = " & V & vbCrLf & "Площадь поверхности = " & S    
	End Sub
```
Задание 2
```vb
	Sub ex2()
	    Dim raw As Double, yarn As Double, waste As Double, loss As Double
	    raw = 12000
	    yarn = raw * 0.93
	    waste = raw * 0.06
	    loss = raw * 0.01
	    
	    MsgBox "Пряжа: " & yarn & " кг" & vbCrLf & _
	           "Отходы: " & waste & " кг" & vbCrLf & _
	           "Потери: " & loss & " кг"
	           
	End Sub
```
Задание 3
```vb
	Sub ex3()
	    Dim L As Double, W As Double, r As Double
	    Dim n As Long, fabric As Double, circles As Double, waste As Double
	    
	    L = 12
	    W = 1.4
	    r = 0.15
	    
	    n = Int(L / (2 * r))
	    fabric = L * W
	    circles = n * 3.14159265358979 * r ^ 2
	    waste = fabric - circles
	    
	    MsgBox "Количество заготовок = " & n & vbCrLf & _
	           "Площадь отходов = " & waste & " м^2"
	           
	End Sub
```	
Задание 4
```vb
	Sub ex4()
	    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, d As Double
	    
	    x1 = CDbl(InputBox("Введите X1:"))
	    y1 = CDbl(InputBox("Введите Y1:"))
	    x2 = CDbl(InputBox("Введите X2:"))
	    y2 = CDbl(InputBox("Введите Y2:"))
	    
	    d = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
	    
	    MsgBox "Расстояние = " & d
	
	End Sub
```
Задание 5
```vb
	Sub ex5()
	    Dim a As Double, b As Double, c As Double, t As Double
	    
	    a = CDbl(InputBox("A="))
	    b = CDbl(InputBox("B="))
	    c = CDbl(InputBox("C="))
	    
	    t = a + b - c
	    
	    If t = 0 Then
	        MsgBox "A+B-C равно нулю; проценты не определены."
	    Else
	        MsgBox "Доля A = " & 100 * a / t & "%" & vbCrLf & _
	               "Доля B = " & 100 * b / t & "%" & vbCrLf & _
	               "Доля C = " & 100 * c / t & "%"
	    End If
	    
	End Sub
```
Задание 6
```vb
	Sub ex6()
	    Dim productionPerMinute As Long, machines As Long, hours As Long, minutes As Long, total As Long
	
	    productionPerMinute = 7
	    machines = 3
	    hours = 6
	
	    minutes = hours * 60
	    total = productionPerMinute * machines * minutes
	    
	    MsgBox "Будет выпущено " & total & " тарелок."
```
	
	    
	End Sub

