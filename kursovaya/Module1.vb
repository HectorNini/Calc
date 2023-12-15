Module Module1

    Sub Main()
        Dim primer As String
        Dim ext As Boolean = False
        Menu(ext)
        Do While Not ext
            primer = vvod()
            Dim mas() As String = Massiv(primer)
            mas = Merge(mas)
            Scob(mas)
            Resh(mas)
            Console.Write("Ответ: ")
            Print(mas)
            ext = Qou(ext)
        Loop
        Console.ForegroundColor = ConsoleColor.Blue
        Console.WriteLine("Нажмите любую кнопку, чтобы выйти")
        Console.ResetColor()


        Console.ReadKey()
    End Sub
    Function Qou(e As Boolean) As Boolean
        Dim ans As String
        Console.WriteLine("1) Продолжить")
        Console.WriteLine("2) Выход")
        Do
            ans = Console.ReadLine()
            If ans <> "2" And ans <> "1" Then
                Console.ForegroundColor = ConsoleColor.Red
                Console.Write("Ошибка. ")
                Console.ResetColor()
                Console.WriteLine("Проверьте корректность написания.")
            End If
        Loop While ans <> "2" And ans <> "1"
        If ans = "2" Then
            e = True
        ElseIf ans = "1" Then
            Console.Clear()
        End If

        Return e
    End Function
    Sub Pause(k As Integer)
        Dim t As Single
        t = Timer
        Do
        Loop While Timer - t < 0.35 + k
    End Sub
    Sub Menu(ByRef e As Boolean)
        Dim kk As Boolean = False
        Dim ans As String
        Dim zastavka() As String =
            {"                                                ",
             "        RRRRRRRRRRRRRR    RRRRRRRRRRRRRR        ",
             "     RRRR       RRRR        RRRR       RRRR     ",
             "     RRR        RRRR        RRRR        RRR     ",
             "     RRR        RRRR        RRRR        RRR     ",
             "     RRRR       RRRR        RRRR       RRRR     ",
             "      RRR       RRRR        RRRR       RRR      ",
             "        RRRRRRRRRRRR        RRRRRRRRRRRR        ",
             "         RRRR   RRRR        RRRR   RRRR         ",
             "        RRRR    RRRR        RRRR    RRRR        ",
             "      RRRRR     RRRR        RRRR     RRRRR      ",
             "     RRRR       RRRR        RRRR       RRRR     ",
             "   RRRRR        RRRR        RRRR        RRRRR   ",
             "RRRRRRR      RRRRRRRRRR  RRRRRRRRRR      RRRRRRR",
             "                                                "}

        For i = 0 To UBound(zastavka)
            Console.SetCursorPosition(25, 3 + i)
            Console.WriteLine(zastavka(i))
            pause(0)
        Next
        Console.SetCursorPosition(40, 3 + UBound(zastavka) + 1)
        Console.WriteLine("By Ruslan Ryabinin")
        pause(1)
        Do
            Do
                Console.Clear()
                Console.SetCursorPosition(47, 3 + 1)
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.WriteLine("Меню")
                Console.ResetColor()
                Console.SetCursorPosition(47, 3 + 2)
                Console.WriteLine("1) Калькулятор")
                Console.SetCursorPosition(47, 3 + 3)
                Console.WriteLine("2) Помощь")
                Console.SetCursorPosition(47, 3 + 4)
                Console.WriteLine("3) Выход")
                Console.SetCursorPosition(47, 3 + 5)
                Console.Write("Выберите номер действия: ")
                ans = Console.ReadLine()
            Loop While (ans <> "1") And (ans <> "2") And (ans <> "3")

            Select Case ans
                Case "1"
                    Console.Clear()
                    kk = False
                Case "2"
                    Console.Clear()
                    Console.ForegroundColor = ConsoleColor.Yellow
                    Console.WriteLine("Используемые символы:")
                    Console.ResetColor()
                    Console.WriteLine("Знак сложения [+]                                               Открывающая скобка [(]")
                    Console.WriteLine("Знак вычитания [-]                                              Закрывающая скобка [)]")
                    Console.WriteLine("Знак умножения [*]                                              Деление без остатка [\]")
                    Console.WriteLine("Знак деления [/]                                                Знак степени [^]")
                    Console.ForegroundColor = ConsoleColor.Yellow
                    Console.WriteLine("Правила пользования:")
                    Console.ResetColor()
                    Console.WriteLine("1) Правильная запись: 3*(2+1)                                   Неправильная запись: 3(2+1)")
                    Console.WriteLine("2) Правильная запись: (-8)*12                                   Неправильная запись: -8*12")
                    Console.WriteLine("3) Правильная запись: 3/(-3)                                    Неправильная запись: 3/-3")
                    Console.WriteLine("4) Правильная запись: 3^2                                       Неправильная запись: 3^2=")
                    Console.WriteLine("5) Правильная запись: 3,14                                      Неправильная запись: 3.14")
                    Console.ForegroundColor = ConsoleColor.Blue
                    Console.WriteLine("Нажмите любую кнопку")
                    Console.ResetColor()
                    Console.ReadKey()
                    kk = True
                Case "3"
                    Console.SetCursorPosition(47, 3 + 6)
                    kk = False
                    e = True
            End Select
        Loop While kk


    End Sub

    Sub Print(ByRef m() As String)
        For i = 0 To UBound(m)
            Console.Write(m(i))
        Next
        Console.WriteLine()
    End Sub


    Function Durak(str As String) As Boolean
        Dim pr1, pr2, pr3, pr4, pr5 As Boolean
        Dim alph() As String = {"+", "-", "*", "/", "\", "^", ","}
        Dim sc() As String = {"(", ")"}
        Dim mm(Len(str) - 1) As String
        For i = 0 To UBound(mm)
            mm(i) = str(i)
        Next

        pr1 = True
        For i = 0 To UBound(mm)
            If i = 0 Or i = UBound(mm) Then
                If alph.Contains(mm(i)) Then
                    pr1 = False
                End If
            End If
            If Not IsNumeric(mm(i)) And Not alph.Contains(mm(i)) And Not sc.Contains(mm(i)) Then
                pr1 = False
            End If
        Next

        If pr1 Then
            Dim n As Integer = 0
            Dim k As Integer = 0
            pr2 = True
            Do
                If mm(k) = "(" Then
                    n += 1
                ElseIf mm(k) = ")" Then
                    n -= 1
                End If
                k += 1
            Loop While k <= UBound(mm) And n >= 0
            If n <> 0 Then
                pr2 = False
            End If
        End If

        If pr2 Then
            pr3 = True
            For i = 0 To UBound(mm) - 1
                If alph.Contains(mm(i)) And alph.Contains(mm(i + 1)) Then
                    pr3 = False
                ElseIf (IsNumeric(mm(i)) And mm(i + 1) = "(") Or (IsNumeric(mm(i + 1)) And mm(i) = ")") Then
                    pr3 = False
                End If
            Next
        End If

        If pr3 Then
            pr4 = True
            If UBound(mm) > 0 Then
                If (alph.Contains(mm(LBound(mm))) And Not IsNumeric(mm(LBound(mm) + 1))) Or alph.Contains(mm(UBound(mm))) Then
                    pr4 = False
                End If
            ElseIf UBound(mm) = 0 And (alph.Contains(mm(0)) Or sc.Contains(mm(0))) Then
                pr4 = False
            End If
        End If

        Dim zap As Boolean = False

        If pr4 Then
            pr5 = True
            For i = 0 To UBound(mm) - 1
                If mm(i) = "," And zap Then
                    pr5 = False
                End If
                If mm(i) = "," Then
                    zap = True
                End If
                If zap And alph.Contains(mm(i)) And mm(i) <> "," Then
                    zap = False
                End If


                If i <> 0 Then
                    If (mm(i) = "0" And (Not IsNumeric(mm(i - 1)) And mm(i - 1) <> "," And IsNumeric(mm(i + 1)))) Then
                        pr5 = False
                    End If
                ElseIf i = 0 Then
                    If (mm(i) = "0" And IsNumeric(mm(i + 1))) Then
                        pr5 = False
                    End If
                End If
            Next
        End If



        Dim itog As Boolean = pr1 And pr2 And pr3 And pr4 And pr5
        Return itog
    End Function

    Function Vvod() As String
        Dim p As String
        Dim ist As Boolean

        Do
            Console.WriteLine("Введите пример")
            p = Console.ReadLine()
            p = Replace(p, " ", "")
            If p <> "" Then
                ist = Durak(p)
            Else
                ist = False
            End If
            If Not ist Then
                Console.Clear()
                Console.ForegroundColor = ConsoleColor.Red
                Console.Write("Ошибка. ")
                Console.ResetColor()
                Console.WriteLine("Проверьте корректность написания.")
            End If
        Loop While Not ist
        Return p
    End Function

    Sub Spc(ByRef m() As String)
        Dim k As Integer = 0
        Dim kk As Boolean
        Do
            If m(UBound(m)) = " " Then
                ReDim Preserve m(UBound(m) - 1)
            End If
        Loop While m(UBound(m)) = " "
        Do
            kk = False
            For i = 0 To UBound(m) - 1
                If m(i) = " " Then
                    kk = True
                    For j = i To UBound(m) - 1
                        m(j) = m(j + 1)
                    Next
                    k += 1
                End If
            Next
        Loop While kk
        ReDim Preserve m(UBound(m) - k)

    End Sub
    Function Massiv(p As String) As String()
        Dim m(Len(p) - 1) As String
        For i = 0 To UBound(m)
            m(i) = CStr(p(i))
        Next

        Return m
    End Function

    Function Merge(ByRef m() As String) As String()
        Dim kat1, kat2, kat3, kk As Boolean
        Do
            kk = False

            For i = 0 To UBound(m) - 1
                kat1 = IsNumeric(m(i))
                kat2 = IsNumeric(m(i + 1))
                kat3 = False
                If m(i + 1) = "," Then
                    kat3 = True
                End If
                If (kat1 And kat2) Or (kat1 And kat3) Then
                    m(i) = m(i) & m(i + 1)
                    m(i + 1) = " "
                    kk = True
                End If
            Next
            Spc(m)
        Loop While kk
        Spc(m)

        Return m
    End Function
    Sub Scob(ByRef ms() As String)
        Dim sc1, sc2 As Integer
        Dim kk As Boolean
        Dim last_m As Integer
        Dim kol As Integer = 1

        Do
            last_m = UBound(ms)
            kk = False
            For i = 0 To last_m

                If ms(i) = "(" Then
                    sc1 = i
                    kk = True
                End If
            Next
            For i = sc1 To last_m
                If ms(i) = ")" Then
                    sc2 = i
                    Exit For
                End If
            Next
            If kk Then
                Dim m_dop(sc2 - sc1 - 2) As String
                For i = 0 To UBound(m_dop)
                    m_dop(i) = ms(sc1 + i + 1)

                Next
                ms(sc1) = Resh(m_dop)
                For i = sc1 + 1 To sc2
                    ms(i) = " "
                Next
                Spc(ms)
            End If
            Console.Write(kol & ") ")
            Print(ms)
            kol += 1
        Loop While kk

    End Sub

    Function Resh(ByRef m() As String) As String
        Dim kk As Integer
        Do
            kk = False
            For i = 0 To UBound(m) - 1
                If m(i) = "^" And Not kk Then
                    m(i) = CDbl(m(i - 1)) ^ CDbl(m(i + 1))
                    m(i - 1) = " "
                    m(i + 1) = " "
                    kk = True
                End If
            Next
            Spc(m)
        Loop While kk

        Do
            kk = False
            For i = 0 To UBound(m) - 1
                If (m(i) = "*" Or m(i) = "/" Or m(i) = "\") And Not kk Then
                    If m(i) = "*" Then
                        m(i) = CDbl(m(i - 1)) * CDbl(m(i + 1))
                        m(i - 1) = " "
                        m(i + 1) = " "
                        kk = True
                    ElseIf m(i) = "/" Then
                        m(i) = CDbl(m(i - 1)) / CDbl(m(i + 1))
                        m(i - 1) = " "
                        m(i + 1) = " "
                        kk = True
                    ElseIf m(i) = "\" Then
                        If m(i + 1) = "0" Then
                            m(i) = CDbl(m(i - 1)) / CDbl(m(i + 1))
                            m(i - 1) = " "
                            m(i + 1) = " "
                            kk = True
                        Else
                            m(i) = CDbl(m(i - 1)) \ CDbl(m(i + 1))
                            m(i - 1) = " "
                            m(i + 1) = " "
                            kk = True
                        End If
                    End If
                End If
            Next
            Spc(m)
        Loop While kk

        Do
            kk = False
            For i = 0 To UBound(m) - 1
                If (m(i) = "+" Or m(i) = "-") And Not kk Then
                    If m(i) = "+" Then
                        m(i) = CDbl(m(i - 1)) + CDbl(m(i + 1))
                        m(i - 1) = " "
                        m(i + 1) = " "
                        kk = True
                    ElseIf m(i) = "-" Then
                        If i <> LBound(m) Then
                            m(i) = CDbl(m(i - 1)) - CDbl(m(i + 1))
                            m(i - 1) = " "
                            m(i + 1) = " "
                            kk = True
                        ElseIf i = LBound(m) Then
                            m(i) = CDbl(m(i + 1)) * -1
                            m(i + 1) = " "
                            kk = True
                        End If
                    End If
                End If
            Next
            Spc(m)

        Loop While kk



        Return m(0)
    End Function

End Module
