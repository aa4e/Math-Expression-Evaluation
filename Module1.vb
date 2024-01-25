Imports System.Text
Imports System.Text.RegularExpressions

Module Module1

    ''' <summary>
    ''' Промежуточный шаг NUM вычислен с результатом RESULT.
    ''' </summary>
    ''' <param name="num">Номер шага.</param>
    ''' <param name="result">Результат вычисления текущего шага.</param>
    Public Event StepEvaluated(num As Integer, result As String)

    ''' <summary>
    ''' Вычисляет заданное математическое выражение.
    ''' </summary>
    ''' <param name="expression">Математическое выражение. Десятичный разделитель - только точки!</param>
    ''' <remarks>
    ''' Допустимые операторы и функции:<para />
    ''' <list type="bullet">
    '''   <item>+, -, *, /, ^;</item>
    '''   <item>exp(x), ln(x), lg(x), abs(x), sqr(x), sin(x), cos(x), tan(x), asin(x), atan(x), acos(x);</item>
    '''   <item>min(x, y), max(x, y), x mod y.</item>
    ''' </list>
    ''' Синтаксис для указания значения из предыдущей операции: {-1}, где число - смещение относительно текущей операции.<para />
    ''' Также можно использовать константы: {const_name}.
    ''' </remarks>
    ''' <example>
    ''' <list type="bullet">
    '''   <item>sin(45 * {const_pi} / 180)^3 + 5/2 - min(2, 4) + 1.2E1</item>
    '''   <item>min({-1}, {-2}) * 1.5E-2</item>
    '''   <item>sqr({const_pi} / 2)</item>
    '''   <item>5 Mod 2 - остаток от деления 5 на 2, равный т.е. "1"</item>
    ''' </list>
    ''' </example>
    Public Function Evaluate(expression As String) As Double

        Const num As String = "\s*[-]?\d+\.?\d*(E[+-]?\d+)?\b\s*" 'число с плавающей точкой или в экспоненциальном формате и необязательными начальными и конечными пробелами
        Const nump As String = "\s*\((?<nump>" & num & ")\)\s*" 'число, заключённое в круглые скобки

        'Математические операции:
        Const add As String = "(?<![*/^]\s*)(?<add1>" & num & ")\+(?<add2>" & num & ")(?!\s*[*/^])"
        Const subt As String = "(?<![*/^]\s*)(?<sub1>" & num & ")\-(?<sub2>" & num & ")(?!\s*[*/^])"
        Const mul As String = "(?<!\^\s*)(?<mul1>" & num & ")\*(?<mul2>" & num & ")(?!\s*\^)"
        Const div As String = "(?<!\^\s*)(?<div1>" & num & ")\/(?<div2>" & num & ")(?!\s*\^)"
        Const modu As String = "(?<!\^\s*)(?<mod1>" & num & ")\s+mod\s+(?<mod2>" & num & ")(?!\s*\^)"
        Const pow As String = "(?<pow1>" & num & ")\^(?<pow2>" & num & ")"

        'Функции с одним и двумя операндами:
        Const fOne As String = "(?<fone>(exp|log|ln|log10|lg|abs|sqr|sqrt|sin|cos|tan|asin|acos|atan))\s*\((?<fone1>" & num & ")\)"
        Const fTwo As String = "(?<ftwo>(min|max|mod)\s*)\((?<ftwo1>" & num & "),(?<ftwo2>" & num & ")\)"

        'Всё вместе:
        Const pattern As String = "(" & fOne & "|" & fTwo & "|" & modu & "|" & pow & "|" & div & "|" & mul & "|" & subt & "|" & add & "|" & nump & ")"

        'Для операций сложения и вычитания после +/- добавляется пробел, чтобы эти знаки нельзя было перепутать со знаком числа:
        expression = Regex.Replace(expression, "(?<=[0-9)]\s*)[+-](?=[0-9(])", "$0")

        Dim reNumber As New Regex("^" & num & "$")
        Dim reEval As New Regex(pattern, RegexOptions.IgnoreCase)

        'Цикл выполняется, пока из выражения не будет получено число:
        Dim stepNum As Integer = 1
        Do Until reNumber.IsMatch(expression)

            'замена только одного подвыражения, которое может быть обработано:
            Dim newExpr As String = reEval.Replace(expression, AddressOf PerformOperation, 1)

            RaiseEvent StepEvaluated(stepNum, newExpr)
            stepNum += 1

            If (expression = newExpr) Then 'если выражение не было упрощено, может возникнуть проблема
                Throw New ArgumentException($"Недопустимое выражение: [ {expression} ]")
            End If

            expression = newExpr 'повторный вход в цикл с новым выражением
        Loop
        Return ParseTextToDouble(expression)

    End Function

    ''' <summary>
    ''' Заменяет совпадение <paramref name="m"/> и возвращает обновлённое выражение.
    ''' </summary>
    Private Function PerformOperation(m As Match) As String
        Dim result As Double = 0
        If (m.Groups("nump").Length > 0) Then
            result = ParseTextToDouble(m.Groups("nump").Value)

        ElseIf (m.Groups("neg").Length > 0) Then
            Return "+"

        ElseIf (m.Groups("add1").Length > 0) Then
            Dim add1 As Double = ParseTextToDouble(m.Groups("add1").Value)
            Dim add2 As Double = ParseTextToDouble(m.Groups("add2").Value)
            result = add1 + add2

        ElseIf (m.Groups("sub1").Length > 0) Then
            Dim sub1 As Double = ParseTextToDouble(m.Groups("sub1").Value)
            Dim sub2 As Double = ParseTextToDouble(m.Groups("sub2").Value)
            result = sub1 - sub2

        ElseIf (m.Groups("mul1").Length > 0) Then
            Dim mul1 As Double = ParseTextToDouble(m.Groups("mul1").Value)
            Dim mul2 As Double = ParseTextToDouble(m.Groups("mul2").Value)
            result = mul1 * mul2

        ElseIf (m.Groups("mod1").Length > 0) Then
            Dim mod1 As Double = ParseTextToDouble(m.Groups("mod1").Value)
            Dim mod2 As Double = ParseTextToDouble(m.Groups("mod2").Value)
            result = Math.IEEERemainder(mod1, mod2)

        ElseIf (m.Groups("div1").Length > 0) Then
            Dim div1 As Double = ParseTextToDouble(m.Groups("div1").Value)
            Dim div2 As Double = ParseTextToDouble(m.Groups("div2").Value)
            result = div1 / div2

        ElseIf (m.Groups("pow1").Length > 0) Then
            Dim pow1 As Double = ParseTextToDouble(m.Groups("pow1").Value)
            Dim pow2 As Double = ParseTextToDouble(m.Groups("pow2").Value)
            result = pow1 ^ pow2

        ElseIf (m.Groups("fone").Length > 0) Then
            Dim operand As Double = ParseTextToDouble(m.Groups("fone1").Value)
            Select Case m.Groups("fone").Value.ToLower()
                Case "exp" : result = Math.Exp(operand)
                Case "ln", "log" : result = Math.Log(operand)
                Case "lg", "log10" : result = Math.Log10(operand)
                Case "abs" : result = Math.Abs(operand)
                Case "sqr", "sqrt" : result = Math.Sqrt(operand)
                Case "sin" : result = Math.Sin(operand)
                Case "cos" : result = Math.Cos(operand)
                Case "tan" : result = Math.Tan(operand)
                Case "asin" : result = Math.Asin(operand)
                Case "acos" : result = Math.Acos(operand)
                Case "atan" : result = Math.Atan(operand)
            End Select

        ElseIf (m.Groups("ftwo").Length > 0) Then
            Dim operand1 As Double = ParseTextToDouble(m.Groups("ftwo1").Value)
            Dim operand2 As Double = ParseTextToDouble(m.Groups("ftwo2").Value)
            Select Case m.Groups("ftwo").Value.ToLower()
                Case "min" : result = Math.Min(operand1, operand2)
                Case "max" : result = Math.Max(operand1, operand2)
                Case "mod" : result = operand1 Mod operand2
            End Select
        End If

        Return result.ToString(New Globalization.CultureInfo("en-US")) 'важно чтобы была десятичная ".", поэтому CultInfo.

    End Function

    ''' <summary>
    ''' Преобразует текст в число с плавающей точкой вне зависимости, используется "." или ",".
    ''' </summary>
    Friend Function ParseTextToDouble(val As String) As Double
        Dim d As Double = 0
        If (val.Length > 0) Then
            For i As Integer = 0 To 2

                If (Not Double.TryParse(val, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture, d)) Then
                    If (Not Double.TryParse(val, Globalization.NumberStyles.Any, New Globalization.CultureInfo("en-US"), d)) Then
                        If (Not Double.TryParse(val, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d)) Then

                            Dim expDblValue As New RegularExpressions.Regex("(?i)[+-]?\d+\.\d+E[+-]?\d+") 'если это строка, содержащая несколько значений в экспоненциальной форме,
                            If expDblValue.IsMatch(val) Then 'например, "+8.97678375244E-001,+0.00000000000E+000"
                                val = expDblValue.Match(val).Value
                                Continue For
                            End If
                            Debug.WriteLine("format error, return 0")
                        End If
                        Exit For
                    End If
                    Exit For
                End If
                Exit For
            Next
        End If
        Return d
    End Function

End Module
