'    __  __         
'   / /_/ /___      
'  / __/ / __ \     
' / /_/ / /_/ / _ _ 
' \__/_/ .___(_|_|_)
'     /_/ 
' .tlp
' Manuel Porras Ojeda
' manuelporrasojeda@gmail.com
' https://www.linkedin.com/in/manuelporrasojeda/
' https://github.com/manon42bcn
'
' ENG
' For legal purposes, calculate an employee's time of service presents 
' some challenges. Legally, the month is always calendar-based, as are the years, 
' regardless of whether a particular month has 28, 30, or 31 days.
' From February 1st to March 2nd, we must have one month and one day of service,
' just like from December 1st to January 2nd. The function has two versions:
' one that returns a text string with the employee's length of service expressed
' in years, months, and days, and another that can return any of the elements.

' ESP
' Por condiciones legales, el cálculo de la antigüedad de un empleado
' presenta algunos desafíos. A nivel legal el mes es siempre calendario,
' al igual que los años, sin importar que un mes en concreto tenga 28, 30 o 31 días. 
' Del 1ro de febrero al 2 de marzo debemos tener un mes y un día de antigüedad, 
' lo mismo que del 1ro de diciembre al 2 de enero. La función tiene dos versiones: 
' una donde se retorna una cadena de texto con la antigüedad del empleado expresada en años, 
' meses y días, y otra donde se pueden retornar cualquiera de los elementos.

' RETURN String
Public Function LOTTT_antique(start_date As Date, end_date As Date) As String

Dim evl_days As Long
Dim evl_month_cond As Long
Dim evl_month_1 As Long
Dim evl_month_2 As Long
Dim rst_months As Long
Dim rst_years As Long
Dim rst_days As Long

evl_days = Day(end_date) - Day(start_date)

If evl_days < 0 Then
    evl_month_cond = -1
Else
    evl_month_cond = 0
End If

evl_month_1 = DateDiff("m", start_date, end_date)
evl_month_2 = evl_month_1 + evl_month_cond
rst_months = evl_month_2 Mod 12
rst_years = Int(evl_month_2 / 12)

If evl_month_cond = -1 Then
    rst_days = 30 + evl_days
Else
    rst_days = evl_days
End If

LOTTT_antique = "Antigüedad: " & rst_years & " año(s), " & rst_months & " mes(es), " & rst_days & " día(s)."

End Function

' RETURN VALUES
' Y = for years.
' M = for months.
' D = for days.
Public Function LOTTT_antiqueValues(start_date As Date, end_date As Date, YearsMonthsDays as String) As Long

Dim evl_days As Long
Dim evl_month_cond As Long
Dim evl_month_1 As Long
Dim evl_month_2 As Long
Dim rst_months As Long
Dim rst_years As Long
Dim rst_days As Long

evl_days = Day(end_date) - Day(start_date)

If evl_days < 0 Then
    evl_month_cond = -1
Else
    evl_month_cond = 0
End If

evl_month_1 = DateDiff("m", start_date, end_date)
evl_month_2 = evl_month_1 + evl_month_cond
rst_months = evl_month_2 Mod 12
rst_years = Int(evl_month_2 / 12)

If evl_month_cond = -1 Then
    rst_days = 30 + evl_days
Else
    rst_days = evl_days
End If

if YearsMonthsDays = "Y" Then
    LOTTT_antiqueValues = rst_years
ElseIf YearsMonthsDays = "M" Then
    LOTTT_antiqueValues = rst_months
Elseif YearsMonthsDays = "D" Then
    LOTTT_antiqueValues = rst_days
End if

End Function