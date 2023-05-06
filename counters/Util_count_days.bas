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

' ENG
' Functions to count days. The function utl_count_days counts the number of days
' (all, Saturdays, Sundays, or Mondays *) between two dates. 'Mondays' are a useful
' option in the Venezuelan legal framework for calculating social security withholdings.
' The function utl_is_holiday returns whether a date is considered a holiday or not.
' To do this, you must have a list of dates considered holidays in a column of your
' Excel file. The dates in the column must be in 'dd/mm/yyyy' format as a string.
' If it is an annual holiday, you can include it with the year 0000, for example,
' '01/01/0000' will consider all January 1st as holidays.
' utl_counter combines both functions to count "labor_days" to obtain working days,
' "holidays" to obtain only holidays, and "weekdays" to obtain weekdays without
' taking into account weekends.

' ESP
' Funciones para contar días. La función utl_count_days cuenta el número de
' días (todos, sábados, domingos o lunes *) entre dos fechas. Los lunes son
' una opción útil en el marco legal venezolano para el cálculo de retenciones
' de la seguridad social.
' La función utl_is_holiday devuelve si una fecha es considerada feriado o no.
' Para ello debe tener una lista de fechas consideradas feriados en una columna
' de su archivo de Excel. Las fechas de la columna deben estar en formato
' ‘dd/mm/aaaa’ en formato cadena. Si se trata de un feriado anual puede incluirlo
' con el año en 0000, por ejemplo ‘01/01/0000’ considerará todos los primero de
' enero como feriados.
' utl_counter une ambas funciones, para contar “labor_days” para obtener días
' hábiles, “holidays” para obtener sólo los feriados y “weekdays” para obtener
' los días de semana sin tomar en cuenta los fines de semana.

Public Function ult_counter(start_date As Date, end_date As Date, holidays_range as Range, filter_as as String) As Long

Dim counter As Long
Dim holidays As Long
Dim saturdays As Long
Dim sundays As Long

counter = 1
holidays = 0
If end_date < start_date Then
    ult_counter = 0
    Exit Function
End If
Do Until start_date = end_date
    counter = counter + 1
    holidays = holidays + utl_is_holiday(start_date, holidays_range)
    start_date = DateAdd("d", 1, start_date)
Loop
saturdays = ult_count_days(start_date, end_date, 7)
sundays = ult_count_days(start_date, end_date, 1)
If filter_as = "labor_days" Then
    utl_counter = counter - (saturdays + sundays + holidays)
Else If filter_as = "holidays" Then
    ult_counter = holidays
Else If filter_as = "weekdays" Then
    utl_counter = counter - (saturdays + sundays)
End If

End Function

Public Function utl_count_days(start_date As Date, end_date As Date, sat7sun1mon2 As Long, Optional alldays As Long = 0) As Long

Dim fix_start As Date
Dim fix_end As Date
    
    If Weekday(start_date) = sat7sun1mon2 Then
    fix_start = DateAdd("d", -1, start_date)
    Else
    fix_start = start_date
    End If
    
    If Weekday(end_date) = sat7sun1mon2 Then
    fix_end = DateAdd("d", 0, end_date)
    Else
    fix_end = end_date
    End If
    
If alldays = 0 Then
    If sat7sun1mon2 = 1 Then
    utl_count_days = DateDiff("ww", (fix_start), (fix_end), vbSunday)
    ElseIf sat7sun1mon2 = 2 Then
    utl_count_days = DateDiff("ww", (fix_start), (fix_end), vbMonday)
    ElseIf sat7sun1mon2 = 3 Then
    utl_count_days = DateDiff("ww", (fix_start), (fix_end), vbTuesday)
    ElseIf sat7sun1mon2 = 4 Then
    utl_count_days = DateDiff("ww", (fix_start), (fix_end), vbWednesday)
    ElseIf sat7sun1mon2 = 5 Then
    utl_count_days = DateDiff("ww", (fix_start), (fix_end), vbThursday)
    ElseIf sat7sun1mon2 = 6 Then
    utl_count_days = DateDiff("ww", (fix_start), (fix_end), vbFriday)
    ElseIf sat7sun1mon2 = 7 Then
    utl_count_days = DateDiff("ww", (fix_start), (fix_end), vbSaturday)
    End If
Else
    utl_count_days = DateDiff("d", (fix_start), (fix_end))
End If

End Function


Public Function utl_is_holiday(toCheck As Date, holidays As Range) As Long

Dim get_holiday As String
Dim i As Integer
Dim cell_value As String
Dim is_holiday As Long

is_holiday = 0
For i = 1 To holidays.Rows.Count
    If IsEmpty(holidays.Cells(i, 1)) Then
        Exit For
    End If
    cell_value = CStr(holidays.Cells(i, 1).Value)
    If Right(cell_value, 4) = "0000" Then
        get_holiday = Left(cell_value, 6)
        holiday_date = CDate(get_holiday & Year(toCheck))
    Else
        holiday_date = CDate(cell_value)
    End If
        If Not Weekday(holiday_date) = 1 And Not Weekday(holiday_date) = 7 And holiday_date = toCheck Then
                utl_is_holiday = 1
        End If
Next i

End Function