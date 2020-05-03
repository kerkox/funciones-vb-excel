Function CalcularFechasDiasRetiro(fecha_ingreso As Date, fecha_retiro As Date, fecha_ultimo_pago As Date)
    
    Dim fecha_ingreso_texto As String
    Dim fecha_ingreso_organizada As Date
    Dim fecha_ultimo_pago_texto As String
    Dim fecha_ultimo_pago_organizada As Date
    Dim cantidad_dias_ultimo_pago As Integer
    Dim anio As Integer
    
    fecha_ultimo_pago_texto = "1/" & Month(fecha_ultimo_pago) & "/" & Year(fecha_ultimo_pago)
    fecha_ultimo_pago_organizada = DateValue(fecha_ultimo_pago_texto)
    fecha_ultimo_pago_organizada = DateAdd("m", 1, fecha_ultimo_pago_organizada)
    fecha_ultimo_pago_organizada = DateAdd("d", -1, fecha_ultimo_pago_organizada)
    If (Day(fecha_ultimo_pago_organizada) = 31) Then
        fecha_ultimo_pago_organizada = DateAdd("d", -1, fecha_ultimo_pago_organizada)
    End If
    cantidad_dias_ultimo_pago = Application.WorksheetFunction.Days360(fecha_ultimo_pago_organizada, fecha_retiro)
    
    If (Year(fecha_retiro) = Year(fecha_ultimo_pago_organizada)) Then
        If (cantidad_dias_ultimo_pago > 0) Then
            
            
            If (Month(fecha_ingreso) > Month(fecha_retiro)) Then
                anio = Year(fecha_retiro) - 1
            ElseIf (Month(fecha_ingreso) < Month(fecha_retiro)) Then
                anio = Year(fecha_retiro)
            ElseIf (Day(fecha_ingreso) > Day(fecha_retiro)) Then
                anio = Year(fecha_retiro) - 1
            ElseIf (Day(fecha_ingreso) < Day(fecha_retiro)) Then
                anio = Year(fecha_retiro)
            Else
                anio = Year(fecha_retiro)
            End If
            
            fecha_ingreso_texto = Day(fecha_ingreso) & "/" & Month(fecha_ingreso) & "/" & anio
            fecha_ingreso_organizada = DateValue(fecha_ingreso_texto)
            
            cantidad_dias_ultimo_pago = Application.WorksheetFunction.Days360(fecha_ingreso_organizada, fecha_retiro)
            
            
        End If
    End If
    CalcularFechasDiasRetiro = cantidad_dias_ultimo_pago
    
End Function