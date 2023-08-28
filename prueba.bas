
Option Explicit

'Arreglo de tipo Elemento
Type Elemento
    Peso As Integer
    Calorias As Integer
End Type

'Procedimiento 
Sub EncontrarElementosViables()
    Dim numElementos As Integer
    numElementos = 5 ' Número total de elementos
    
    Dim pesoMaximo As Integer
    Dim caloriasMinimas As Integer
    pesoMaximo = 10 ' Peso máximo 
    caloriasMinimas = 20 ' Calorías mínimas 
    
    Dim elementos() As Elemento
    ReDim elementos(1 To numElementos)
    
    elementos(1).Peso = 4
    elementos(1).Calorias = 12
    elementos(2).Peso = 2
    elementos(2).Calorias = 7
    elementos(3).Peso = 9
    elementos(3).Calorias = 9
    elementos(4).Peso = 8
    elementos(4).Calorias = 10
    elementos(5).Peso = 4
    elementos(5).Calorias = 1
    
    Dim mejorConjunto() As Boolean
    Dim mejorPeso As Integer
    Dim mejorCalorias As Integer
    Dim i As Integer, j As Integer
    
    ReDim mejorConjunto(1 To numElementos)
    mejorPeso = 0
    mejorCalorias = 0
    
    ' Algoritmo de fuerza bruta para probar todas las combinaciones
    For i = 1 To 2 ^ numElementos - 1
        Dim pesoActual As Integer
        Dim caloriasActuales As Integer
        Dim conjuntoActual() As Boolean
        ReDim conjuntoActual(1 To numElementos)
        
        ' Generar el conjunto actual
        For j = 1 To numElementos
            conjuntoActual(j) = (i And 2 ^ (j - 1)) > 0
        Next j
        
        ' Calculo de peso y calorías totales del conjunto
        pesoActual = 0
        caloriasActuales = 0
        For j = 1 To numElementos
            If conjuntoActual(j) Then
                pesoActual = pesoActual + elementos(j).Peso
                caloriasActuales = caloriasActuales + elementos(j).Calorias
            End If
        Next j
        
        ' Verificación del conjunto actual
        If pesoActual <= pesoMaximo And caloriasActuales >= caloriasMinimas Then
            If mejorPeso = 0 Or (pesoActual < mejorPeso) Or (pesoActual = mejorPeso And caloriasActuales > mejorCalorias) Then
                mejorPeso = pesoActual
                mejorCalorias = caloriasActuales
                mejorConjunto = conjuntoActual
            End If
        End If
    Next i
    
    ' Muestro el contenido de la lista y el resultado
    Dim resultado As String
    resultado = "Elementos viables: " 
    For i = 1 To numElementos
        If mejorConjunto(i) Then
            resultado = resultado & "     - " & "E" & i & " Peso: " & elementos(i).Peso & " Calorías: " & elementos(i).Calorias
        End If
    Next i
    
    if resultado = "Elementos viables:" then
    	msgbox "No es posible escoger un elemento"
    else
    	MsgBox resultado
    end if
    
End Sub

