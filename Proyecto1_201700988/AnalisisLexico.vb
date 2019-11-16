Imports Proyecto1_201700988.Token
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Public Class AnalisisLexico

    'Creando variables para su manejo'
    'Variables para hacer uso en el automata'
    Private Salida As List(Of Token)
    Private Estado As Integer
    Private AuxLex As String
    Private ValAux As String
    Private Texto As String

    Public Function Scaner(ByVal Entrada As String) As List(Of Token)

        'Inicializando las variables'
        'Colocando final a mi entrada'
        Entrada = Entrada + "#"
        Salida = New List(Of Token)
        Estado = 0

        AuxLex = ""
        ValAux = ""

        Dim Val As Char
        'Recorrer toda la entrada, caracter por caracter'
        For i As Integer = 0 To Entrada.Length - 1 Step 1
            Val = Entrada.Chars(i)
            If (i = Entrada.Length - 1) Then
                ValAux = ""
            Else
                ValAux = Entrada.Chars(i + 1)
            End If
            Select Case Estado
                Case 0
                    'Validamos todos los estados
                    If Char.IsLetter(Val) Then
                        Estado = 1
                        AuxLex += Val
                    ElseIf (Val = "{") Then
                        AuxLex += Val
                        addToken(Tipo.LLAVE_IZQ)
                    ElseIf (Val = "}") Then
                        AuxLex += Val
                        addToken(Tipo.LLAVE_DER)
                    ElseIf (Val = "-") Then
                        Estado = 2
                        AuxLex += Val
                    ElseIf Char.IsDigit(Val) Then
                        Estado = 9
                        AuxLex += Val
                    ElseIf (Val = "(") Then
                        AuxLex += Val
                        addToken(Tipo.PARENTESIS_IZQ)
                    ElseIf (Val = ")") Then
                        AuxLex += Val
                        addToken(Tipo.PARENTESIS_DER)
                    ElseIf (Val = "[") Then
                        Estado = 3
                        AuxLex += Val
                    ElseIf (Val = "]") Then
                        AuxLex += Val
                        addToken(Tipo.CORCHETE_DER)
                    ElseIf (Val = """") Or (Val = Chr(147)) Or (Val = Chr(148)) Then
                        Estado = 4
                        AuxLex += Val
                    ElseIf (Val = "/") Then
                        Estado = 5
                        AuxLex += Val
                    ElseIf (Val = ",") Then
                        AuxLex += Val
                        addToken(Tipo.COMA)
                    ElseIf (Val = ";") Then
                        AuxLex += Val
                        addToken(Tipo.PUNTO_COMA)
                    ElseIf (Val = ":") Then
                        AuxLex += Val
                        addToken(Tipo.DOS_PUNTOS)
                    ElseIf (Val = "=") Then
                        AuxLex += Val
                        addToken(Tipo.IGUAL)
                    Else
                        If (Val = "#" And i = Entrada.Length() - 1) Then
                            'Se ha terminado el analisis del contenido total
                            addToken(Tipo.Ultimo)
                            Console.WriteLine("Hemos concluido con el análisis lexico exitosamente.")
                        Else
                            If Val = " " Or Val = "" Then
                                'No tomar en cuenta
                            ElseIf Val = Chr(10) Then

                            Else
                                addError(Tipo.ErrorLex, Val)
                            End If
                            Estado = 0
                        End If
                    End If
               'Iniciando Estado 1
                Case 1
                    If Char.IsLetter(Val) Then
                        Estado = 6
                        AuxLex += Val
                    ElseIf (Val = "_") Then
                        Estado = 7
                        AuxLex += Val
                    ElseIf Char.IsDigit(Val) Then
                        Estado = 8
                        AuxLex += Val
                    Else
                        addToken(Tipo.Identificador)
                        i -= 1
                    End If
               'Inicio Estado 2
                Case 2
                    If Char.IsDigit(Val) Then
                        Estado = 9
                        AuxLex += Val
                    Else
                        If Val = " " Or Val = "" Or Val = Chr(10) Then
                            'No tomar en cuenta
                        Else
                            addError(Tipo.ErrorLex, Val)
                        End If
                        Console.WriteLine("Error sintactico con: " & Val & ", después del signo - se espera un digito (0-9)")
                        Estado = 0
                    End If
                'Inicio Estado 3
                Case 3
                    If (Val = "+") Then
                        Estado = 10
                        AuxLex += Val
                    ElseIf (Val = "*") Then
                        Estado = 11
                        AuxLex += Val
                    Else
                        addToken(Tipo.CORCHETE_IZQ)
                        i -= 1
                    End If
                'Inicio Estado 4
                Case 4
                    If (Val = """") Or (Val = Chr(147)) Or (Val = Chr(148)) Then
                        AuxLex += Val
                        addToken(Tipo.Texto_normal)
                    ElseIf (Val = "#") Then
                        addError(Tipo.ErrorLex, AuxLex)
                    Else
                        Estado = 4
                        AuxLex += Val
                    End If
               'Inicio Estado 5
                Case 5
                    If (Val = "*") Then
                        Estado = 12
                        AuxLex += Val
                    Else
                        If Val = " " Or Val = "" Or Val = Chr(10) Then
                            addError(Tipo.ErrorLex, AuxLex)
                        Else
                            addError(Tipo.ErrorLex, AuxLex)
                            i -= 1
                        End If
                        Console.WriteLine("Error sintactico con: " & Val & ", después del signo / se espera un *")
                        AuxLex = ""
                        Estado = 0
                    End If
                'Iniciando Estado 6
                Case 6
                    If Char.IsLetter(Val) Then
                        Estado = 6
                        AuxLex += Val
                    ElseIf Char.IsDigit(Val) Then
                        Estado = 8
                        AuxLex += Val
                    ElseIf (Val = "_") Then
                        Estado = 7
                        AuxLex += Val
                    Else
                        If (AuxLex.ToLower = "INSTRUCCIONES".ToLower) Then
                            addToken(Tipo.INSTRUCCIONES)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "VARIABLES".ToLower) Then
                            addToken(Tipo.VARIABLES)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "TEXTO".ToLower) Then
                            addToken(Tipo.TEXTO)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Interlineado".ToLower) Then
                            addToken(Tipo.Interlineado)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Entero".ToLower) Then
                            addToken(Tipo.Entero)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Cadena".ToLower) Then
                            addToken(Tipo.Cadena)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Imagen".ToLower) Then
                            addToken(Tipo.Imagen)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Numeros".ToLower) Then
                            addToken(Tipo.Numeros)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Var".ToLower) Then
                            addToken(Tipo.Var)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Promedio".ToLower) Then
                            addToken(Tipo.Promedio)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Suma".ToLower) Then
                            addToken(Tipo.Suma)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Resta".ToLower) Then
                            addToken(Tipo.Resta)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Multiplicar".ToLower) Then
                            addToken(Tipo.Multiplicar)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Division".ToLower) Then
                            addToken(Tipo.Division)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Asignar".ToLower) Then
                            addToken(Tipo.Asignar)
                            i -= 1
                        Else
                            addToken(Tipo.Identificador)
                            i -= 1
                        End If
                    End If
               'Iniciando Estado 7
                Case 7
                    If Char.IsLetter(Val) Then
                        Estado = 13
                        AuxLex += Val
                    Else
                        If Val = " " Or Val = "" Or Val = Chr(10) Then
                            'No tomar en cuenta
                        Else
                            addError(Tipo.ErrorLex, Val)
                        End If
                        Console.WriteLine("Error sintactico con: " & Val & ", después del signo _ se espera una letra")
                        Estado = 0
                    End If
               'Iniciando Estado 8
                Case 8
                    If Char.IsLetterOrDigit(Val) Then
                        Estado = 8
                        AuxLex += Val
                    Else
                        addToken(Tipo.Identificador)
                        i -= 1
                    End If
                'Iniciando Estado 9
                Case 9
                    If Char.IsDigit(Val) Then
                        Estado = 9
                        AuxLex += Val
                    ElseIf (Val = ".") Then
                        Estado = 14
                        AuxLex += Val
                    Else
                        addToken(Tipo.Numero_entero)
                        i -= 1
                    End If
                'Iniciando Estado 10
                Case 10
                    If (Val = "+" And ValAux = "]") Then
                        AuxLex += Val & "]"
                        addToken(Tipo.Texto_negrita)
                        i += 1
                    ElseIf (Val = "#") Then
                        addError(Tipo.ErrorLex, AuxLex)
                    Else
                        Estado = 10
                        AuxLex += Val
                    End If
                'Iniciando Estado 11
                Case 11
                    If (Val = "*" And ValAux = "]") Then
                        AuxLex += Val & "]"
                        addToken(Tipo.Texto_subrayado)
                        i += 1
                    ElseIf (Val = "#") Then
                        addError(Tipo.ErrorLex, AuxLex)
                    Else
                        Estado = 11
                        AuxLex += Val
                    End If
                'Iniciando Estado 12
                Case 12
                    If (Val = "*" And ValAux = "/") Then
                        AuxLex += Val & "/"
                        addToken(Tipo.Comentario)
                        i += 1
                    ElseIf (Val = "#") Then
                        addError(Tipo.ErrorLex, AuxLex)
                    Else
                        Estado = 12
                        AuxLex += Val
                    End If
                'Iniciando Estado 13
                Case 13
                    If Char.IsLetter(Val) Then
                        Estado = 13
                        AuxLex += Val
                    ElseIf (Val = "_") Then
                        Estado = 7
                        AuxLex += Val
                    Else
                        If (AuxLex.ToLower = "Tamanio_letra".ToLower) Then
                            addToken(Tipo.Tamanio_letra)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Nombre_archivo".ToLower) Then
                            addToken(Tipo.Nombre_archivo)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Direccion_archivo".ToLower) Then
                            addToken(Tipo.Direccion_archivo)
                            i -= 1
                        ElseIf (AuxLex.ToLower = "Linea_en_blanco".ToLower) Then
                            addToken(Tipo.Linea_en_blanco)
                            i -= 1
                        Else
                            If Val = " " Or Val = "" Or Val = Chr(10) Then
                                'No tomar en cuenta
                            Else
                                addError(Tipo.ErrorLex, AuxLex)
                                AuxLex = ""
                                i -= 1
                            End If
                            Console.WriteLine("Error sintactico con: " & Val & ", no se encuentra palabra reservada")
                            Estado = 0
                        End If
                    End If
                'Iniciando Estado 14
                Case 14
                    If Char.IsDigit(Val) Then
                        Estado = 15
                        AuxLex += Val
                    Else
                        If Val = " " Or Val = "" Or Val = Chr(10) Then
                            'No tomar en cuenta
                        Else
                            addError(Tipo.ErrorLex, Val)
                        End If
                        Console.WriteLine("Error sintactico con: " & Val & ", después del punto se espera un dígito (0-9)")
                        Estado = 0
                    End If
                'Iniciando Estado 15
                Case 15
                    If Char.IsDigit(Val) Then
                        Estado = 15
                        AuxLex += Val
                    Else
                        addToken(Tipo.Numero_decimal)
                        i -= 1
                    End If
            End Select
        Next
        Return Salida
    End Function

    Private Sub addToken(ByVal tipo As Tipo)
        Salida.Add(New Token(tipo, AuxLex))
        AuxLex = ""
        Estado = 0
    End Sub

    Private Sub addError(ByVal tipo As Tipo, datoError As String)
        Salida.Add(New Token(tipo, datoError))
    End Sub

    Public Function CantidadErrores(ByVal lista As List(Of Token)) As Integer
        Dim Errores As Integer = 0
        For Each Tok As Token In lista
            If Tok.getTipo = 37 Then
                Errores += 1
            End If
        Next
        Return Errores
    End Function

    Public Function CantidadCaracteres(ByVal listas As List(Of Token))
        Dim Total As Integer = 1
        For Each dato As Token In listas
            Total += 1
        Next
        Return Total
    End Function

    Public Sub CrearDocumentoErrores(ByVal m As List(Of Token))
        Try
            Dim ListaErroresPDF As New Document
            Dim EscribirDoc As PdfWriter = PdfWriter.GetInstance(ListaErroresPDF, New FileStream("C:\Users\Luis Fer\Desktop\Lista_Errores_Léxicos.PDF", FileMode.Create))
            Dim columna As Integer = 1
            Dim fila As Integer = 1
            ListaErroresPDF.Open()
            'Colocar logo al documento
            Dim LogoUsac As Image = Image.GetInstance("LOGO.jpg")
            LogoUsac.ScalePercent(CInt(15%))
            LogoUsac.SetAbsolutePosition(400.0F, 720.0F)
            ListaErroresPDF.Add(LogoUsac)
            'Colocar encabezado al documento
            Dim Encabezado As String = "UNIVERSIDAD SAN CARLOS DE GUATEMALA" & vbLf & "FACULTAD DE INGENIERÍA" & vbLf & "ESCUELA DE CIENCIAS" & vbLf & "LENGUAJES FORMALES Y DE PROGRAMACIÓN" & vbLf & vbLf & vbLf
            ListaErroresPDF.Add(New Paragraph(Encabezado))
            'Crear tabla para cargar los errores
            Dim TablaErrores As PdfPTable = New PdfPTable(6)
            Dim cell As New PdfPCell(New Phrase("Listado de Errores Léxicos"))
            Dim Contador As Integer = 1
            cell.Colspan = 6
            cell.HorizontalAlignment = 1
            TablaErrores.AddCell(cell)
            TablaErrores.AddCell("No.")
            TablaErrores.AddCell("Error")
            TablaErrores.AddCell("Tipo")
            TablaErrores.AddCell("Descripción")
            TablaErrores.AddCell("Fila")
            TablaErrores.AddCell("Columna")
            For Each t As Token In m
                If t.getTipoString = "Error Léxico" Then
                    TablaErrores.AddCell(Contador)
                    TablaErrores.AddCell(t.getValor)
                    TablaErrores.AddCell("Léxico")
                    TablaErrores.AddCell(t.getTipoString)
                    TablaErrores.AddCell(fila)
                    TablaErrores.AddCell(columna)
                    If columna < 10 Then
                        columna += 1
                    Else
                        columna = 1
                        fila += 1
                    End If

                    Contador += 1
                Else
                    'No hacer nada
                End If
            Next
            ListaErroresPDF.Add(TablaErrores)
            'Cerrar documento
            ListaErroresPDF.Close()
            Process.Start("C:\Users\Luis Fer\Desktop\Lista_Errores_Léxicos.PDF")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub CrearDocumentoTokens(ByVal n As List(Of Token))
        Try
            Dim ListaTokensPDF As New Document
            Dim Escribir As PdfWriter = PdfWriter.GetInstance(ListaTokensPDF, New FileStream("C:\Users\Luis Fer\Desktop\Lista_Tokens.PDF", FileMode.Create))
            Dim columna As Integer = 1
            Dim fila As Integer = 1
            ListaTokensPDF.Open()
            'Colocar logo al documento
            Dim LogoUsac As Image = Image.GetInstance("LOGO.jpg")
            LogoUsac.ScalePercent(CInt(15%))
            LogoUsac.SetAbsolutePosition(400.0F, 720.0F)
            ListaTokensPDF.Add(LogoUsac)
            'Colocar encabezado al documento
            Dim Encabezado As String = "UNIVERSIDAD SAN CARLOS DE GUATEMALA" & vbLf & "FACULTAD DE INGENIERÍA" & vbLf & "ESCUELA DE CIENCIAS" & vbLf & "LENGUAJES FORMALES Y DE PROGRAMACIÓN" & vbLf & vbLf & vbLf
            ListaTokensPDF.Add(New Paragraph(Encabezado))

            'Crear tabla para agregar al documento
            Dim TablaTokens As PdfPTable = New PdfPTable(5)
            Dim cell As New PdfPCell(New Phrase("Listado de Tokens"))
            Dim Contador As Integer = 1
            cell.Colspan = 5
            cell.HorizontalAlignment = 1
            TablaTokens.AddCell(cell)
            TablaTokens.AddCell("No.")
            TablaTokens.AddCell("Lexema")
            TablaTokens.AddCell("Tipo")
            TablaTokens.AddCell("Fila")
            TablaTokens.AddCell("Columna")
            For Each t As Token In n
                If t.getTipoString = "Error Léxico" Then
                    'No lo agregue a esta lista
                ElseIf t.getTipoString = "Ultimo" Then
                    'No lo agregue a esta lista
                Else
                    TablaTokens.AddCell(Contador)
                    TablaTokens.AddCell(t.getValor)
                    TablaTokens.AddCell(t.getTipoString)
                    TablaTokens.AddCell(fila)
                    TablaTokens.AddCell(columna)
                    If columna < 10 Then
                        columna += 1
                    Else
                        columna = 1
                        fila += 1
                    End If

                    Contador += 1
                End If
            Next
            ListaTokensPDF.Add(TablaTokens)
            'Cerrar documento
            ListaTokensPDF.Close()
            Process.Start("C:\Users\Luis Fer\Desktop\Lista_Tokens.PDF")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub imprimirLista(ByVal l As List(Of Token))
        For Each t As Token In l
            Console.WriteLine(t.getTipoString() & " <--> " & t.getValor)
        Next
    End Sub

End Class


