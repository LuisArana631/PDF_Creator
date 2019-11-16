Imports Proyecto1_201700988.Token
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Public Class AnalisisSintactico

    Dim numPreanalisis As Integer
    Dim preanalisis As Token
    Dim listaTokens As List(Of Token)

    Dim listaErrores(900) As String
    Dim listaDescripcion(900) As String
    Dim contarErrores As Integer = 0


    Public Sub parsear(tok As List(Of Token))
        listaTokens = tok
        preanalisis = listaTokens.Item(0)
        numPreanalisis = 0
        A()
    End Sub

    'Analizar Sintacticamente
    'Ingresar a un segmento
    Private Sub A()
        If preanalisis.getTipo() = Tipo.INSTRUCCIONES Then
            match(Tipo.INSTRUCCIONES)
            match(Tipo.LLAVE_IZQ)
            B()
        ElseIf preanalisis.getTipo() = Tipo.VARIABLES Then
            match(Tipo.VARIABLES)
            match(Tipo.LLAVE_IZQ)
            D()
        ElseIf preanalisis.getTipo() = Tipo.TEXTO Then
            match(Tipo.TEXTO)
            match(Tipo.LLAVE_IZQ)
            C()
        Else
            If Not preanalisis.getTipo = Tipo.Ultimo Then
                Console.WriteLine("Error con: " + preanalisis.getValor + ", se esperaba inicio de segmento")
                listaErrores(contarErrores) = preanalisis.getValor
                listaDescripcion(contarErrores) = "Se esperaba un inicio de segmento"
                contarErrores += 1
                numPreanalisis += 1
                preanalisis = listaTokens.Item(numPreanalisis)
                A()
            End If
        End If
    End Sub
    'Instrucciones
    Private Sub B()
        If preanalisis.getTipo() = Tipo.Nombre_archivo Then
            match(Tipo.Nombre_archivo)
            match(Tipo.PARENTESIS_IZQ)
            match(Tipo.Texto_normal)
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            B()
        ElseIf preanalisis.getTipo() = Tipo.Interlineado Then
            match(Tipo.Interlineado)
            match(Tipo.PARENTESIS_IZQ)
            match(Tipo.Numero_decimal)
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            B()
        ElseIf preanalisis.getTipo() = Tipo.Direccion_archivo Then
            match(Tipo.Direccion_archivo)
            match(Tipo.PARENTESIS_IZQ)
            match(Tipo.Texto_normal)
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            B()
        ElseIf preanalisis.getTipo() = Tipo.Tamanio_letra Then
            match(Tipo.Tamanio_letra)
            match(Tipo.PARENTESIS_IZQ)
            match(Tipo.Numero_entero)
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            B()
        ElseIf preanalisis.getTipo = Tipo.LLAVE_DER Then
            match(Tipo.LLAVE_DER)
            A()
        Else
            If Not preanalisis.getTipo = Tipo.Ultimo Then
                Console.WriteLine("Error con: " + preanalisis.getValor + ", se esperaba instruccion de INSTRUCCIONES")
                listaErrores(contarErrores) = preanalisis.getValor
                listaDescripcion(contarErrores) = "Se esperaba instruccion de INSTRUCCIONES"
                contarErrores += 1
                numPreanalisis += 1
                preanalisis = listaTokens.Item(numPreanalisis)
                B()
            End If
        End If
    End Sub

    'Texto
    Private Sub C()
        If preanalisis.getTipo() = Tipo.Imagen Then
            match(Tipo.Imagen)
            match(Tipo.PARENTESIS_IZQ)
            match(Tipo.Texto_normal)
            match(Tipo.COMA)
            match(Tipo.Numero_entero)
            match(Tipo.COMA)
            match(Tipo.Numero_entero)
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.Texto_normal Then
            match(Tipo.Texto_normal)
            C()
        ElseIf preanalisis.getTipo = Tipo.Texto_negrita Then
            match(Tipo.Texto_negrita)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.Texto_subrayado Then
            match(Tipo.Texto_subrayado)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.Linea_en_blanco Then
            match(Tipo.Linea_en_blanco)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.Comentario Then
            match(Tipo.Comentario)
            C()
        ElseIf preanalisis.getTipo = Tipo.Var Then
            match(Tipo.Var)
            match(Tipo.CORCHETE_IZQ)
            match(Tipo.Identificador)
            match(Tipo.CORCHETE_DER)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.Asignar Then
            match(Tipo.Asignar)
            match(Tipo.PARENTESIS_IZQ)
            match(Tipo.Identificador)
            match(Tipo.COMA)
            G()
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.Numeros Then
            match(Tipo.Numeros)
            match(Tipo.PARENTESIS_IZQ)
            H()
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.Promedio Then
            match(Tipo.Promedio)
            match(Tipo.PARENTESIS_IZQ)
            I()
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.Suma Then
            match(Tipo.Suma)
            match(Tipo.PARENTESIS_IZQ)
            I()
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.Resta Then
            match(Tipo.Resta)
            match(Tipo.PARENTESIS_IZQ)
            I()
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.Division Then
            match(Tipo.Division)
            match(Tipo.PARENTESIS_IZQ)
            I()
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.Multiplicar Then
            match(Tipo.Multiplicar)
            match(Tipo.PARENTESIS_IZQ)
            I()
            match(Tipo.PARENTESIS_DER)
            match(Tipo.PUNTO_COMA)
            C()
        ElseIf preanalisis.getTipo = Tipo.LLAVE_DER Then
            match(Tipo.LLAVE_DER)
            A()
        Else
            If Not preanalisis.getTipo = Tipo.Ultimo Then
                Console.WriteLine("Error con: " + preanalisis.getValor + ", se esperaba instruccion de TEXTO")
                listaErrores(contarErrores) = preanalisis.getValor
                listaDescripcion(contarErrores) = "Se esperaba instruccion de TEXTO"
                contarErrores += 1
                numPreanalisis += 1
                preanalisis = listaTokens.Item(numPreanalisis)
                C()
            End If
        End If
    End Sub

    Private Sub G()
        If preanalisis.getTipo = Tipo.Texto_normal Then
            match(Tipo.Texto_normal)
        Else
            match(Tipo.Numero_entero)
        End If
    End Sub

    Private Sub GP()
        If preanalisis.getTipo = Tipo.Texto_normal Then
            match(Tipo.Texto_normal)
        ElseIf preanalisis.getTipo = Tipo.Identificador Then
            match(Tipo.Identificador)
        Else
            match(Tipo.Numero_entero)
        End If
    End Sub

    Private Sub H()
        GP()
        HP()
    End Sub

    Private Sub HP()
        If preanalisis.getTipo = Tipo.COMA Then
            match(Tipo.COMA)
            GP()
            HP()
        End If
    End Sub

    Private Sub I()
        IP()
        J()
    End Sub

    Private Sub IP()
        If preanalisis.getTipo = Tipo.Numero_entero Then
            match(Tipo.Numero_entero)
        Else
            match(Tipo.Identificador)
        End If
    End Sub

    Private Sub J()
        If preanalisis.getTipo = Tipo.COMA Then
            match(Tipo.COMA)
            IP()
            J()
        End If
    End Sub

    'Variables
    Private Sub D()
        If preanalisis.getTipo() = Tipo.Identificador Then
            match(Tipo.Identificador)
            DP()
        ElseIf preanalisis.getTipo = Tipo.LLAVE_DER Then
            match(Tipo.LLAVE_DER)
            A()
        Else
            If Not preanalisis.getTipo = Tipo.Ultimo Then
                Console.WriteLine("Error con: " + preanalisis.getValor + ", se esperaba identificador")
                listaErrores(contarErrores) = preanalisis.getValor
                listaDescripcion(contarErrores) = "Se esperaba identificador"
                contarErrores += 1
                numPreanalisis += 1
                preanalisis = listaTokens.Item(numPreanalisis)
                D()
            End If
        End If
    End Sub

    Private Sub DP()
        If preanalisis.getTipo() = Tipo.COMA Then
            match(Tipo.COMA)
            match(Tipo.Identificador)
            DP()
        Else
            match(Tipo.DOS_PUNTOS)
            E()
        End If
    End Sub

    Private Sub E()
        If preanalisis.getTipo = Tipo.Entero Then
            match(Tipo.Entero)
            F()
        Else
            match(Tipo.Cadena)
            FP()
        End If
    End Sub

    Private Sub F()
        If preanalisis.getTipo = Tipo.IGUAL Then
            match(Tipo.IGUAL)
            match(Tipo.Numero_entero)
            match(Tipo.PUNTO_COMA)
            D()
        Else
            match(Tipo.PUNTO_COMA)
            D()
        End If
    End Sub

    Private Sub FP()
        If preanalisis.getTipo = Tipo.IGUAL Then
            match(Tipo.IGUAL)
            match(Tipo.Texto_normal)
            match(Tipo.PUNTO_COMA)
            D()
        Else
            match(Tipo.PUNTO_COMA)
            D()
        End If
    End Sub

    'Metodos auxiliares
    Public Sub match(p As Tipo)
        If Not p = preanalisis.getTipo Then
            Console.WriteLine("Error con: " + preanalisis.getValor() + ", se esperaba " + getTipoError(p))
            listaErrores(contarErrores) = preanalisis.getValor
            listaDescripcion(contarErrores) = "Se esperaba " + getTipoError(p)
            contarErrores += 1
        End If
        If Not preanalisis.getTipo = Tipo.Ultimo Then
            numPreanalisis += 1
            preanalisis = listaTokens.Item(numPreanalisis)
        Else
            Console.WriteLine("Hemos concluido con el analisis sintactico exitosamente")
        End If
    End Sub

    Private Function getTipoError(p As Tipo) As String
        Select Case p
            Case Tipo.Asignar
                Return "Asignar"
            Case Tipo.Cadena
                Return "Cadena"
            Case Tipo.COMA
                Return "Coma"
            Case Tipo.Comentario
                Return "Comentario"
            Case Tipo.CORCHETE_DER
                Return "Corechete Derecho"
            Case Tipo.CORCHETE_IZQ
                Return "Corchete Izquierdo"
            Case Tipo.Direccion_archivo
                Return "Direccion_archivo"
            Case Tipo.Division
                Return "Division"
            Case Tipo.DOS_PUNTOS
                Return "Dos puntos"
            Case Tipo.Entero
                Return "Entero"
            Case Tipo.Identificador
                Return "Identificador"
            Case Tipo.IGUAL
                Return "Igual"
            Case Tipo.Imagen
                Return "Imagen"
            Case Tipo.INSTRUCCIONES
                Return "Instrucciones"
            Case Tipo.Interlineado
                Return "Interlineado"
            Case Tipo.Linea_en_blanco
                Return "Salto de linea"
            Case Tipo.LLAVE_DER
                Return "Llave Derecha"
            Case Tipo.LLAVE_IZQ
                Return "Llave Izquierda"
            Case Tipo.Multiplicar
                Return "Multiplicar"
            Case Tipo.Nombre_archivo
                Return "Nombre_archivo"
            Case Tipo.Numeros
                Return "Numeros"
            Case Tipo.Numero_decimal
                Return "Numero"
            Case Tipo.Numero_entero
                Return "Numero"
            Case Tipo.PARENTESIS_DER
                Return "Parentesis Derecho"
            Case Tipo.PARENTESIS_IZQ
                Return "Parentesis Izquierdo"
            Case Tipo.Promedio
                Return "Promedio"
            Case Tipo.PUNTO_COMA
                Return "Punto y coma"
            Case Tipo.Resta
                Return "Resta"
            Case Tipo.Suma
                Return "Suma"
            Case Tipo.Tamanio_letra
                Return "Tamaño_letra"
            Case Tipo.TEXTO
                Return "Texto"
            Case Tipo.Texto_negrita
                Return "Texto en negrita"
            Case Tipo.Texto_normal
                Return "Texto normal"
            Case Tipo.Texto_subrayado
                Return "Texto subrayado"
            Case Tipo.Var
                Return "Var"
            Case Tipo.VARIABLES
                Return "Variable"
            Case Else
                Return "Desconocido"
        End Select
    End Function

    Public Sub CrearDocumentoErrorSintactico()
        Try
            Dim ListaErroresSinPDF As New Document
            Dim EscribirDoc As PdfWriter = PdfWriter.GetInstance(ListaErroresSinPDF, New FileStream("C:\Users\Luis Fer\Desktop\Lista_Errores_Sintácticos.PDF", FileMode.Create))
            ListaErroresSinPDF.Open()
            'Colocar logo al documento
            Dim LogoUsac As Image = Image.GetInstance("LOGO.jpg")
            LogoUsac.ScalePercent(CInt(15%))
            LogoUsac.SetAbsolutePosition(400.0F, 720.0F)
            ListaErroresSinPDF.Add(LogoUsac)
            'Colocar encabezado al documento
            Dim Encabezado As String = "UNIVERSIDAD SAN CARLOS DE GUATEMALA" & vbLf & "FACULTAD DE INGENIERÍA" & vbLf & "ESCUELA DE CIENCIAS" & vbLf & "LENGUAJES FORMALES Y DE PROGRAMACIÓN" & vbLf & vbLf & vbLf
            ListaErroresSinPDF.Add(New Paragraph(Encabezado))
            'Crear tabla para cargar los errores
            Dim TablaErrores As PdfPTable = New PdfPTable(6)
            Dim cell As New PdfPCell(New Phrase("Listado de Errores Sintácticos"))
            Dim Contador As Integer = 1
            Dim recorridolistades As Integer = 0
            Dim fila As Integer = 1
            Dim columna As Integer = 1
            cell.Colspan = 6
            cell.HorizontalAlignment = 1
            TablaErrores.AddCell(cell)
            TablaErrores.AddCell("No.")
            TablaErrores.AddCell("Error")
            TablaErrores.AddCell("Tipo")
            TablaErrores.AddCell("Descripción")
            TablaErrores.AddCell("Fila")
            TablaErrores.AddCell("Columna")
            'Llenar la tabla 
            For Each er As String In listaErrores
                If er = "" Then
                    'ni madres
                Else
                    TablaErrores.AddCell(Contador)
                    TablaErrores.AddCell(er)
                    TablaErrores.AddCell("Sintáctico")
                    TablaErrores.AddCell(listaDescripcion(recorridolistades))
                    TablaErrores.AddCell(fila)
                    TablaErrores.AddCell(columna)
                    recorridolistades += 1
                    Contador += 1
                End If
            Next
            'Agregar tabla al documento
            ListaErroresSinPDF.Add(TablaErrores)
            'Cerrar documento
            ListaErroresSinPDF.Close()
            Process.Start("C:\Users\Luis Fer\Desktop\Lista_Errores_Sintácticos.PDF")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class

