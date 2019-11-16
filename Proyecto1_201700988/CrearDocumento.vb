Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Public Class CrearDocumento
    'Creando documento
    Public Sub CrearDocumento(ByVal list As List(Of Token))
        Dim contarerrores As Integer = 0
        For Each erro As Token In list
            If erro.getTipo = Token.Tipo.ErrorLex Then
                contarerrores += 1
            Else
                'No haga nada
            End If
        Next
        If contarerrores = 0 Then
            'Creacion de los parametros por medio de las instrucciones
            'Para seleccionar funcion de la palabra reservada
            Dim Funcion As Integer = 0
            'Para tener los componentesPDF
            Dim ArchivoNombre As String = "Resultado.PDF"
            Dim ArchivoDir As String = "C:\Users\Luis Fer\Desktop"
            Dim archivoRuta As String
            Dim Interlineado As Double = 1.5
            Dim TamañoLetra As Integer = 11
            'Utilizar las funciones
            For Each tok As Token In list
                'Validando  todas las funciones
                'Funciones de instrucciones
                If tok.getTipo = Token.Tipo.Nombre_archivo Then
                    Funcion = 1
                ElseIf tok.getTipo = Token.Tipo.Direccion_archivo Then
                    Funcion = 2
                ElseIf tok.getTipo = Token.Tipo.Interlineado Then
                    Funcion = 3
                ElseIf tok.getTipo = Token.Tipo.Tamanio_letra Then
                    Funcion = 4
                End If
                'Llenando los campos con las instrucciones
                Select Case Funcion
                    Case 1
                        If tok.getTipo = Token.Tipo.Texto_normal Then
                            ArchivoNombre = tok.getValor
                            ArchivoNombre = Strings.Replace(ArchivoNombre, """", "")
                            ArchivoNombre = Strings.Replace(ArchivoNombre, Chr(148), "")
                            ArchivoNombre = Strings.Replace(ArchivoNombre, Chr(147), "")
                            Funcion = 0
                        ElseIf tok.getTipo = Token.Tipo.PARENTESIS_IZQ Then
                            'nada
                        ElseIf Tok.getTipo = Token.Tipo.PARENTESIS_DER Or Tok.gettipo = Token.Tipo.PUNTO_COMA Then
                            Funcion = 0
                        ElseIf tok.getTipo = Token.Tipo.Nombre_archivo Then
                            'nada
                        Else
                            MessageBox.Show("Debe de ir una cadena de texto normal dentro de la funcion 'Nombre_archivo'", "Error Sintactico", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        End If
                    Case 2
                        If tok.getTipo = Token.Tipo.Texto_normal Then
                            ArchivoDir = tok.getValor
                            ArchivoDir = Strings.Replace(ArchivoDir, """", "")
                            ArchivoDir = Strings.Replace(ArchivoDir, Chr(147), "")
                            ArchivoDir = Strings.Replace(ArchivoDir, Chr(148), "")
                            Funcion = 0
                        ElseIf tok.getTipo = Token.Tipo.PARENTESIS_IZQ Then
                            'nada
                        ElseIf Tok.getTipo = Token.Tipo.PARENTESIS_DER Or Tok.gettipo = Token.Tipo.PUNTO_COMA Then
                            Funcion = 0
                        ElseIf tok.getTipo = Token.Tipo.Direccion_archivo Then
                            'nada
                        Else
                            MessageBox.Show("Debe de ir una cadena de texto normal dentro de la funcion 'Direccion_archivo'", "Error Sintactico", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        End If
                    Case 3
                        If tok.getTipo = Token.Tipo.Numero_decimal Then
                            Interlineado = tok.getValor
                            Funcion = 0
                        ElseIf tok.getTipo = Token.Tipo.PARENTESIS_IZQ Then
                            'nada
                        ElseIf Tok.getTipo = Token.Tipo.PARENTESIS_DER Or Tok.gettipo = Token.Tipo.PUNTO_COMA Then
                            Funcion = 0
                        ElseIf tok.getTipo = Token.Tipo.Interlineado Then
                            'nada
                        Else
                            MessageBox.Show("Debe de ir un numero decimal dentro de la funcion 'Interlineado'", "Error Sintactico", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        End If
                    Case 4
                        If tok.getTipo = Token.Tipo.Numero_entero Then
                            TamañoLetra = tok.getValor
                            Funcion = 0
                        ElseIf tok.getTipo = Token.Tipo.PARENTESIS_IZQ Then
                            'nada
                        ElseIf Tok.getTipo = Token.Tipo.PARENTESIS_DER Or Tok.gettipo = Token.Tipo.PUNTO_COMA Then
                            Funcion = 0
                        ElseIf tok.getTipo = Token.Tipo.Tamanio_letra Then
                            'nada
                        Else
                            MessageBox.Show("Debe de ir un numero entero dentro de la funcion 'Tamanio_letra'", "Error Sintactico", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        End If
                End Select
            Next
            'Creacion de las variables para su uso
            Dim listavar As List(Of String) = New List(Of String)
            Dim listval As List(Of String) = New List(Of String)
            Dim Asignar As Integer = 0
            Dim RecorrerAsig As Integer = 0
            Dim varyes As Boolean = False
            For Each t As Token In list
                If t.getTipo = Token.Tipo.VARIABLES Then
                    varyes = True
                ElseIf t.getTipo = Token.Tipo.TEXTO Then
                    varyes = False
                ElseIf t.getTipo = Token.Tipo.INSTRUCCIONES Then
                    varyes = False
                End If
                If varyes Then
                    If t.getTipo = Token.Tipo.Identificador Then
                        listavar.Add(t.getValor)
                        Asignar += 1
                    ElseIf t.getTipo = Token.Tipo.Numero_entero Then
                        While RecorrerAsig < Asignar
                            listval.Add(t.getValor)
                            RecorrerAsig += 1
                        End While
                        RecorrerAsig = 0
                        Asignar = 0
                    ElseIf t.getTipo = Token.Tipo.Texto_normal Then
                        Dim auxvariables As String
                        auxvariables = t.getValor
                        auxvariables = Strings.Replace(auxvariables, """", "")
                        auxvariables = Strings.Replace(auxvariables, Chr(148), "")
                        auxvariables = Strings.Replace(auxvariables, Chr(147), "")
                        While RecorrerAsig < Asignar
                            listval.Add(auxvariables)
                            RecorrerAsig += 1
                        End While
                        RecorrerAsig = 0
                        Asignar = 0
                    ElseIf t.getTipo = Token.Tipo.LLAVE_DER Then
                        varyes = False
                    End If
                End If
            Next
            Dim reco As Integer = 0
            While reco < listavar.Count
                If reco < listval.Count Then
                    Console.WriteLine("variable: " & listavar(reco) & " valor: " & listval(reco))
                    reco += 1
                Else
                    Console.WriteLine("variable: " & listavar(reco))
                    reco += 1
                End If
            End While
            'Textos
            Dim listaTextos As List(Of String) = New List(Of String)
            Dim listatipo As List(Of String) = New List(Of String)
            Dim auxiliarTexto As String = ""
            Dim textoyes As Boolean = False
            Dim Listadovar As Boolean = False
            Dim ImagenLista As Boolean = False
            Dim mostrarVar As Boolean = False
            Dim asignarvar As Boolean = False
            Dim valx As List(Of Integer) = New List(Of Integer)
            Dim valy As List(Of Integer) = New List(Of Integer)
            Dim Promyes As Boolean = False
            Dim ListProm As List(Of Integer) = New List(Of Integer)
            Dim sumyes As Boolean = False
            Dim listsum As List(Of Integer) = New List(Of Integer)
            'Variables del proyecto 2
            Dim restyes As Boolean = False
            Dim listRes As List(Of Integer) = New List(Of Integer)
            Dim multyes As Boolean = False
            Dim listMult As List(Of Integer) = New List(Of Integer)
            Dim divyes As Boolean = False
            Dim listDiv As List(Of Integer) = New List(Of Integer)

            '  Dim ValaCambiar As Integer
            For Each t As Token In list
                If t.getTipo = Token.Tipo.TEXTO Then
                    textoyes = True
                ElseIf t.getTipo = Token.Tipo.INSTRUCCIONES Then
                    textoyes = False
                ElseIf t.getTipo = Token.Tipo.VARIABLES Then
                    textoyes = False
                End If
                If t.getTipo = Token.Tipo.Numeros Then
                    Listadovar = True
                End If
                If t.getTipo = Token.Tipo.Imagen Then
                    ImagenLista = True
                End If
                If t.getTipo = Token.Tipo.Var Then
                    mostrarVar = True
                End If
                If t.getTipo = Token.Tipo.Promedio Then
                    Promyes = True
                End If
                If t.getTipo = Token.Tipo.Suma Then
                    sumyes = True
                End If
                If t.getTipo = Token.Tipo.Resta Then
                    restyes = True
                End If
                If t.getTipo = Token.Tipo.Division Then
                    divyes = True
                End If
                If t.getTipo = Token.Tipo.Multiplicar Then
                    multyes = True
                End If
                'If t.getTipo = Token.Tipo.Asignar Then
                'asignarvar = True
                'End If
                If textoyes Then
                    If Listadovar Then
                        If t.getTipo = Token.Tipo.Texto_normal Then
                            auxiliarTexto = t.getValor
                            auxiliarTexto = Strings.Replace(auxiliarTexto, """", "")
                            auxiliarTexto = Strings.Replace(auxiliarTexto, Chr(148), "")
                            auxiliarTexto = Strings.Replace(auxiliarTexto, Chr(147), "")
                            listaTextos.Add(auxiliarTexto)
                            listatipo.Add("listado")
                        ElseIf t.getTipo = Token.Tipo.PUNTO_COMA Then
                            Listadovar = False
                        End If
                    ElseIf ImagenLista Then
                        If t.getTipo = Token.Tipo.Texto_normal Then
                            auxiliarTexto = t.getValor
                            auxiliarTexto = Strings.Replace(auxiliarTexto, """", "")
                            auxiliarTexto = Strings.Replace(auxiliarTexto, Chr(148), "")
                            auxiliarTexto = Strings.Replace(auxiliarTexto, Chr(147), "")
                            listaTextos.Add(auxiliarTexto)
                            listatipo.Add("imagen")
                        ElseIf t.getTipo = Token.Tipo.Numero_entero Then
                            auxiliarTexto = t.getValor
                            If valx.Count = valy.Count Then
                                Console.WriteLine("Ingresando valx: " & auxiliarTexto)
                                valx.Add(auxiliarTexto)
                            Else
                                Console.WriteLine("Ingresando valy: " & auxiliarTexto)
                                valy.Add(auxiliarTexto)
                            End If
                        ElseIf t.getTipo = Token.Tipo.PUNTO_COMA Then
                            Listadovar = False
                            ImagenLista = False
                        End If
                    ElseIf mostrarVar Then
                        If t.getTipo = Token.Tipo.Identificador Then
                            listatipo.Add("variable")
                            Dim recono As Integer = 0
                            Dim encontrovar As Boolean = False
                            While recono < listavar.Count
                                If listavar(recono) = t.getValor Then
                                    If recono < listval.Count Then
                                        listaTextos.Add(listval(recono))
                                        encontrovar = True
                                    Else
                                        'Nothing
                                    End If
                                    mostrarVar = False
                                End If
                                recono += 1
                            End While
                            If encontrovar Then
                                encontrovar = False
                            Else
                                MessageBox.Show("No se encontro la variable de dicha función", "Error con la función Var")
                                encontrovar = False
                            End If
                        End If
                    ElseIf Promyes Then
                        If t.getTipo = Token.Tipo.Numero_entero Then
                            ListProm.Add(t.getValor)
                        ElseIf t.getTipo = Token.Tipo.Identificador Then
                            Dim recono As Integer = 0
                            While recono < listval.Count
                                If listavar(recono) = t.getValor Then
                                    If recono < listval.Count Then
                                        ListProm.Add(listval(recono))
                                    Else
                                        'no  haga nada
                                    End If
                                    mostrarVar = False
                                End If
                                recono += 1
                            End While
                        ElseIf t.getTipo = Token.Tipo.PARENTESIS_DER Or t.getTipo = Token.Tipo.PUNTO_COMA Then
                            Dim sumaProm As Integer = 0
                            Dim valores As Integer = 0
                            Dim prom As Decimal = 0
                            For Each valorlista As Integer In ListProm
                                sumaProm += valorlista
                                valores += 1
                            Next
                            prom = sumaProm / valores
                            listatipo.Add("Promedio")
                            listaTextos.Add(prom)
                            ListProm.Clear()
                            Promyes = False
                        End If
                    ElseIf multyes Then
                        If t.getTipo = Token.Tipo.Numero_entero Then
                            listMult.Add(t.getValor)
                        ElseIf t.getTipo = Token.Tipo.Identificador Then
                            Dim reconocer As Integer = 0
                            While reconocer < listavar.Count
                                If listavar(reconocer) = t.getValor Then
                                    If reconocer < listval.Count Then
                                        listMult.Add(listval(reconocer))
                                    Else
                                        'nada
                                    End If
                                    mostrarVar = False
                                End If
                                reconocer += 1
                            End While
                        ElseIf t.getTipo = Token.Tipo.PARENTESIS_DER Or t.getTipo = Token.Tipo.PUNTO_COMA Then
                            Dim multTotal As Integer = 1
                            For Each valmult As Integer In listMult
                                multTotal = valmult * multTotal
                            Next
                            listatipo.Add("mult")
                            listaTextos.Add(multTotal)
                            listMult.Clear()
                            multyes = False
                        End If
                    ElseIf restyes Then
                        If t.getTipo = Token.Tipo.Numero_entero Then
                            listRes.Add(t.getValor)
                        ElseIf t.getTipo = Token.Tipo.Identificador Then
                            Dim reconocer As Integer = 0
                            While reconocer < listavar.Count
                                If listavar(reconocer) = t.getValor Then
                                    If reconocer < listval.Count Then
                                        listRes.Add(listval(reconocer))
                                    Else
                                        'nada
                                    End If
                                    mostrarVar = False
                                End If
                                reconocer += 1
                            End While
                        ElseIf t.getTipo = Token.Tipo.PARENTESIS_DER Or t.getTipo = Token.Tipo.PUNTO_COMA Then
                            Dim restotal As Integer = 0
                            Dim recorrido As Integer = 0
                            For Each valrest As Integer In listRes
                                If recorrido = 0 Then
                                    restotal = valrest
                                Else
                                    restotal = restotal - valrest
                                End If
                                recorrido += 1
                            Next
                            listatipo.Add("rest")
                            listaTextos.Add(restotal)
                            listRes.Clear()
                            restyes = False
                        End If
                    ElseIf divyes Then
                        If t.getTipo = Token.Tipo.Numero_entero Then
                            listDiv.Add(t.getValor)
                        ElseIf t.getTipo = Token.Tipo.Identificador Then
                            Dim reconocer As Integer = 0
                            While reconocer < listavar.Count
                                If listavar(reconocer) = t.getValor Then
                                    If reconocer < listval.Count Then
                                        listDiv.Add(listval(reconocer))
                                    Else
                                        'nada
                                    End If
                                    mostrarVar = False
                                End If
                                reconocer += 1
                            End While
                        ElseIf t.getTipo = Token.Tipo.PARENTESIS_DER Or t.getTipo = Token.Tipo.PUNTO_COMA Then
                            Dim divtotal As Double
                            Dim recorrido As Integer = 0
                            For Each valdiv As Integer In listDiv
                                If recorrido = 0 Then
                                    divtotal = valdiv
                                Else
                                    divtotal = divtotal / valdiv
                                End If
                            Next
                            listatipo.Add("div")
                            listaTextos.Add(divtotal)
                            listDiv.Clear()
                            divyes = False
                        End If
                    ElseIf sumyes Then
                        If t.getTipo = Token.Tipo.Numero_entero Then
                            listsum.Add(t.getValor)
                        ElseIf t.getTipo = Token.Tipo.Identificador Then
                            Dim reconocer As Integer = 0
                            While reconocer < listavar.Count
                                If listavar(reconocer) = t.getValor Then
                                    If reconocer < listval.Count Then
                                        listsum.Add(listval(reconocer))
                                    Else
                                        'Ni madres
                                    End If
                                    mostrarVar = False
                                End If
                                reconocer += 1
                            End While
                        ElseIf t.getTipo = Token.Tipo.PARENTESIS_DER Or t.getTipo = Token.Tipo.PUNTO_COMA Then
                            Dim sumatotal As Integer = 0
                            For Each valsuma As Integer In listsum
                                sumatotal += valsuma
                            Next
                            listatipo.Add("suma")
                            listaTextos.Add(sumatotal)
                            listsum.Clear()
                            sumyes = False
                        End If
                        'ElseIf asignarvar Then
                        '   If t.getTipo = Token.Tipo.Identificador Then
                        '  For Each variable As String In listavar
                        ' If variable = t.getValor Then
                        'Exit For
                        'ElseIf t.getTipo = Token.Tipo.Numero_entero Or t.getTipo = Token.Tipo.Texto_normal Then
                        '   listval(ValaCambiar) = t.getValor
                        'ElseIf t.getTipo = Token.Tipo.PARENTESIS_DER Or t.getTipo = Token.Tipo.PUNTO_COMA Then
                        '   asignarvar = False
                        'End If
                        '   ValaCambiar += 1
                        'Next
                        '       End If
                    Else
                        If t.getTipo = Token.Tipo.Texto_normal Then
                            auxiliarTexto = t.getValor
                            auxiliarTexto = Strings.Replace(auxiliarTexto, """", "")
                            auxiliarTexto = Strings.Replace(auxiliarTexto, Chr(148), "")
                            auxiliarTexto = Strings.Replace(auxiliarTexto, Chr(147), "")
                            listaTextos.Add(auxiliarTexto)
                            listatipo.Add("normal")
                        ElseIf t.getTipo = Token.Tipo.Texto_negrita Then
                            auxiliarTexto = t.getValor
                            auxiliarTexto = Strings.Replace(auxiliarTexto, "[", "")
                            auxiliarTexto = Strings.Replace(auxiliarTexto, "]", "")
                            auxiliarTexto = Strings.Replace(auxiliarTexto, "+", "")
                            listaTextos.Add(auxiliarTexto)
                            listatipo.Add("negrita")
                        ElseIf t.getTipo = Token.Tipo.Texto_subrayado Then
                            auxiliarTexto = t.getValor
                            auxiliarTexto = Strings.Replace(auxiliarTexto, "[", "")
                            auxiliarTexto = Strings.Replace(auxiliarTexto, "]", "")
                            auxiliarTexto = Strings.Replace(auxiliarTexto, "*", "")
                            listaTextos.Add(auxiliarTexto)
                            listatipo.Add("subrayada")
                        ElseIf t.getTipo = Token.Tipo.Linea_en_blanco Then
                            listaTextos.Add("salto")
                            listatipo.Add("salto")
                        End If
                    End If
                End If
            Next

            archivoRuta = ArchivoDir + "\" + ArchivoNombre
            Try
                Dim DocumentoResult As New Document
                Dim EscribirDoc As PdfWriter = PdfWriter.GetInstance(DocumentoResult, New FileStream(archivoRuta, FileMode.Create))
                DocumentoResult.Open()
                'Escribir los textos
                Dim recorrerListaTexto As Integer = 0
                Dim listado As Integer = 1
                Dim imagenesColocadas As Integer = 0
                While recorrerListaTexto < listaTextos.Count
                    Console.WriteLine(listaTextos(recorrerListaTexto))
                    If listaTextos(recorrerListaTexto) = "salto" Then
                        DocumentoResult.Add(New Paragraph(vbLf))
                        recorrerListaTexto += 1
                    Else
                        If listatipo(recorrerListaTexto) = "normal" Then
                            DocumentoResult.Add(New Paragraph(listaTextos(recorrerListaTexto), New Font(FontStyle.Regular, TamañoLetra)))
                            recorrerListaTexto += 1
                            listado = 1
                        ElseIf listatipo(recorrerListaTexto) = "negrita" Then
                            DocumentoResult.Add(New Paragraph(listaTextos(recorrerListaTexto), New Font(FontStyle.Regular, TamañoLetra, FontStyle.Bold)))
                            recorrerListaTexto += 1
                            listado = 1
                        ElseIf listatipo(recorrerListaTexto) = "subrayada" Then
                            Console.WriteLine("Estamos en el if de sub")
                            DocumentoResult.Add(New Paragraph(listaTextos(recorrerListaTexto), New Font(FontStyle.Regular, TamañoLetra, FontStyle.Strikeout)))
                            recorrerListaTexto += 1
                            listado = 1
                        ElseIf listatipo(recorrerListaTexto) = "listado" Then
                            DocumentoResult.Add(New Paragraph(listado & ". " & listaTextos(recorrerListaTexto), New Font(FontStyle.Regular, TamañoLetra)))
                            recorrerListaTexto += 1
                            listado += 1
                        ElseIf listatipo(recorrerListaTexto) = "imagen" Then
                            Try
                                Dim imagen As Image = Image.GetInstance(listaTextos(recorrerListaTexto))
                                Console.WriteLine("valx" & valx(imagenesColocadas))
                                Console.WriteLine("valy" & valy(imagenesColocadas))
                                imagen.ScaleAbsolute(CInt(valx(imagenesColocadas)), CInt(valy(imagenesColocadas)))
                                DocumentoResult.Add(imagen)
                                imagenesColocadas += 1
                                listado = 1
                            Catch ex As Exception
                                MessageBox.Show("No colocaste bien la dirección de imagén", "Error con imagén")
                            End Try
                            recorrerListaTexto += 1
                        ElseIf listatipo(recorrerListaTexto) = "variable" Or listatipo(recorrerListaTexto) = "Promedio" Or listatipo(recorrerListaTexto) = "suma" Or listatipo(recorrerListaTexto) = "mult" Or listatipo(recorrerListaTexto) = "rest" Or listatipo(recorrerListaTexto) = "div" Then
                            DocumentoResult.Add(New Phrase(listaTextos(recorrerListaTexto), New Font(FontStyle.Regular, TamañoLetra)))
                            recorrerListaTexto += 1
                        Else
                            recorrerListaTexto += 1
                            listado = 1
                        End If
                    End If
                End While
                DocumentoResult.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            MessageBox.Show("Tienes Errores Léxicos y no se ha podido generar el Archivo", "No se puede Generar el Archivo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End If
    End Sub
End Class
