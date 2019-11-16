Imports Proyecto1_201700988.Token

Public Class Form1

    Dim RutaArchivo As String = ""

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub AbrirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AbrirToolStripMenuItem.Click
        Cargar_Archivo()
    End Sub

    Public Sub Cargar_Archivo()
        Try
            'Filtro para buscar archivos especificos
            BuscadorArchivos.Filter = "Archivos de Texto|*.ACK;"
            BuscadorArchivos.FileName = ""
            'Validación al seleccionar archivo
            If BuscadorArchivos.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                'Colocar la ruta en una variable
                RutaArchivo = BuscadorArchivos.FileName

            End If
            'Lector de Archivos
            Dim Lector As String
            Lector = My.Computer.FileSystem.ReadAllText(BuscadorArchivos.FileName, System.Text.Encoding.Default)
            'Lee el archivo obtenido de la ruta anterior
            txt_texto.Text = Lector

        Catch ex As Exception

        End Try
    End Sub

    Public Sub Guardar_Como()
        Try
            Dim TextoAGuardar As String = ""
            TextoAGuardar = ""

            Dim Objeto As Object
            Dim Archivo As Object

            TextoAGuardar = txt_texto.Text

            Dim GuardarArchivo As New SaveFileDialog()
            GuardarArchivo.Filter = "Archivos de Texto (*.ACK)|*.ACK|Todos Los Archivos (*.*)|*.*"

            If GuardarArchivo.ShowDialog() = DialogResult.OK Then
                Objeto = CreateObject("Scripting.FileSystemObject")
                Archivo = Objeto.CreateTextFile(GuardarArchivo.FileName)
                Archivo.write(TextoAGuardar)
                Archivo.close()
                RutaArchivo = GuardarArchivo.FileName
                MessageBox.Show("Se ha creado el Archivo en " + RutaArchivo, "Guardado con éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error al crear el archivo debido a: " + ex.ToString)
        End Try
    End Sub

    Public Sub Guardar()
        Try
            Dim TextoAGuardar As String = ""
            TextoAGuardar = ""

            TextoAGuardar = txt_texto.Text

            If RutaArchivo = "" Then
                Guardar_Como()
            Else
                My.Computer.FileSystem.WriteAllText(RutaArchivo, TextoAGuardar, False)
                MessageBox.Show("Se ha guardado el Archivo en " + RutaArchivo, "Guardado con éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        If MsgBox("¿Deseas cerrar el programa?", vbQuestion + vbYesNo, "Cerrar programa") = vbYes Then
            End
        End If
    End Sub

    Private Sub GuardarComoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GuardarComoToolStripMenuItem.Click
        Guardar_Como()
    End Sub

    Private Sub GuardarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GuardarToolStripMenuItem.Click
        Guardar()
    End Sub

    Private Sub btn_analizar_Click(sender As Object, e As EventArgs) Handles btn_analizar.Click
        'Le mandamos el texto a la varible
        Dim entrada As String = txt_texto.Text
        If entrada = "" Or entrada = " " Then
            MessageBox.Show("Debes escribir texto o cargar un archivo", "No se puede Realizar Análisis", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            'PintarPalabras
            PintarPalabras()
            'Iniciamos el analisis lexico
            Dim AnalisisLex As AnalisisLexico = New AnalisisLexico()
            Dim ListaToken As List(Of Token) = AnalisisLex.Scaner(entrada)
            'Imprimimos en consola la lista
            'AnalisisLex.imprimirLista(ListaToken)
            'Iniciamos el analisis sintactico
            Console.WriteLine("")
            Console.WriteLine("")
            Console.WriteLine("")
            Dim parser As AnalisisSintactico = New AnalisisSintactico
            parser.parsear(ListaToken)
            'Crear Documento de Errores Lexicos
            AnalisisLex.CrearDocumentoErrores(ListaToken)
            'Crear Documento de tokens
            AnalisisLex.CrearDocumentoTokens(ListaToken)
            'Crear Documento de Errores Sintacticos
            parser.CrearDocumentoErrorSintactico()

            If AnalisisLex.CantidadErrores(ListaToken) = 0 Then
                Dim CrearDoc As CrearDocumento = New CrearDocumento()
                CrearDoc.CrearDocumento(ListaToken)
            Else
                MessageBox.Show("Tienes Errores Léxicos y no se ha podido generar el Archivo", "No se puede Generar el Archivo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End If
        End If
    End Sub

    Private Sub AcercaDeToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AcercaDeToolStripMenuItem1.Click
        AcercaDE.Show()
    End Sub

    Private Sub PintarPalabras()
        Dim temp As String = txt_texto.Text
        txt_texto.Clear()
        txt_texto.Text = temp
        Try
            'PALABRAS RESERVADAS
            Dim busqueda As String = "INSTRUCCIONES"
            Dim index As Integer
            While index < txt_texto.Text.LastIndexOf(busqueda)
                txt_texto.Find(busqueda, index, txt_texto.TextLength, RichTextBoxFinds.MatchCase)
                txt_texto.SelectionColor = Color.Blue
                index = txt_texto.Text.IndexOf(busqueda, index) + 1
            End While
            Dim busqueda1 As String = "variables".ToUpper
            Dim index1 As Integer
            While index1 < txt_texto.Text.LastIndexOf(busqueda1)
                txt_texto.Find(busqueda1, index1, txt_texto.TextLength, RichTextBoxFinds.MatchCase)
                txt_texto.SelectionColor = Color.Blue
                index1 = txt_texto.Text.IndexOf(busqueda1, index1) + 1
            End While
            Dim busqueda2 As String = "texto".ToUpper
            Dim index2 As Integer
            While index2 < txt_texto.Text.LastIndexOf(busqueda2)
                txt_texto.Find(busqueda2, index2, txt_texto.TextLength, RichTextBoxFinds.MatchCase)
                txt_texto.SelectionColor = Color.Blue
                index2 = txt_texto.Text.IndexOf(busqueda2, index2) + 1
            End While
            Dim busqueda3 As String = "Interlineado"
            Dim index3 As Integer
            While index3 < txt_texto.Text.LastIndexOf(busqueda3)
                txt_texto.Find(busqueda3, index3, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index3 = txt_texto.Text.IndexOf(busqueda3, index3) + 1
            End While
            Dim busqueda4 As String = "Tamanio_letra"
            Dim index4 As Integer
            While index4 < txt_texto.Text.LastIndexOf(busqueda4)
                txt_texto.Find(busqueda4, index4, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index4 = txt_texto.Text.IndexOf(busqueda4, index4) + 1
            End While
            Dim busqueda5 As String = "Nombre_archivo"
            Dim index5 As Integer
            While index5 < txt_texto.Text.LastIndexOf(busqueda5)
                txt_texto.Find(busqueda5, index5, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index5 = txt_texto.Text.IndexOf(busqueda5, index5) + 1
            End While
            Dim busqueda6 As String = "Direccion_archivo"
            Dim index6 As Integer
            While index6 < txt_texto.Text.LastIndexOf(busqueda6)
                txt_texto.Find(busqueda6, index6, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index6 = txt_texto.Text.IndexOf(busqueda6, index6) + 1
            End While
            Dim busqueda7 As String = "entero"
            Dim index7 As Integer
            While index7 < txt_texto.Text.LastIndexOf(busqueda7)
                txt_texto.Find(busqueda7, index7, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index7 = txt_texto.Text.IndexOf(busqueda7, index7) + 1
            End While
            Dim busqueda8 As String = "cadena"
            Dim index8 As Integer
            While index8 < txt_texto.Text.LastIndexOf(busqueda8)
                txt_texto.Find(busqueda8, index8, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index8 = txt_texto.Text.IndexOf(busqueda8, index8) + 1
            End While
            Dim busqueda9 As String = "Imagen"
            Dim index9 As Integer
            While index9 < txt_texto.Text.LastIndexOf(busqueda9)
                txt_texto.Find(busqueda9, index9, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index9 = txt_texto.Text.IndexOf(busqueda9, index9) + 1
            End While
            Dim busqueda10 As String = "Numeros"
            Dim index10 As Integer
            While index10 < txt_texto.Text.LastIndexOf(busqueda10)
                txt_texto.Find(busqueda10, index10, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index10 = txt_texto.Text.IndexOf(busqueda10, index10) + 1
            End While
            Dim busqueda11 As String = "Linea_en_blanco"
            Dim index11 As Integer
            While index11 < txt_texto.Text.LastIndexOf(busqueda11)
                txt_texto.Find(busqueda11, index11, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index11 = txt_texto.Text.IndexOf(busqueda11, index11) + 1
            End While
            Dim busqueda12 As String = "Var"
            Dim index12 As Integer
            While index12 < txt_texto.Text.LastIndexOf(busqueda12)
                txt_texto.Find(busqueda12, index12, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index12 = txt_texto.Text.IndexOf(busqueda12, index12) + 1
            End While
            Dim busqueda13 As String = "Promedio"
            Dim index13 As Integer
            While index13 < txt_texto.Text.LastIndexOf(busqueda13)
                txt_texto.Find(busqueda13, index13, txt_texto.TextLength, RichTextBoxFinds.MatchCase)
                txt_texto.SelectionColor = Color.Blue
                index13 = txt_texto.Text.IndexOf(busqueda13, index13) + 1
            End While
            Dim busqueda14 As String = "Suma"
            Dim index14 As Integer
            While index14 < txt_texto.Text.LastIndexOf(busqueda14)
                txt_texto.Find(busqueda14, index14, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index14 = txt_texto.Text.IndexOf(busqueda14, index14) + 1
            End While
            Dim busqueda15 As String = "Resta"
            Dim index15 As Integer
            While index15 < txt_texto.Text.LastIndexOf(busqueda15)
                txt_texto.Find(busqueda15, index15, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index15 = txt_texto.Text.IndexOf(busqueda15, index15) + 1
            End While
            Dim busqueda16 As String = "Multiplicar"
            Dim index16 As Integer
            While index16 < txt_texto.Text.LastIndexOf(busqueda16)
                txt_texto.Find(busqueda16, index16, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index16 = txt_texto.Text.IndexOf(busqueda16, index16) + 1
            End While
            Dim busqueda17 As String = "Division"
            Dim index17 As Integer
            While index17 < txt_texto.Text.LastIndexOf(busqueda17)
                txt_texto.Find(busqueda17, index17, txt_texto.TextLength, RichTextBoxFinds.WholeWord)
                txt_texto.SelectionColor = Color.Blue
                index17 = txt_texto.Text.IndexOf(busqueda17, index17) + 1
            End While
            Dim busqueda18 As String = "Asignar"
            Dim index18 As Integer
            While index18 < txt_texto.Text.LastIndexOf(busqueda18)
                txt_texto.Find(busqueda18, index18, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.Blue
                index18 = txt_texto.Text.IndexOf(busqueda18, index18) + 1
            End While
            'SIMBOLOS
            Dim busqueda19 As String = ","
            Dim index19 As Integer
            While index19 < txt_texto.Text.LastIndexOf(busqueda19)
                txt_texto.Find(busqueda19, index19, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.DarkOrange
                index19 = txt_texto.Text.IndexOf(busqueda19, index19) + 1
            End While
            Dim busqueda20 As String = "="
            Dim index20 As Integer
            While index20 < txt_texto.Text.LastIndexOf(busqueda20)
                txt_texto.Find(busqueda20, index20, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.DarkOrange
                index20 = txt_texto.Text.IndexOf(busqueda20, index20) + 1
            End While
            Dim busqueda21 As String = ":"
            Dim index21 As Integer
            While index21 < txt_texto.Text.LastIndexOf(busqueda21)
                txt_texto.Find(busqueda21, index21, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.DarkOrange
                index21 = txt_texto.Text.IndexOf(busqueda21, index21) + 1
            End While
            Dim busqueda22 As String = ";"
            Dim index22 As Integer
            While index22 < txt_texto.Text.LastIndexOf(busqueda22)
                txt_texto.Find(busqueda22, index22, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.DarkOrange
                index22 = txt_texto.Text.IndexOf(busqueda22, index22) + 1
            End While
            Dim busqueda23 As String = "{"
            Dim index23 As Integer
            While index23 < txt_texto.Text.LastIndexOf(busqueda23)
                txt_texto.Find(busqueda23, index23, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.DarkOrange
                index23 = txt_texto.Text.IndexOf(busqueda23, index23) + 1
            End While
            Dim busqueda24 As String = "}"
            Dim index24 As Integer
            While index24 < txt_texto.Text.LastIndexOf(busqueda24)
                txt_texto.Find(busqueda24, index24, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.DarkOrange
                index24 = txt_texto.Text.IndexOf(busqueda24, index24) + 1
            End While
            Dim busqueda25 As String = "("
            Dim index25 As Integer
            While index25 < txt_texto.Text.LastIndexOf(busqueda25)
                txt_texto.Find(busqueda25, index25, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.DarkOrange
                index25 = txt_texto.Text.IndexOf(busqueda25, index25) + 1
            End While
            Dim busqueda26 As String = ")"
            Dim index26 As Integer
            While index26 < txt_texto.Text.LastIndexOf(busqueda26)
                txt_texto.Find(busqueda26, index26, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.DarkOrange
                index26 = txt_texto.Text.IndexOf(busqueda26, index26) + 1
            End While
            Dim busqueda27 As String = "["
            Dim index27 As Integer
            While index27 < txt_texto.Text.LastIndexOf(busqueda27)
                txt_texto.Find(busqueda27, index27, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.DarkOrange
                index27 = txt_texto.Text.IndexOf(busqueda27, index27) + 1
            End While
            Dim busqueda28 As String = "]"
            Dim index28 As Integer
            While index28 < txt_texto.Text.LastIndexOf(busqueda28)
                txt_texto.Find(busqueda28, index28, txt_texto.TextLength, RichTextBoxFinds.None)
                txt_texto.SelectionColor = Color.DarkOrange
                index28 = txt_texto.Text.IndexOf(busqueda28, index28) + 1
            End While
            'NUMEROS
            Dim AnalisisLex As AnalisisLexico = New AnalisisLexico()
            Dim ListaToken As List(Of Token) = AnalisisLex.Scaner(temp)
            For Each numero As Token In ListaToken
                If numero.getTipo = Tipo.Numero_decimal Then
                    Dim index29 As Integer
                    Console.WriteLine("Entramos con: " + numero.getValor)
                    While index29 < txt_texto.Text.LastIndexOf(numero.getValor)
                        txt_texto.Find(numero.getValor, index29, txt_texto.TextLength, RichTextBoxFinds.None)
                        txt_texto.SelectionColor = Color.Yellow
                        index29 = txt_texto.Text.IndexOf(numero.getValor, index29) + 1
                    End While
                End If
            Next
            For Each numero As Token In ListaToken
                If numero.getTipo = Tipo.Numero_entero Then
                    Dim index30 As Integer
                    Console.WriteLine("Entramos con: " + numero.getValor)
                    While index30 < txt_texto.Text.LastIndexOf(numero.getValor)
                        txt_texto.Find(numero.getValor, index30, txt_texto.TextLength, RichTextBoxFinds.None)
                        txt_texto.SelectionColor = Color.Yellow
                        index30 = txt_texto.Text.IndexOf(numero.getValor, index30) + 1
                    End While
                End If
            Next
            'COMENTARIOS
            For Each comentario As Token In ListaToken
                If comentario.getTipo = Tipo.Comentario Then
                    Dim index31 As Integer
                    While index31 < txt_texto.Text.LastIndexOf(comentario.getValor)
                        txt_texto.Find(comentario.getValor, index31, txt_texto.TextLength, RichTextBoxFinds.MatchCase)
                        txt_texto.SelectionColor = Color.Gray
                        index31 = txt_texto.Text.IndexOf(comentario.getValor, index31) + 1
                    End While
                End If
            Next
            'CADENAS
            For Each cadena As Token In ListaToken
                If cadena.getTipo = Tipo.Texto_negrita Then
                    Dim index32 As Integer
                    While index32 < txt_texto.Text.LastIndexOf(cadena.getValor)
                        txt_texto.Find(cadena.getValor, index32, txt_texto.TextLength, RichTextBoxFinds.MatchCase)
                        txt_texto.SelectionColor = Color.Green
                        index32 = txt_texto.Text.IndexOf(cadena.getValor, index32) + 1
                    End While
                ElseIf cadena.getTipo = Tipo.Texto_subrayado Then
                    Dim index33 As Integer
                    While index33 < txt_texto.Text.LastIndexOf(cadena.getValor)
                        txt_texto.Find(cadena.getValor, index33, txt_texto.TextLength, RichTextBoxFinds.MatchCase)
                        txt_texto.SelectionColor = Color.Green
                        index33 = txt_texto.Text.IndexOf(cadena.getValor, index33) + 1
                    End While
                ElseIf cadena.getTipo = Tipo.Texto_normal Then
                    Dim index32 As Integer
                    While index32 < txt_texto.Text.LastIndexOf(cadena.getValor)
                        txt_texto.Find(cadena.getValor, index32, txt_texto.TextLength, RichTextBoxFinds.MatchCase)
                        txt_texto.SelectionColor = Color.Green
                        index32 = txt_texto.Text.IndexOf(cadena.getValor, index32) + 1
                    End While
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub
End Class
