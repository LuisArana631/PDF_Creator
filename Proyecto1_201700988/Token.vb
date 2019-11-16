Public Class Token
    'Asignando numero de token a todas las herramientas'
    Enum Tipo
        'Tokens de Palabras Reservadas
        INSTRUCCIONES
        VARIABLES
        TEXTO
        Interlineado
        Tamanio_letra
        Nombre_archivo
        Direccion_archivo
        Entero
        Cadena
        Imagen
        Numeros
        Linea_en_blanco
        Var
        Promedio
        Suma
        Resta
        Multiplicar
        Division
        Asignar

        'Tokens de Signos
        COMA
        IGUAL
        DOS_PUNTOS
        PUNTO_COMA
        LLAVE_IZQ
        LLAVE_DER
        PARENTESIS_IZQ
        PARENTESIS_DER
        CORCHETE_IZQ
        CORCHETE_DER

        'Tokens extras
        Numero_entero
        Numero_decimal
        Identificador
        Comentario
        Texto_normal
        Texto_negrita
        Texto_subrayado
        Ultimo

        'Para Guardar los errores
        ErrorLex

    End Enum
    'Variables para auxiliar en el manejo de tokens'
    Private TipoToken As Tipo
    Private Valor As String

    Public Sub New(ByVal tipo As Tipo, ByVal auxlexico As String)
        Me.TipoToken = tipo
        Me.Valor = auxlexico
    End Sub

    Public Function getTipo() As Tipo
        Return TipoToken
    End Function

    Public Function getValor() As String
        Return Valor
    End Function

    Public Function getTipoString() As String
        Select Case TipoToken
            Case Tipo.INSTRUCCIONES
                Return "Palabra Reservada"
            Case Tipo.VARIABLES
                Return "Palabra Reservada"
            Case Tipo.TEXTO
                Return "Palabra Reservada"
            Case Tipo.Cadena
                Return "Palabra Reservada"
            Case Tipo.Entero
                Return "Palabra Reservada"
            Case Tipo.Interlineado
                Return "Palabra Reservada"
            Case Tipo.Tamanio_letra
                Return "Palabra Reservada"
            Case Tipo.Nombre_archivo
                Return "Palabra Reservada"
            Case Tipo.Direccion_archivo
                Return "Palabra Reservada"
            Case Tipo.Imagen
                Return "Palabra Reservada"
            Case Tipo.Texto_negrita
                Return "Texto en Negrita"
            Case Tipo.Texto_subrayado
                Return "Texto Subrayado"
            Case Tipo.Numeros
                Return "Palabra Reservada"
            Case Tipo.Linea_en_blanco
                Return "Palabra Reservada"
            Case Tipo.Var
                Return "Palabra Reservada"
            Case Tipo.Promedio
                Return "Palabra Reservada"
            Case Tipo.Suma
                Return "Palabra Reservada"
            Case Tipo.Resta
                Return "Palabra Reservada"
            Case Tipo.Multiplicar
                Return "Palabra Reservada"
            Case Tipo.Division
                Return "Palabra Reservada"
            Case Tipo.Asignar
                Return "Palabra Reservada"
            Case Tipo.Comentario
                Return "Comentario"
            Case Tipo.Texto_normal
                Return "Texto  Normal"
            Case Tipo.LLAVE_IZQ
                Return "Llave Izquierda"
            Case Tipo.LLAVE_DER
                Return "Llave Derecha"
            Case Tipo.Identificador
                Return "Identificador"
            Case Tipo.Numero_decimal
                Return "Número Decimal"
            Case Tipo.Numero_entero
                Return "Número Entero"
            Case Tipo.PARENTESIS_IZQ
                Return "Paréntesis Izquierdo"
            Case Tipo.PARENTESIS_DER
                Return "Paréntesis Derecho"
            Case Tipo.CORCHETE_IZQ
                Return "Corchete Izquierdo"
            Case Tipo.CORCHETE_DER
                Return "Corchete Derecho"
            Case Tipo.COMA
                Return "Coma"
            Case Tipo.PUNTO_COMA
                Return "Punto y Coma"
            Case Tipo.DOS_PUNTOS
                Return "Dos Puntos"
            Case Tipo.IGUAL
                Return "Signo Igual"
            Case Tipo.ErrorLex
                Return "Error Léxico"
            Case Tipo.Ultimo
                Return "Ultimo"
            Case Else
                Return "Desconocido"
        End Select
    End Function

End Class
