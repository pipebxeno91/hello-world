Imports ExcelDna.Integration
Public Module Module1
    <ExcelFunction(Description:="Mi primera función UDF en .NET")>
    Public Function DimeHola(nombre As String) As String
        Return "Hola " & nombre
    End Function

    <ExcelFunction(Description:="Función para convertir números a letras")>
    Public Function ConvertirNumerosALetras(num As Integer) As String
        Dim texto As String
        Dim nUnidades As Integer
        Dim cUnidades() As String = {"cero", "uno", "dos", "tres", "cuatro", "cinco",
                                        "seis", "siete", "ocho", "nueve"}
        nUnidades = num

        texto = cUnidades(nUnidades)

        ConvertirNumerosALetras = texto

    End Function
End Module
