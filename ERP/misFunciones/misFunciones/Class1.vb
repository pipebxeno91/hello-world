Imports System
Imports ExcelDna.Integration
Imports ExcelDna.ComInterop
Imports System.Runtime.InteropServices

<ComVisible(True)>
<ClassInterface(ClassInterfaceType.AutoDispatch)>
<ProgId("excelymas")>
Public Class Class1
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
End Class

Public Class AddIn
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ComServer.DllRegisterServer()
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        ComServer.DllRegisterServer()
    End Sub
End Class
