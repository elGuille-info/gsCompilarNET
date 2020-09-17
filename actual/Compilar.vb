'------------------------------------------------------------------------------
' Clase para compilar código de VB y C#                             (14/Sep/20)
'
' Usando Microsoft.CodeAnalysis.VisualBasic y Microsoft.CodeAnalysis.CSharp
'
' Código basado en las clases de C# Copyright (c) 2019 Laurent Kempé
' https://laurentkempe.com/2019/02/18/dynamically-compile-and-run-code-using-dotNET-Core-3.0/
'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Text


Public Class Compilar

    ''' <summary>
    ''' Una colección con los fallos al compilar o nulo si no hubo error al compilar.
    ''' </summary>
    Public Shared Function FallosCompilar() As IEnumerable(Of Microsoft.CodeAnalysis.Diagnostic)
        Return Compiler.FallosCompilar
    End Function

    ''' <summary>
    ''' Devuelve true si el código a compilar contiene InitializeComponent()
    ''' </summary>
    Public Shared Property EsWinForm As Boolean

    ''' <summary>
    ''' Compila, guarda y ejecuta el contenido del fichero indicado
    ''' (aplicaciones de consola o de Windows Forms).
    ''' Devuelve una cadena vacía si hubo error.
    ''' </summary>
    ''' <remarks>Si run es false no lo ejecuta.</remarks>
    Public Shared Function CompilarRun(ByVal file As String, ByVal run As Boolean) As String
        Return CompilarGuardar(file, run)
    End Function

    ''' <summary>
    ''' Compila, guarda y ejecuta el contenido del fichero indicado
    ''' (aplicaciones de consola o de Windows Forms).
    ''' Devuelve una cadena vacía si hubo error.
    ''' </summary>
    ''' <remarks>Si run es false no lo ejecuta.</remarks>
    Public Shared Function CompilarRunWinF(ByVal file As String, ByVal run As Boolean) As String
        Return CompilarGuardar(file, run)
    End Function

    ''' <summary>
    ''' Compilar y ejecutar el contenido del fichero indicado
    ''' (solo aplicaciones de consola).
    ''' Devuelve una cadena vacía si todo fue bien.
    ''' </summary>
    Public Shared Function CompilarRunArgs(ByVal file As String, ByVal args As String()) As String
        Return CompilarEjecutar(file, args)
    End Function

    ''' <summary>
    ''' Compilar y ejecutar el contenido del fichero indicado
    ''' (solo aplicaciones de consola).
    ''' Devuelve una cadena vacía si todo fue bien.
    ''' </summary>
    Public Shared Function CompilarRunConsole(ByVal file As String, ByVal args As String()) As String
        Return CompilarEjecutar(file, args)
    End Function

    ''' <summary>
    ''' Compila y ejecuta el código del fichero indicado.
    ''' No utiliza process ni nada externo.
    ''' Tampoco crea la DLL para reutilizarla.
    ''' Devuelve una cadena vacía si todo fue bien.
    ''' </summary>
    ''' <remarks>No usar para aplicaciones de Windows Forms, solo de consola.</remarks>
    Public Shared Function CompilarEjecutar(ByVal file As String, ByVal Optional args As String() = Nothing) As String
        Dim compiler = New Compiler()
        Dim runner = New Runner()

        Dim filepath = file

        Dim res = compiler.Compile(filepath)
        EsWinForm = compiler.EsWinForm

        If res Is Nothing Then
            Return "ERROR al Compilar."
        ElseIf compiler.EsWinForm Then
            Return "No se puede usar CompilarEjecutar para aplicaciones de Windows. Usa CompilarGuardar."
            'Return CompilarGuardar(file, True)
        Else
            Return runner.Execute(res, Nothing)
        End If

        'Return ""
    End Function

    ''' <summary>
    ''' Compila el código del fichero indicado.
    ''' Crea el ensamblado y lo guarda en el directorio del código fuente.
    ''' Crea el fichero runtimeconfig.json
    ''' Devuelve el nombre del ejecutable para usar con "dotnet nombreExe".
    ''' Ejecuta (si así se indica) el ensamblado generado.
    ''' Devuelve una cadena vacía si hubo error.
    ''' </summary>
    Public Shared Function CompilarGuardar(ByVal file As String, ByVal Optional run As Boolean = True) As String
        Dim compiler = New Compiler()

        Dim outputExe = compiler.CompileAsFile(file)
        EsWinForm = compiler.EsWinForm
        If outputExe Is Nothing Then Return Nothing

        ' para ejecutar una DLL usando dotnet, necesitamos un fichero de configuración
        Dim jsonFile = Path.ChangeExtension(outputExe, "runtimeconfig.json")

        Dim jsonText = ""

        If compiler.EsWinForm Then
            Dim version = Compiler.WindowsDesktopApp().Version
            ' Aplicación de escritorio (Windows Forms)
            ' Microsoft.WindowsDesktop.App
            ' 5.0.0-preview.8.20411.6
            jsonText = "
{
    ""runtimeOptions"": {
    ""tfm"": ""net5.0-windows"",
    ""framework"": {
        ""name"": ""Microsoft.WindowsDesktop.App"",
        ""version"": """ & version & """
    }
    }
}"
        Else
            Dim version = Compiler.NETCoreApp().Version
            ' Tipo consola
            ' Microsoft.NETCore.App
            ' 5.0.0-preview.8.20407.11
            jsonText = "
{
    ""runtimeOptions"": {
    ""tfm"": ""net5.0"",
    ""framework"": {
        ""name"": ""Microsoft.NETCore.App"",
        ""version"": """ & version & """
    }
    }
}"
        End If

        Using sw = New StreamWriter(jsonFile, False, Encoding.UTF8)
            sw.WriteLine(jsonText)
        End Using

        If run Then

            Try
                ' Algunas veces no se ejecuta,                      (17/Sep/20)
                ' porque el path contiene espacios.
                Process.Start("dotnet", $"{ChrW(34)}{outputExe}{ChrW(34)}")
                'Process.Start("dotnet", outputExe)

            Catch
            End Try
        End If

        Return outputExe
    End Function

    ''' <summary>
    ''' Devuelve el directorio y la versión mayor 
    ''' del path con las DLL para aplicaciones NETCore.App.
    ''' </summary>
    ''' <remarks>08/Sep/2020</remarks>
    Public Shared Function NETCoreApp() As (Dir As String, Version As String)
        Return Compiler.NETCoreApp()
    End Function

    ''' <summary>
    ''' Devuelve el directorio y la versión mayor
    ''' del path con las DLL de Microsoft.WindowsDesktop.App.
    ''' </summary>
    ''' <remarks>08/Sep/2020</remarks>
    Public Shared Function WindowsDesktopApp() As (Dir As String, Version As String)
        Return Compiler.WindowsDesktopApp()
    End Function

    ''' <summary>
    ''' Devuelve la versión de la DLL.
    ''' Si completa es True, se devuelve también el nombre de la DLL:
    ''' gsCompilarCore v 1.0.0.0 (para .NET Core 3.1 revisión del dd/MMM/yyyy)
    ''' </summary>
    Public Shared Function Version(ByVal Optional completa As Boolean = False) As String
        Dim ensamblado = System.Reflection.Assembly.GetExecutingAssembly()

        Dim versionAttr = ensamblado.GetCustomAttributes(GetType(System.Reflection.AssemblyVersionAttribute), False)
        Dim vers = If(versionAttr.Length > 0,
                        (TryCast(versionAttr(0), System.Reflection.AssemblyVersionAttribute)).Version,
                        "1.0.0.0")

        Dim fileVerAttr = ensamblado.GetCustomAttributes(GetType(System.Reflection.AssemblyFileVersionAttribute), False)
        Dim versF = If(fileVerAttr.Length > 0,
                        (TryCast(fileVerAttr(0), System.Reflection.AssemblyFileVersionAttribute)).Version,
                        "1.0.0.1")

        Dim res = $"v {vers} ({versF})"

        If completa Then
            Dim prodAttr = ensamblado.GetCustomAttributes(GetType(System.Reflection.AssemblyProductAttribute), False)
            Dim producto = If(prodAttr.Length > 0,
                                (TryCast(prodAttr(0), System.Reflection.AssemblyProductAttribute)).Product,
                                "gsCompilarNET")

            Dim descAttr = ensamblado.GetCustomAttributes(GetType(System.Reflection.AssemblyDescriptionAttribute), False)
            Dim desc = If(descAttr.Length > 0,
                                (TryCast(descAttr(0), System.Reflection.AssemblyDescriptionAttribute)).Description,
                                "(para .NET Core 3.1 revisión del 14/Sep/2020)")

            desc = desc.Substring(desc.IndexOf("(para .NET"))

            res = $"{producto} {res} {desc}"
        End If

        Return res
    End Function

End Class
