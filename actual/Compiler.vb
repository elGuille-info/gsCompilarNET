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

Imports Microsoft.VisualBasic

Imports System
Imports System.IO
Imports System.Linq
Imports Microsoft.CodeAnalysis
Imports csc = Microsoft.CodeAnalysis.CSharp
Imports vbc = Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.Text
Imports System.Collections.Generic

Friend Class Compiler

    ''' <summary>
    ''' Una colección con los fallos al compilar o nulo si no hubo error al compilar.
    ''' </summary>
    Friend Shared FallosCompilar As IEnumerable(Of Diagnostic)

    ''' <summary>
    ''' Devuelve true si el código a compilar contiene InitializeComponent()
    ''' </summary>
    Friend Property EsWinForm As Boolean

    ''' <summary>
    ''' Compila el fichero indicado y devuelve un array de bytes con el ensamblado compilado
    ''' o nulo si hubo error. 
    ''' Si hubo error los fallo estarán en <see cref="FallosCompilar"/>.
    ''' </summary>
    Friend Function Compile(ByVal filepath As String) As Byte()
        Dim sourceCode As String

        Using sr = New StreamReader(filepath, System.Text.Encoding.UTF8, True)
            sourceCode = sr.ReadToEnd()
        End Using

        EsWinForm = sourceCode.IndexOf("InitializeComponent()") > -1

        ' Solo debe ser el nombre, sin path
        Dim outputExe = Path.GetFileNameWithoutExtension(filepath) & ".dll"
        Dim extension = Path.GetExtension(filepath).ToLowerInvariant()

        Using peStream = New MemoryStream()
            Dim result As Microsoft.CodeAnalysis.Emit.EmitResult

            If extension = ".vb" Then
                result = VBGenerateCode(sourceCode, outputExe).Emit(peStream)
            Else
                result = CSGenerateCode(sourceCode, outputExe).Emit(peStream)
            End If

            FallosCompilar = Nothing

            If Not result.Success Then
                Dim failures = result.Diagnostics.Where(
                    Function(diagnostic) diagnostic.IsWarningAsError OrElse diagnostic.Severity = DiagnosticSeverity.[Error])

                FallosCompilar = failures

                'For Each diagnostic In failures
                '    Dim lin = diagnostic.Location.GetLineSpan().StartLinePosition.Line + 1
                '    Dim pos = diagnostic.Location.GetLineSpan().StartLinePosition.Character + 1
                '    Console.WriteLine("{0}: {1} (en línea {2}, posición {3})",
                '                      diagnostic.Id,
                '                      diagnostic.GetMessage(),
                '                      lin, pos)
                'Next

                Return Nothing
            End If

            peStream.Seek(0, SeekOrigin.Begin)

            Return peStream.ToArray()
        End Using
    End Function

    ''' <summary>
    ''' Compila el fichero indicado y devuelve el nombre de la DLL generada.
    ''' Si hay error devuelve nulo. 
    ''' Si hubo error los fallo estarán en <see cref="FallosCompilar"/>.
    ''' </summary>
    Friend Function CompileAsFile(ByVal filepath As String) As String
        Dim sourceCode As String

        Using sr = New StreamReader(filepath, System.Text.Encoding.UTF8, True)
            sourceCode = sr.ReadToEnd()
        End Using

        EsWinForm = sourceCode.IndexOf("InitializeComponent()") > -1

        Dim outputExe = Path.GetFileNameWithoutExtension(filepath) & ".dll"
        Dim outputPath = Path.Combine(Path.GetDirectoryName(filepath), outputExe)
        Dim extension = Path.GetExtension(filepath).ToLowerInvariant()

        Dim result As Microsoft.CodeAnalysis.Emit.EmitResult

        If extension = ".vb" Then
            result = VBGenerateCode(sourceCode, outputExe).Emit(outputPath)
        Else
            result = CSGenerateCode(sourceCode, outputExe).Emit(outputPath)
        End If

        FallosCompilar = Nothing

        If Not result.Success Then

            Dim failures = result.Diagnostics.Where(
                Function(diagnostic) diagnostic.IsWarningAsError OrElse diagnostic.Severity = DiagnosticSeverity.[Error])

            FallosCompilar = failures

            'For Each diagnostic In failures
            '    Dim lin = diagnostic.Location.GetLineSpan().StartLinePosition.Line + 1
            '    Dim pos = diagnostic.Location.GetLineSpan().StartLinePosition.Character + 1
            '    Console.WriteLine("{0}: {1} (en línea {2}, posición {3})",
            '                      diagnostic.Id,
            '                      diagnostic.GetMessage(),
            '                      lin, pos)
            'Next

            Return Nothing
        End If

        Return outputPath
    End Function

    Private Shared ColReferencias As List(Of MetadataReference)

    Private Shared Function Referencias() As List(Of MetadataReference)
        If ColReferencias IsNot Nothing Then Return ColReferencias

        ColReferencias = New List(Of MetadataReference)

        Dim dirCore = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory()
        ColReferencias = ReferenciasDir(dirCore)

        ' Para las aplicaciones de Windows Forms

        ' Buscar la versión mayor del directorio de aplicaciones de escritorio
        Dim dirWinDesk As String = WindowsDesktopApp().Dir
        ColReferencias.AddRange(ReferenciasDir(dirWinDesk))

        Return ColReferencias
    End Function

    Private Shared Function ReferenciasDir(ByVal dirCore As String) As List(Of MetadataReference)
        Dim col = New List(Of MetadataReference)()
        Dim dll = New List(Of String)()

        dll.AddRange(Directory.GetFiles(dirCore, "System*.dll"))
        dll.AddRange(Directory.GetFiles(dirCore, "Microsoft*.dll"))

        Dim noInc = Path.Combine(dirCore, "Microsoft.DiaSymReader.Native.amd64.dll")
        If dll.Contains(noInc) Then dll.Remove(noInc)

        For i = 0 To dll.Count - 1
            col.Add(MetadataReference.CreateFromFile(dll(i)))
        Next

        Return col
    End Function

    ''' <summary>
    ''' Devuelve el directorio y la versión mayor
    ''' del path con las DLL de Microsoft.WindowsDesktop.App.
    ''' </summary>
    ''' <remarks>08/Sep/2020</remarks>
    Friend Shared Function WindowsDesktopApp() As (Dir As String, Version As String)
        Dim dirCore = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory()
        ' Buscar la versión mayor del directorio de aplicaciones de escritorio
        Dim dirWinDesk As String
        Dim mayor As String
        Dim dirSep = Path.DirectorySeparatorChar
        Dim j = dirCore.IndexOf($"dotnet{dirSep}shared{dirSep}")

        If j = -1 Then
            mayor = "5.0.0-preview.8.20411.6"
            dirWinDesk = "C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App\5.0.0-preview.8.20411.6"
        Else
            j += ($"dotnet{dirSep}shared{dirSep}").Length
            dirWinDesk = Path.Combine(dirCore.Substring(0, j), "Microsoft.WindowsDesktop.App")
            Dim dirs = Directory.GetDirectories(dirWinDesk).ToList()
            dirs.Sort()
            mayor = Path.GetFileName(dirs.Last())
            dirWinDesk = Path.Combine(dirWinDesk, mayor)
        End If

        Return (dirWinDesk, mayor)
    End Function

    ''' <summary>
    ''' Devuelve el directorio y la versión mayor 
    ''' del path con las DLL para aplicaciones NETCore.App.
    ''' </summary>
    ''' <remarks>08/Sep/2020</remarks>
    Friend Shared Function NETCoreApp() As (Dir As String, Version As String)
        Dim dirCore = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory()
        Dim mayor As String
        Dim j = dirCore.IndexOf("Microsoft.NETCore.App")

        If j = -1 Then
            mayor = "5.0.0-preview.8.20407.11"
        Else
            j += ("Microsoft.NETCore.App").Length
            Dim dirCoreApp = dirCore.Substring(0, j)
            Dim dirs = Directory.GetDirectories(dirCoreApp).ToList()
            dirs.Sort()
            mayor = Path.GetFileName(dirs.Last())
        End If

        Return (dirCore, mayor)
    End Function

    Private Function VBGenerateCode(ByVal sourceCode As String, ByVal outputExe As String) As vbc.VisualBasicCompilation
        Dim codeString = SourceText.From(sourceCode)
        Dim options = vbc.VisualBasicParseOptions.[Default].WithLanguageVersion(vbc.LanguageVersion.Latest)

        Dim parsedSyntaxTree = vbc.SyntaxFactory.ParseSyntaxTree(codeString, options)

        ' Añadir todas las referencias
        Dim references = Referencias().ToArray()

        Dim outpKind = OutputKind.ConsoleApplication
        If EsWinForm Then _
            outpKind = OutputKind.WindowsApplication

        Return vbc.VisualBasicCompilation.Create(
                outputExe,
                {parsedSyntaxTree},
                references:=references,
                options:=New vbc.VisualBasicCompilationOptions(
                    outpKind,
                    optimizationLevel:=OptimizationLevel.Release,
                    assemblyIdentityComparer:=DesktopAssemblyIdentityComparer.[Default]))

    End Function

    Private Function CSGenerateCode(ByVal sourceCode As String, ByVal outputExe As String) As csc.CSharpCompilation
        Dim codeString = SourceText.From(sourceCode)
        Dim options = csc.CSharpParseOptions.[Default].WithLanguageVersion(csc.LanguageVersion.Latest)

        Dim parsedSyntaxTree = csc.SyntaxFactory.ParseSyntaxTree(codeString, options)

        ' Añadir todas las referencias
        Dim references = Referencias().ToArray()

        Dim outpKind = Microsoft.CodeAnalysis.OutputKind.ConsoleApplication
        If EsWinForm Then _
            outpKind = OutputKind.WindowsApplication

        Return csc.CSharpCompilation.Create(
                outputExe,
                {parsedSyntaxTree},
                references:=references,
                options:=New csc.CSharpCompilationOptions(
                    outpKind,
                    optimizationLevel:=OptimizationLevel.Release,
                    assemblyIdentityComparer:=DesktopAssemblyIdentityComparer.[Default]))

    End Function

End Class
