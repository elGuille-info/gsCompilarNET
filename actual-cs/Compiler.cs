//
// Clase para compilar el código indicado en un fichero
// Admite código de C# y Visual Basic .NET tanto para consola como para Windows Forms.
//  Si la extensión es .vb se considera de Visual Basic, si no, será de C#.
//
// El código de Windows Forms debe estar completo en el fichero.
//
// Por el código original: Copyright (c) 2019 Laurent Kempé
//

using System;
using System.IO;
using System.Linq;
using Microsoft.CodeAnalysis;
using csc = Microsoft.CodeAnalysis.CSharp;
using vbc = Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.Text;
using System.Collections.Generic;

namespace gsCompilarCore
{
    internal class Compiler
    {
        /// <summary>
        /// Una colección con los fallos al compilar o nulo si no hubo error al compilar.
        /// </summary>
        internal static IEnumerable<Diagnostic> FallosCompilar;

        /// <summary>
        /// Devuelve true si el código a compilar contiene InitializeComponent()
        /// </summary>
        internal bool EsWinForm { get; private set; }

        /// <summary>
        /// Compila el fichero indicado y devuelve un array de bytes con el ensamblado compilado
        /// o nulo si hubo error.
        /// Si hubo error los fallo estarán en <see cref="FallosCompilar"/>.
        /// </summary>
        internal byte[] Compile(string filepath)
        {

            string sourceCode;
            using (var sr = new StreamReader(filepath, System.Text.Encoding.UTF8, true))
            {
                sourceCode = sr.ReadToEnd();
            }

            EsWinForm = sourceCode.IndexOf("InitializeComponent()") > -1;

            // Solo debe ser el nombre, sin path
            var outputExe = Path.GetFileNameWithoutExtension(filepath) + ".dll";
            var extension = Path.GetExtension(filepath).ToLowerInvariant();

            using (var peStream = new MemoryStream())
            {
                Microsoft.CodeAnalysis.Emit.EmitResult result;

                if (extension == ".vb")
                    result = VBGenerateCode(sourceCode, outputExe).Emit(peStream);
                else
                    result = CSGenerateCode(sourceCode, outputExe).Emit(peStream);

                FallosCompilar = null;

                if (!result.Success)
                {
                    var failures = result.Diagnostics.Where(diagnostic => diagnostic.IsWarningAsError || diagnostic.Severity == DiagnosticSeverity.Error);

                    FallosCompilar = failures;

                    //foreach (var diagnostic in failures)
                    //{
                    //    var lin = diagnostic.Location.GetLineSpan().StartLinePosition.Line + 1;
                    //    var pos = diagnostic.Location.GetLineSpan().StartLinePosition.Character + 1;
                    //    Console.WriteLine("{0}: {1} (en línea {2}, posición {3})",
                    //                      diagnostic.Id,
                    //                      diagnostic.GetMessage(),
                    //                      lin, pos);
                    //}

                    return null;
                }

                peStream.Seek(0, SeekOrigin.Begin);

                return peStream.ToArray();
            }
        }

        /// <summary>
        /// Compila el fichero indicado y devuelve el nombre de la DLL generada.
        /// Si hay error devuelve nulo. 
        /// Si hubo error los fallo estarán en <see cref="FallosCompilar"/>.
        /// </summary>
        internal string CompileAsFile(string filepath)
        {
            string sourceCode;
            using (var sr = new StreamReader(filepath, System.Text.Encoding.UTF8, true))
            {
                sourceCode = sr.ReadToEnd();
            }

            EsWinForm = sourceCode.IndexOf("InitializeComponent()") > -1;

            var outputExe = Path.GetFileNameWithoutExtension(filepath) + ".dll";
            var outputPath = Path.Combine(Path.GetDirectoryName(filepath), outputExe);
            var extension = Path.GetExtension(filepath).ToLowerInvariant();

            Microsoft.CodeAnalysis.Emit.EmitResult result;

            if (extension == ".vb")
                result = VBGenerateCode(sourceCode, outputExe).Emit(outputPath);
            else
                result = CSGenerateCode(sourceCode, outputExe).Emit(outputPath);

            FallosCompilar = null;

            if (!result.Success)
            {
                var failures = result.Diagnostics.Where(
                    diagnostic => diagnostic.IsWarningAsError || diagnostic.Severity == DiagnosticSeverity.Error);

                FallosCompilar = failures;

                //foreach (var diagnostic in failures)
                //{
                //    var lin = diagnostic.Location.GetLineSpan().StartLinePosition.Line + 1;
                //    var pos = diagnostic.Location.GetLineSpan().StartLinePosition.Character + 1;
                //    Console.WriteLine("{0}: {1} (en línea {2}, posición {3})",
                //                      diagnostic.Id,
                //                      diagnostic.GetMessage(),
                //                      lin, pos);
                //}

                return null;
            }
            return outputPath;
        }

        private static List<MetadataReference> ColReferencias;

        private static List<MetadataReference> Referencias()
        {
            if (ColReferencias != null)
                return ColReferencias;

            ColReferencias = new List<MetadataReference>();

            var dirCore = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory();
            ColReferencias = ReferenciasDir(dirCore);

            // Para las aplicaciones de Windows Forms

            // Buscar la versión mayor del directorio de aplicaciones de escritorio
            string dirWinDesk = WindowsDesktopApp().Dir;
            ColReferencias.AddRange(ReferenciasDir(dirWinDesk));

            return ColReferencias;
        }

        private static List<MetadataReference> ReferenciasDir(string dirCore)
        {
            var col = new List<MetadataReference>();
            var dll = new List<string>();

            dll.AddRange(Directory.GetFiles(dirCore, "System*.dll"));
            dll.AddRange(Directory.GetFiles(dirCore, "Microsoft*.dll"));

            var noInc = Path.Combine(dirCore, "Microsoft.DiaSymReader.Native.amd64.dll");
            if (dll.Contains(noInc))
                dll.Remove(noInc);

            for (var i = 0; i < dll.Count; i++)
            {
                col.Add(MetadataReference.CreateFromFile(dll[i]));
            }

            return col;
        }

        /// <summary>
        /// Devuelve el directorio y la versión mayor
        /// del path con las DLL de Microsoft.WindowsDesktop.App.
        /// </summary>
        /// <remarks>08/Sep/2020</remarks>
        internal static (string Dir, string Version) WindowsDesktopApp()
        {
            var dirCore = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory();
            // Buscar la versión mayor del directorio de aplicaciones de escritorio
            string dirWinDesk;
            string mayor;
            var dirSep = Path.DirectorySeparatorChar;
            var j = dirCore.IndexOf($"dotnet{dirSep}shared{dirSep}");
            if (j == -1)
            {
                mayor = "5.0.0-preview.8.20411.6";
                dirWinDesk = @"C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App\5.0.0-preview.8.20411.6";
            }
            else
            {
                j += ($"dotnet{dirSep}shared{dirSep}").Length;
                dirWinDesk = Path.Combine(dirCore.Substring(0, j), "Microsoft.WindowsDesktop.App");
                var dirs = Directory.GetDirectories(dirWinDesk).ToList();
                dirs.Sort();
                mayor = Path.GetFileName(dirs.Last());
                dirWinDesk = Path.Combine(dirWinDesk, mayor);
            }
            return (dirWinDesk, mayor);
        }

        /// <summary>
        /// Devuelve el directorio y la versión mayor 
        /// del path con las DLL para aplicaciones NETCore.App.
        /// </summary>
        /// <remarks>08/Sep/2020</remarks>
        internal static (string Dir, string Version) NETCoreApp()
        {
            var dirCore = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory();
            string mayor;
            var j = dirCore.IndexOf("Microsoft.NETCore.App");
            if (j == -1)
            {
                mayor = "5.0.0-preview.8.20407.11";
            }
            else
            {
                j += ("Microsoft.NETCore.App").Length;
                var dirCoreApp = dirCore.Substring(0, j);
                var dirs = Directory.GetDirectories(dirCoreApp).ToList();
                dirs.Sort();
                mayor = Path.GetFileName(dirs.Last());
            }
            return (dirCore, mayor);
        }

        private vbc.VisualBasicCompilation VBGenerateCode(string sourceCode, string outputExe)
        {
            
            var codeString = SourceText.From(sourceCode);
            var options = vbc.VisualBasicParseOptions.Default.WithLanguageVersion(vbc.LanguageVersion.Latest);

            var parsedSyntaxTree = vbc.SyntaxFactory.ParseSyntaxTree(codeString, options);

            // Añadir todas las referencias
            var references = Referencias().ToArray();

            var outputKind = OutputKind.ConsoleApplication;
            if (EsWinForm)
                outputKind = OutputKind.WindowsApplication;

            return vbc.VisualBasicCompilation.Create(outputExe,
                new[] { parsedSyntaxTree },
                references: references,
                options: new vbc.VisualBasicCompilationOptions(outputKind,
                    optimizationLevel: OptimizationLevel.Release,
                    assemblyIdentityComparer: DesktopAssemblyIdentityComparer.Default));
        }

        private csc.CSharpCompilation CSGenerateCode(string sourceCode, string outputExe)
        {
            var codeString = SourceText.From(sourceCode);
            var options = csc.CSharpParseOptions.Default.WithLanguageVersion(csc.LanguageVersion.Latest);

            var parsedSyntaxTree = csc.SyntaxFactory.ParseSyntaxTree(codeString, options);

            // Añadir todas las referencias
            var references = Referencias().ToArray();

            var outputKind = OutputKind.ConsoleApplication;
            if (EsWinForm)
                outputKind = OutputKind.WindowsApplication;

            return csc.CSharpCompilation.Create(outputExe,
                new[] { parsedSyntaxTree },
                references: references,
                options: new csc.CSharpCompilationOptions(outputKind,
                    optimizationLevel: OptimizationLevel.Release,
                    assemblyIdentityComparer: DesktopAssemblyIdentityComparer.Default));
        }
    }
}