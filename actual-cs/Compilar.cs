//-----------------------------------------------------------------------------
// Compilar y ejecutar código de VB y C#                            (02/Sep/20)
// Biblioteca de clases (.NET Core 3.1) para compilar código de C# y VB
//
// Basado en el código para C# de Laurent Kempé y otros ejemplos de la red.
// https://laurentkempe.com/2019/02/18/dynamically-compile-and-run-code-using-dotNET-Core-3.0/
//
// 
//
// (c) Guillermo (elGuille) Som, 2020
//-----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace gsCompilarCore
{
    public class Compilar
    {
        /// <summary>
        /// Una colección con los fallos al compilar o nulo si no hubo error al compilar.
        /// </summary>
        public static IEnumerable<Microsoft.CodeAnalysis.Diagnostic> FallosCompilar()
        {
            return Compiler.FallosCompilar;
        }
            

        /// <summary>
        /// Devuelve true si el código a compilar contiene InitializeComponent()
        /// </summary>
        public static bool EsWinForm { get; private set; }

        /// <summary>
        /// Compila, guarda y ejecuta el contenido del fichero indicado
        /// (aplicaciones de consola o de Windows Forms).
        /// </summary>
        /// <remarks>Si run es false no lo ejecuta.</remarks>
        public static string CompilarRun(string file, bool run)
        {
            return CompilarGuardar(file, run);
        }

        /// <summary>
        /// Compila, guarda y ejecuta el contenido del fichero indicado
        /// (aplicaciones de consola o de Windows Forms).
        /// </summary>
        /// <remarks>Si run es false no lo ejecuta.</remarks>
        public static string CompilarRunWinF(string file, bool run)
        {
            return CompilarGuardar(file, run);
        }

        /// <summary>
        /// Compilar y ejecutar el contenido del fichero indicado
        /// (solo aplicaciones de consola).
        /// </summary>
        public static string CompilarRunArgs(string file, string[] args)
        {
            return CompilarEjecutar(file, args);
        }

        /// <summary>
        /// Compilar y ejecutar el contenido del fichero indicado
        /// (solo aplicaciones de consola).
        /// </summary>
        public static string CompilarRunConsole(string file, string[] args)
        {
            return CompilarEjecutar(file, args);
        }


        /// <summary>
        /// Compila y ejecuta el código del fichero indicado.
        /// No utiliza process ni nada externo.
        /// Tampoco crea la DLL para reutilizarla.
        /// Devuelve una cadena vacía si todo fue bien.
        /// </summary>
        /// <remarks>No usar para aplicaciones de Windows Forms, solo de consola.</remarks>
        public static string CompilarEjecutar(string file, string[] args = null)
        {
            var compiler = new Compiler();
            var runner = new Runner();

            var filepath = file;
            
            var res = compiler.Compile(filepath);
            EsWinForm = compiler.EsWinForm;
            if (res == null)
            {
                return "ERROR al Compilar.";
            }
            else if (compiler.EsWinForm)
            {
                return "No se puede usar CompilarEjecutar para aplicaciones de Windows. Usa CompilarGuardar.";
                //return CompilarGuardar(file, true);
            }
            else
            {
                runner.Execute(res, null);
            }

            return "";
        }

        /// <summary>
        /// Compila el código del fichero indicado.
        /// Crea el ensamblado y lo guarda en el directorio del código fuente.
        /// Crea el fichero runtimeconfig.json
        /// Devuelve el nombre del ejecutable para usar con "dotnet nombreExe".
        /// Ejecuta (si así se indica) el ensamblado generado.
        /// </summary>
        public static string CompilarGuardar(string file, bool run = true)
        {
            var compiler = new Compiler();

            var outputExe = compiler.CompileAsFile(file);
            EsWinForm = compiler.EsWinForm;
            if (outputExe == null)
                return null;

            // para ejecutar una DLL usando dotnet, necesitamos un fichero de configuración
            var jsonFile = Path.ChangeExtension(outputExe, "runtimeconfig.json");

            var jsonText = "";

            if (compiler.EsWinForm)
            {
                var version = Compiler.WindowsDesktopApp().Version;
                // Aplicación de escritorio (Windows Forms)
                // Microsoft.WindowsDesktop.App
                // 5.0.0-preview.8.20411.6
                jsonText = @"
{
    ""runtimeOptions"": {
    ""tfm"": ""net5.0-windows"",
    ""framework"": {
        ""name"": ""Microsoft.WindowsDesktop.App"",
        ""version"": """ + version + @"""
    }
    }
}";
            }
            else
            {
                var version = Compiler.NETCoreApp().Version;
                // Tipo consola
                // Microsoft.NETCore.App
                // 5.0.0-preview.8.20407.11
                jsonText = @"
{
    ""runtimeOptions"": {
    ""tfm"": ""net5.0"",
    ""framework"": {
        ""name"": ""Microsoft.NETCore.App"",
        ""version"": """ + version + @"""
    }
    }
}";
            }

            using (var sw = new StreamWriter(jsonFile, false, Encoding.UTF8))
            {
                sw.WriteLine(jsonText);
            }

            if (run) {
                try
                {
                    //Process.Start("dotnet", outputExe);
                    // Algunas veces no se ejecuta,                      (17/Sep/20)
                    // porque el path contiene espacios.
                    Process.Start("dotnet", $"{'\"'}{outputExe}{'\"'}");
                }
                catch { }
            }
            return outputExe;
        }

        /// <summary>
        /// Devuelve el directorio y la versión mayor 
        /// del path con las DLL para aplicaciones NETCore.App.
        /// </summary>
        /// <remarks>08/Sep/2020</remarks>
        public static (string Dir, string Version) NETCoreApp()
        {
            return Compiler.NETCoreApp();
        }

        /// <summary>
        /// Devuelve el directorio y la versión mayor
        /// del path con las DLL de Microsoft.WindowsDesktop.App.
        /// </summary>
        /// <remarks>08/Sep/2020</remarks>
        public static (string Dir, string Version) WindowsDesktopApp()
        {
            return Compiler.WindowsDesktopApp();
        }

        /// <summary>
        /// Devuelve la versión de la DLL.
        /// Si completa es True, se devuelve también el nombre de la DLL:
        /// gsCompilarCore v 1.0.0.0 (para .NET Core 3.1 revisión del dd/MMM/yyyy)
        /// </summary>
        public static string Version(bool completa = false)
        {
            var ensamblado = System.Reflection.Assembly.GetExecutingAssembly();
            
            var versionAttr = ensamblado.GetCustomAttributes(typeof(System.Reflection.AssemblyVersionAttribute), false);
            var vers = versionAttr.Length > 0 ? (versionAttr[0] as System.Reflection.AssemblyVersionAttribute).Version
                                              : "1.0.0.0";

            var fileVerAttr = ensamblado.GetCustomAttributes(typeof(System.Reflection.AssemblyFileVersionAttribute), false);
            var versF = fileVerAttr.Length > 0 ? (fileVerAttr[0] as System.Reflection.AssemblyFileVersionAttribute).Version
                                               : "1.0.0.13";

            var res = $"v {vers} ({versF})";

            if (completa)
            {
                var prodAttr = ensamblado.GetCustomAttributes(typeof(System.Reflection.AssemblyProductAttribute), false);
                var producto = prodAttr.Length > 0 ? (prodAttr[0] as System.Reflection.AssemblyProductAttribute).Product 
                                                   : "gsCompilarCore";

                // La descripción, tomar solo el final              (11/Sep/20)
                var descAttr = ensamblado.GetCustomAttributes(typeof(System.Reflection.AssemblyDescriptionAttribute), false);
                var desc = descAttr.Length > 0 ? (descAttr[0] as System.Reflection.AssemblyDescriptionAttribute).Description 
                                               : "(para .NET Core 3.1 revisión del 14/Sep/2020)";
                desc = desc.Substring(desc.IndexOf("(para .NET"));

                res = $"{producto} {res} {desc}";
            }
            return res;
        }
    }
}