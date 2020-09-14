'------------------------------------------------------------------------------
' Clase Runner para ejecutar código tipo consola                    (14/Sep/20)
'
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
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.Loader


Friend Class Runner
    ''' <summary>
    ''' Ejecuta el contenido del ensamblado pasado
    ''' (solo aplicaciones de consola).
    ''' Devuelve una cadena vacía si no hubo error.
    ''' </summary>
    Public Function Execute(ByVal compiledAssembly As Byte(), ByVal args As String()) As String
        Dim assemblyLoadContextWeakRef = LoadAndExecute(compiledAssembly, args)

        Dim i = 0
        Do While i < 8 AndAlso assemblyLoadContextWeakRef.IsAlive
            GC.Collect()
            GC.WaitForPendingFinalizers()
            i += 1
        Loop

        If assemblyLoadContextWeakRef.IsAlive Then
            'Console.WriteLine("Unloading failed!")
            Return "Fallo al descargar el ensamblado de la memoria."
        End If

        Return ""
    End Function

    <MethodImpl(MethodImplOptions.NoInlining)>
    Private Shared Function LoadAndExecute(ByVal compiledAssembly As Byte(), ByVal args As String()) As WeakReference
        Using asm = New MemoryStream(compiledAssembly)
            Dim assemblyLoadContext = New SimpleUnloadableAssemblyLoadContext()
            Dim assembly = assemblyLoadContext.LoadFromStream(asm)
            Dim entry = assembly.EntryPoint

            Dim _o = If(entry IsNot Nothing AndAlso entry.GetParameters().Length > 0,
                            entry.Invoke(Nothing, New Object() {args}),
                            entry.Invoke(Nothing, Nothing)
                        )

            assemblyLoadContext.Unload()

            Return New WeakReference(assemblyLoadContext)
        End Using
    End Function
End Class

Friend Class SimpleUnloadableAssemblyLoadContext
    Inherits AssemblyLoadContext

    Public Sub New()
        MyBase.New(True)
    End Sub

    Protected Overrides Function Load(ByVal assemblyName As AssemblyName) As Assembly
        Return Nothing
    End Function
End Class

