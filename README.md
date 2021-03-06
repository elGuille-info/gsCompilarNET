# gsCompilarNET
Biblioteca de clases (.NET Core 3.1) para compilar código de C# y VB.<br>
Basada en código en C# de Laurent Kempé (Copyright (c) 2019 Laurent Kempé) y otras fuentes de la web.<br>
<br>
El lenguaje de código usado es Visual Basic .NET.<br>
<br>
<br>
La carpeta <b>actual</b> tiene el código del proyecto actual con las modificaciones.<br>
La carpeta <b>actual-cs</b> tiene el código en C# que adapté (modifiqué) a partir del código original de Laurent Kempé<br>
<br>
A fecha de hoy 14 de septiembre de 2019 (18:33 GMT+2) es el mismo código que el que he convertido a Visual Basic .NET<br>
En realidad después de convertirlo a VB lo he mejorado, añadido nuevas características, etc., y después lo he pasado al proyecto de C#.
<br>
<br>
<h2>IMPORTANTE:</h2>
Esta biblioteca está codificada para usar en .NET Core 3.1<br>
<br>
Puedes usar este código en tus proyectos sin ninguna restricción, así como la DLL una vez compilada, indicando que parte del código 
está basado en las clases de C# Copyright (c) 2019 Laurent Kempé<br>
<br>
<br>
<h2>Uso de esta DLL:</h2>
Actualmente esta biblioteca la utilizo personalmente en:<br>
gsCompilarEjecutarNET v1.0.0.5 para .NET 5.0 RC2.<br>
<b>NOTA:</b><br>
Con fecha del 25 de octubre de 2020 la aplicación gsCompilarEjecutarNET la considero obsoloeta.<br>
En su lugar recomiendo usar <a href="https://github.com/elGuille-info/gsEvaluarColorearCodigoNET">gsEvaluarColorearCodigoNET</a> (que no usa esta DLL, pero sí código mejorado y ampliado).<br>
<br>
Y en varias aplicaciones de prueba para .NET 5.0 preview 8:<br>
CompilarCore_App_VB y CompilarCore_App_CS (aplicaciones de consola de prueba)<br>
<br> 
<br>
He publicado un paquete en NuGet con el código compilado (Release) por si lo quieres usar en tus proyectos de Visual Studio.<br>
https://www.nuget.org/packages/gsCompilarNET/<br>
<br>
<br>
<h2>Versiones</h2>
v1.0.0.5 del 25 de octubre de 2020<br>
Cambio las versiones de json para usar .NET 5.0 RC2 (WinForms: 5.0.0-rc.2.20475.6, Consola: 5.0.0-rc.2.20475.5)
v1.0.0.4 del 21 de septiembre de 2020<br>
Cambio las versiones de json para aplicaciones de WinForms a 5.0.0-rc.1.20452.2 y Consola a 5.0.0-rc.1.20451.14.<br>
<br>
v1.0.0.3 del 19 de septiembre de 2020<br>
Se pueden indicar las versiones del lenguaje para compilar.<br>
La versión predeterminada es Default. Que es la última versión soportada.<br>
En VB Latest o Default para la última versión (16.0). En C# Latest (8.0) o Default o Preview para 9.0.<br>
<br>
v1.0.0.2 <br>
Falla si el path de la DLL a ejecutar contiene espacios.<br>
El paquete de NuGet lo sincronizo con la versión revisada (FileVersion).<br>
<br>
<br>
Guillermo<br>
<br>
Actualizado el 25 de octubre de 2020 a eso de las 14:16 GMT+1

