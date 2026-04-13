using System.Reflection;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Runtime;

[assembly: AssemblyTitle("WordToCadPlugin")]
[assembly: AssemblyDescription("AutoCAD plugin that converts each page of a Word document to PNG and inserts them as a grid into the current drawing.")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("WordToCadPlugin")]
[assembly: AssemblyCopyright("Copyright © 2026")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

[assembly: ComVisible(false)]
[assembly: Guid("9F3A8E2D-4B5C-4F1E-9D7A-1A2B3C4D5E6F")]

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

// Register command class with AutoCAD so [CommandMethod] attributes are scanned at NETLOAD time.
[assembly: CommandClass(typeof(WordToCadPlugin.Commands))]
