using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("NPOI OpenXml4Net")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("Nissl")]
[assembly: AssemblyProduct("NPOI")]
[assembly: AssemblyCopyright("Apache 2.0")]
[assembly: AssemblyTrademark("NPOI")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("13f2a810-331a-40b6-8d7a-1322b405fab7")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("2.5.3.0")]
[assembly: AssemblyFileVersion("2.5.3.0")]
[assembly: AssemblyInformationalVersion("2.0.0.0")]
#if NETSTANDARD2_1 || NETSTANDARD2_0 || NET40
[assembly: AllowPartiallyTrustedCallers]
#endif
#if NETSTANDARD2_1 || NETSTANDARD2_0 || NET40 || NET45
[assembly: SecurityRules(SecurityRuleSet.Level1)]
#endif
