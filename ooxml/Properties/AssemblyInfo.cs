using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("NPOI OOXML")]
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
[assembly: Guid("cc03d84d-498a-4561-97c1-e39d5d7780a0")]

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

//[assembly: InternalsVisibleTo("ooxml.Testcases")]

[assembly: InternalsVisibleTo("NPOI.OOXML.TestCases, PublicKey=002400000480000094000000060200000024000052534131000400000100010095ccd95af3b39d8bc20544d3f47fd24b53ebc5ccb693eaed116290629f8cd882c827ebd511ad59449224f0718d3f9d03b64945a6c8b6644266001b8c8426185330e3d96da70ae16d4acc21b8d4d480f1385c7e924273179375aa88f81380a72fb115712a313379d16aed4aa36208ee3b4a5dd785b06a07b2d868e3227f4495b5", AllInternalsVisible = true)]
#if NETSTANDARD2_1 || NETSTANDARD2_0 || NET40
[assembly: AllowPartiallyTrustedCallers]
#endif
#if NETSTANDARD2_1 || NETSTANDARD2_0 || NET40 || NET45
[assembly: SecurityRules(SecurityRuleSet.Level1)]
#endif
