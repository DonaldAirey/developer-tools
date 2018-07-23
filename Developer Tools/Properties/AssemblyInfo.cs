// <copyright file="AssemblyInfo.cs" company="Dark Bond, Inc.">
//    Copyright © 2016-2017 - Dark Bond, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
using System;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;

// General information about the assembly.
[assembly: AssemblyTitle("Developer Tools")]
[assembly: AssemblyDescription("Productivity commands for the Visual Studio environment.")]
[assembly: AssemblyCompany("DarkBond, Inc.")]
[assembly: AssemblyProduct("DarkBond")]
[assembly: AssemblyCopyright("Copyright © 2016-2017, DarkBond, Inc.  All rights reserved.")]

// Indicates that this assembly is compliant with the Common Language Specification (CLS).
[assembly: CLSCompliant(false)]

// Disables the accessibility of this assembly to COM.
[assembly: ComVisible(false)]

// Describes the default language used for the resources.
[assembly: NeutralResourcesLanguageAttribute("en-US")]

// Version information for this assembly.
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

// Suppress these FXCop issues.
[assembly: SuppressMessage("Microsoft.Design", "CA2210:AssembliesShouldHaveValidStrongNames", Justification = "Reviewed")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1703:ResourcestringsShouldBeSpelledCorrectly", MessageId = "cref", Scope = "resource", Target = "DarkBond.Tools.Properties.Resources.resources", Justification = "This is a stupid rule")]
[assembly: SuppressMessage("Microsoft.Design", "CA1020:AvoidNamespacesWithFewTypes", Scope = "namespace", Target = "DarkBond.Tools", Justification = "Namespace is used exclusively for tools")]
