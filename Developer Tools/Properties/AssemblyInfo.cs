// <copyright file="AssemblyInfo.cs" company="Gamma Four, Inc.">
//    Copyright © 2018 - Gamma Four, Inc.  All Rights Reserved.
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
[assembly: AssemblyCompany("GammaFour, Inc.")]
[assembly: AssemblyProduct("GammaFour")]
[assembly: AssemblyCopyright("Copyright © 2018, GammaFour, Inc.  All rights reserved.")]

// Indicates that this assembly is compliant with the Common Language Specification (CLS).
[assembly: CLSCompliant(false)]

// Disables the accessibility of this assembly to COM.
[assembly: ComVisible(false)]

// Describes the default language used for the resources.
[assembly: NeutralResourcesLanguageAttribute("en-US")]

// Version information for this assembly.
[assembly: AssemblyVersion("2.10.0.0")]
[assembly: AssemblyFileVersion("2.10.0.0")]

// Suppress these FXCop issues.
[assembly: SuppressMessage("Microsoft.Design", "CA2210:AssembliesShouldHaveValidStrongNames", Justification = "Reviewed")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1703:ResourcestringsShouldBeSpelledCorrectly", MessageId = "cref", Scope = "resource", Target = "GammaFour.DeveloperTools.Properties.Resources.resources", Justification = "This is a stupid rule")]
[assembly: SuppressMessage("Microsoft.Design", "CA1020:AvoidNamespacesWithFewTypes", Scope = "namespace", Target = "GammaFour.DeveloperTools", Justification = "Namespace is used exclusively for tools")]
