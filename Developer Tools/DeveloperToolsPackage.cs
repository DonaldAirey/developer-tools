// <copyright file="DeveloperToolsPackage.cs" company="Dark Bond, Inc.">
//    Copyright © 2016-2017 - Dark Bond, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace DarkBond.Tools
{
    using System;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Microsoft.VisualStudio.Shell;

    /// <summary>
    /// The Developer Tools Package.
    /// </summary>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(DeveloperToolsPackage.PackageGuidstring)]
    public sealed class DeveloperToolsPackage : AsyncPackage
    {
        /// <summary>
        /// DeveloperToolsPackage GUID string.
        /// </summary>
        internal const string PackageGuidstring = "690d896c-c8f4-475a-9d31-9c28e7d1c4ba";

        /// <summary>
        /// The command set.
        /// </summary>
        private static Guid commandSetField = new Guid("387d7b3c-a527-463d-96ce-b1ba9b7c1d85");

        /// <summary>
        /// Gets the command set used to group the commands in this package.
        /// </summary>
        /// <value>
        /// The command set used to group the commands in this package.
        /// </value>
        internal static Guid CommandSet
        {
            get
            {
                return DeveloperToolsPackage.commandSetField;
            }
        }

        /// <summary>
        /// Called when the VSPackage is loaded by Visual Studio.
        /// </summary>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <param name="progress">The progress indicator.</param>
        /// <returns>A <see cref="System.Threading.Tasks.Task"/> representing the asynchronous operation.</returns>
        protected override System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // Initialize the commands.
            FormatCommentCommand.Initialize(this);
            InsertModuleHeaderCommand.Initialize(this);
            SetWrapMarginCommand.Initialize(this);
            SetModuleHeaderCommand.Initialize(this);
            InsertConstructorHeaderCommand.Initialize(this);
            FormatXmlCommand.Initialize(this);

            // This is a no-op, but it satisfies the asynchronous interfae.
            return System.Threading.Tasks.Task.CompletedTask;
        }
    }
}