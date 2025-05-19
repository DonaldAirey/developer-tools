// <copyright file="DeveloperToolsPackage.cs" company="Gamma Four, Inc.">
//    Copyright © 2025 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    using System;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Microsoft.VisualStudio.Shell;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(DeveloperToolsPackage.PackageGuidString)]
    public sealed class DeveloperToolsPackage : AsyncPackage
    {
        /// <summary>
        /// Package identifier.
        /// </summary>
        public const string PackageGuidString = "e86fa1fd-2e90-40a9-a80d-45e405cb5e6b";

        /// <summary>
        /// Gets the Command Set identifier.
        /// </summary>
        internal static Guid CommandSet { get; } = new Guid("4a8eeab1-3e50-413b-ada5-40f8b62cd1fb");

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            await FormatCommentCommand.InitializeAsync(this);
            await FormatXmlCommand.InitializeAsync(this);
            await InsertModuleHeaderCommand.InitializeAsync(this);
            await InsertConstructorHeaderCommand.InitializeAsync(this);
            await SetModuleHeaderCommand.InitializeAsync(this);
            await SetWrapMarginCommand.InitializeAsync(this);
            await ScrubXsdCommand.InitializeAsync(this);
        }
    }
}