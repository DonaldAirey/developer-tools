// <copyright file="SetWrapMarginCommand.cs" company="Gamma Four, Inc.">
//    Copyright © 2018 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    using System;
    using System.ComponentModel.Design;
    using System.Globalization;
    using EnvDTE;
    using EnvDTE80;
    using GammaFour.DeveloperTools.Properties;
    using Microsoft;
    using Microsoft.VisualStudio.Shell;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SetWrapMarginCommand
    {
        /// <summary>
        /// Command identifier.
        /// </summary>
        private const int CommandId = 0x0002;

        /// <summary>
        /// The environment for the developer tools.
        /// </summary>
        private static DTE2 environment;

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        private static SetWrapMarginCommand instance;

        /// <summary>
        /// Initializes a new instance of the <see cref="SetWrapMarginCommand"/> class.
        /// </summary>
        /// <param name="package">Owner package.</param>
        /// <param name="commandService">Command service to add command to.</param>
        private SetWrapMarginCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            // Validate the arguments.
            package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            // The environment is needed to examine and modify the active document.
            IServiceProvider serviceProvider = package as IServiceProvider;
            SetWrapMarginCommand.environment = serviceProvider.GetService(typeof(DTE)) as DTE2;
            Assumes.Present(SetWrapMarginCommand.environment);

            // This installs our custom command into the environment.
            commandService.AddCommand(
                new MenuCommand(this.Execute, new CommandID(DeveloperToolsPackage.CommandSet, SetWrapMarginCommand.CommandId)));
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <returns>An awaitable task.</returns>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Verify the current thread is the UI thread.
            ThreadHelper.ThrowIfNotOnUIThread();

            // Instantiate the command.
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            SetWrapMarginCommand.instance = new SetWrapMarginCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            // This dialog will prompt the user for the header information.
            WrapMarginDialog wrapmarginDialog = new WrapMarginDialog
            {
                WrapMargin = Convert.ToString(Settings.Default.WrapMargin, CultureInfo.InvariantCulture)
            };

            // Prompt the user and wait for the response.
            bool? response = wrapmarginDialog.ShowDialog();

            // If the dialog was accepted then use these values for our header.  Note that the 'IsHeadingSet' can be used to detect whether the user
            // has initialized this information or not and that the values are saved immediately after being set.
            if (response.HasValue && response.Value)
            {
                Settings.Default.WrapMargin = Convert.ToInt32(wrapmarginDialog.WrapMargin, CultureInfo.InvariantCulture);
                Settings.Default.Save();
            }
        }
    }
}