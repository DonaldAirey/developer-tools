// <copyright file="SetModuleHeaderCommand.cs" company="Gamma Four, Inc.">
//    Copyright © 2018 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    using System;
    using System.ComponentModel.Design;
    using EnvDTE;
    using EnvDTE80;
    using GammaFour.DeveloperTools.Properties;
    using Microsoft;
    using Microsoft.VisualStudio.Shell;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Command handler.
    /// </summary>
    internal sealed class SetModuleHeaderCommand
    {
        /// <summary>
        /// Command identifier.
        /// </summary>
        private const int CommandId = 0x0004;

        /// <summary>
        /// The environment for the command.
        /// </summary>
        private static DTE2 environment;

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        private static SetModuleHeaderCommand instance;

        /// <summary>
        /// Initializes a new instance of the <see cref="SetModuleHeaderCommand"/> class.
        /// </summary>
        /// <param name="package">Owner package.</param>
        /// <param name="commandService">Command service to add command to.</param>
        private SetModuleHeaderCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            // Validate the arguments.
            package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            // The environment is needed to examine and modify the active document.
            IServiceProvider serviceProvider = package as IServiceProvider;
            SetModuleHeaderCommand.environment = serviceProvider.GetService(typeof(DTE)) as DTE2;
            Assumes.Present(SetModuleHeaderCommand.environment);

            // This installs our custom command into the environment.
            commandService.AddCommand(
                new MenuCommand(this.Execute, new CommandID(DeveloperToolsPackage.CommandSet, SetModuleHeaderCommand.CommandId)));
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <returns>An awaitable task.</returns>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Execute this method on the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            // Instantiate the command.
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            SetModuleHeaderCommand.instance = new SetModuleHeaderCommand(package, commandService);
        }

        /// <summary>
        /// Prompts the user for the copyright notice.
        /// </summary>
        internal static void SetHeader()
        {
            // This dialog will prompt the user for the header information.
            ModuleHeaderDialog headerDialog = new ModuleHeaderDialog
            {
                Author = Settings.Default.Author,
                Company = Settings.Default.Company,
                CopyrightNotice = Settings.Default.CopyrightNotice,
            };

            // Prompt the user and wait for the response.
            bool? response = headerDialog.ShowDialog();

            // If the dialog was accepted then use these values for our header.  Note that the 'IsHeadingSet' can be used to detect whether the user
            // has initialized this information or not and that the values are saved immediately after being set.
            if (response.HasValue && response.Value)
            {
                Settings.Default.Author = headerDialog.Author;
                Settings.Default.Company = headerDialog.Company;
                Settings.Default.CopyrightNotice = headerDialog.CopyrightNotice;
                Settings.Default.IsHeadingSet = true;
                Settings.Default.Save();
            }
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
            // This does the actal work of prompting the user for the header values.
            SetModuleHeaderCommand.SetHeader();
        }
    }
}