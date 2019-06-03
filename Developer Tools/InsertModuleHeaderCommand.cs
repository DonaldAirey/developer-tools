// <copyright file="InsertModuleHeaderCommand.cs" company="Gamma Four, Inc.">
//    Copyright © 2018 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    using System;
    using System.ComponentModel.Design;
    using System.Globalization;
    using System.IO;
    using EnvDTE;
    using EnvDTE80;
    using GammaFour.DeveloperTools.Properties;
    using Microsoft;
    using Microsoft.VisualStudio.Shell;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Command handler.
    /// </summary>
    internal sealed class InsertModuleHeaderCommand
    {
        /// <summary>
        /// Command identifier.
        /// </summary>
        private const int CommandId = 0x0003;

        /// <summary>
        /// The environment for the command.
        /// </summary>
        private static DTE2 environment;

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        private static InsertModuleHeaderCommand instance;

        /// <summary>
        /// Initializes a new instance of the <see cref="InsertModuleHeaderCommand"/> class.
        /// </summary>
        /// <param name="package">Owner package.</param>
        /// <param name="commandService">Command service to add command to.</param>
        private InsertModuleHeaderCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            // Validate the arguments.
            package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            // The environment is needed to examine and modify the active document.
            IServiceProvider serviceProvider = package as IServiceProvider;
            InsertModuleHeaderCommand.environment = serviceProvider.GetService(typeof(DTE)) as DTE2;
            Assumes.Present(InsertModuleHeaderCommand.environment);

            // This installs our custom command into the environment.
            commandService.AddCommand(
                new MenuCommand(this.Execute, new CommandID(DeveloperToolsPackage.CommandSet, InsertModuleHeaderCommand.CommandId)));
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
            InsertModuleHeaderCommand.instance = new InsertModuleHeaderCommand(package, commandService);
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
            // Make sure that the user has set the (or accepted the default) value for a wrap margin.
            if (!Settings.Default.IsHeadingSet)
            {
                SetModuleHeaderCommand.SetHeader();
            }

            // We can only execute this command when an active document is present.
            if (InsertModuleHeaderCommand.environment.ActiveDocument != null)
            {
                // Get the selected text from the environment.
                TextSelection selection = InsertModuleHeaderCommand.environment.ActiveDocument.Selection as TextSelection;

                // This command will place a copyright notice in the first block of comments it finds.  This is generally the block that describes
                // the class found in the file.  The general idea here is to move the cursor to the start of the file and then search for the first
                // set of block comments.  Once that is found, we'll continue to read until there are no more comments in the block and place the
                // copyright notice there.  This is likely a 95% solution.  The alternative was to place the copyright notice on the same line as the
                // cursor, but that would involve guessing at the indentation and the exact format of the comment prefix.
                EditPoint startPoint = selection.AnchorPoint.CreateEditPoint();
                startPoint.StartOfDocument();

                // The active document's full name is normalized to lower case (so why does StyleCop force us to normalize in upper case?).  So we'd
                // have a problem if we tried to use the name of the module found in this structure: since the analyzers check the case-sensitive
                // name of the file against the name of the module in the header, a name that is normalized to lower case will fail.  So we need to
                // go the long way around to find a function that returns the non-normalized name of the file.
                string directoryPath = Path.GetDirectoryName(InsertModuleHeaderCommand.environment.ActiveDocument.FullName);
                string[] files = Directory.GetFiles(directoryPath, Path.GetFileName(InsertModuleHeaderCommand.environment.ActiveDocument.Name));
                string moduleName = Path.GetFileName(files[0]);

                // This inserts the boilerplate header at the start of the module.
                startPoint.Insert(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        Resources.ModuleHeaderTemplate,
                        "// ",
                        moduleName,
                        Settings.Default.Company,
                        Settings.Default.CopyrightNotice,
                        Settings.Default.Author));
            }
        }
    }
}