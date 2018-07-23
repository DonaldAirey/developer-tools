// <copyright file="InsertModuleHeaderCommand.cs" company="Dark Bond, Inc.">
//    Copyright © 2016-2017 - Dark Bond, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace DarkBond.Tools
{
    using System;
    using System.ComponentModel.Design;
    using System.Globalization;
    using System.IO;
    using System.Windows.Threading;
    using DarkBond.Tools.Properties;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft;
    using Microsoft.VisualStudio.Shell;

    /// <summary>
    /// Inserts a block comment at the start of a C# module.
    /// </summary>
    internal static class InsertModuleHeaderCommand
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
        /// The service provider.
        /// </summary>
        private static IServiceProvider serviceProvider;

        /// <summary>
        /// Initialize the command.
        /// </summary>
        /// <param name="package">The package to which this command belongs.</param>
        internal static void Initialize(Package package)
        {
            // Validate the 'package' argument.
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            // The VS Package provides services for examining the Visual Studio environment.
            InsertModuleHeaderCommand.serviceProvider = package as IServiceProvider;
            InsertModuleHeaderCommand.environment = InsertModuleHeaderCommand.serviceProvider.GetService(typeof(DTE)) as DTE2;
            Assumes.Present(InsertModuleHeaderCommand.environment);

            // This installs our custom command into the environment.
            OleMenuCommandService oleMenuCommandService =
                InsertModuleHeaderCommand.serviceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (oleMenuCommandService != null)
            {
                oleMenuCommandService.AddCommand(
                    new MenuCommand(InsertModuleHeaderCommand.ExecuteCommand, new CommandID(DeveloperToolsPackage.CommandSet, CommandId)));
            }
        }

        /// <summary>
        /// Executes the command.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="eventArgs">An object that contains no event data.</param>
        private static void ExecuteCommand(object sender, EventArgs eventArgs)
        {
            // This insures we're on the main thread.
            Dispatcher.CurrentDispatcher.VerifyAccess();

            // Make sure that the user has set the (or accepted the default) value for a wrap margin.
            if (!Properties.Settings.Default.IsHeadingSet)
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