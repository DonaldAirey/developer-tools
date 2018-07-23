// <copyright file="SetWrapMarginCommand.cs" company="Dark Bond, Inc.">
//    Copyright © 2016-2017 - Dark Bond, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace DarkBond.Tools
{
    using System;
    using System.ComponentModel.Design;
    using System.Globalization;
    using Microsoft.VisualStudio.Shell;

    /// <summary>
    /// Prompts the user for the wrapping margin to use when justifying comments.
    /// </summary>
    internal static class SetWrapMarginCommand
    {
        /// <summary>
        /// Command identifier.
        /// </summary>
        private const int CommandId = 0x0002;

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
            IServiceProvider serviceProvider = package as IServiceProvider;

            // This installs our custom command into the environment.
            OleMenuCommandService oleMenuCommandService = serviceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (oleMenuCommandService != null)
            {
                oleMenuCommandService.AddCommand(
                    new MenuCommand(SetWrapMarginCommand.ExecuteCommand, new CommandID(DeveloperToolsPackage.CommandSet, CommandId)));
            }
        }

        /// <summary>
        /// Executes the command.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="eventArgs">An object that contains no event data.</param>
        private static void ExecuteCommand(object sender, EventArgs eventArgs)
        {
            // This dialog will prompt the user for the header information.
            WrapMarginDialog wrapmarginDialog = new WrapMarginDialog
            {
                WrapMargin = Convert.ToString(Properties.Settings.Default.WrapMargin, CultureInfo.InvariantCulture)
            };

            // Prompt the user and wait for the response.
            bool? response = wrapmarginDialog.ShowDialog();

            // If the dialog was accepted then use these values for our header.  Note that the 'IsHeadingSet' can be used to detect whether the user
            // has initialized this information or not and that the values are saved immediately after being set.
            if (response.HasValue && response.Value)
            {
                Properties.Settings.Default.WrapMargin = Convert.ToInt32(wrapmarginDialog.WrapMargin, CultureInfo.InvariantCulture);
                Properties.Settings.Default.Save();
            }
        }
    }
}