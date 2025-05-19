// <copyright file="ModuleHeaderDialog.xaml.cs" company="Gamma Four, Inc.">
//    Copyright © 2025 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    using System;
    using System.Windows;
    using System.Windows.Input;
    using Microsoft.VisualStudio.PlatformUI;

    /// <summary>
    /// Interaction logic for the header dialog.
    /// </summary>
    internal partial class ModuleHeaderDialog : DialogWindow
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ModuleHeaderDialog"/> class.
        /// </summary>
        internal ModuleHeaderDialog()
        {
            // Initialize the IDE managed resources.
            this.InitializeComponent();

            // Bind the view model to the model.
            this.DataContext = this;

            // Bind the View Model to the model.
            this.CommandBindings.Add(new CommandBinding(EditorCommands.OK, this.OnOK));
            this.CommandBindings.Add(new CommandBinding(EditorCommands.Cancel, this.OnCancel));
        }

        /// <summary>
        /// Gets or sets the author of the module.
        /// </summary>
        /// <value>
        /// The author of the module.
        /// </value>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the company that owns the module.
        /// </summary>
        /// <value>
        /// The company that owns the module.
        /// </value>
        public string Company { get; set; }

        /// <summary>
        /// Gets or sets the copyright notice.
        /// </summary>
        /// <value>
        /// The copyright notice.
        /// </value>
        public string CopyrightNotice { get; set; }

        /// <summary>
        /// Handles the OK command.
        /// </summary>
        /// <param name="sender">The object that originated the event.</param>>
        /// <param name="routedEventArgs">The routed event data.</param>
        private void OnOK(object sender, RoutedEventArgs routedEventArgs)
        {
            this.DialogResult = true;
            this.Close();
        }

        /// <summary>
        /// Handles the Cancel command.
        /// </summary>
        /// <param name="sender">The object that originated the event.</param>>
        /// <param name="routedEventArgs">The routed event data.</param>
        private void OnCancel(object sender, RoutedEventArgs routedEventArgs)
        {
            this.Close();
        }
    }
}