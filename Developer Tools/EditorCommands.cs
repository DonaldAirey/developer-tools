// <copyright file="EditorCommands.cs" company="Gamma Four, Inc.">
//    Copyright © 2025 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    using System.Windows.Input;

    /// <summary>
    /// The commands used for the Editor Extension dialogs.
    /// </summary>
    internal static class EditorCommands
    {
        /// <summary>
        /// Creates a slice from the selected orders and sends them to a destination.
        /// </summary>
        private static RoutedCommand okField = new RoutedCommand("OK", typeof(EditorCommands));

        /// <summary>
        /// Opens the properties dialog for the selected item.
        /// </summary>
        private static RoutedCommand cancelField = new RoutedCommand("Cancel", typeof(EditorCommands));

        /// <summary>
        /// Gets the OK command.
        /// </summary>
        /// <value>
        /// The OK command.
        /// </value>
        public static RoutedCommand OK
        {
            get
            {
                return EditorCommands.okField;
            }
        }

        /// <summary>
        /// Gets the Cancel command.
        /// </summary>
        /// <value>
        /// The Cancel command.
        /// </value>
        public static RoutedCommand Cancel
        {
            get
            {
                return EditorCommands.cancelField;
            }
        }
    }
}