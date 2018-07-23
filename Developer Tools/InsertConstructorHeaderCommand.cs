// <copyright file="InsertConstructorHeaderCommand.cs" company="Dark Bond, Inc.">
//    Copyright © 2016-2017 - Dark Bond, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace DarkBond.Tools
{
    using System;
    using System.ComponentModel.Design;
    using System.Globalization;
    using System.Windows.Threading;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft;
    using Microsoft.VisualStudio.Shell;
    using Properties;

    /// <summary>
    /// Inserts a standard header for a constructor.
    /// </summary>
    internal static class InsertConstructorHeaderCommand
    {
        /// <summary>
        /// Finds the name of the class or structure.
        /// </summary>
        private const string ClassDeclarationPattern = @"^(public|internal|private|abstract|sealed|static|\s+)*(class|struct)\s+(\w+)";

        /// <summary>
        /// Command identifier.
        /// </summary>
        private const int CommandId = 0x0005;

        /// <summary>
        /// Finds the constructor of the class or structure.
        /// </summary>
        private const string ConstructorDeclarationPattern = @"(static\s+)?{0}\s*\([^\)]*\)";

        /// <summary>
        /// Finds the constructor summary comment.
        /// </summary>
        private const string ConstructorSummaryPattern = @"^([ \t]*///\s*)<summary>(.|\r?\n)*?</summary>\s*\r?\n";

        /// <summary>
        /// The environment for the developer tools.
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
            InsertConstructorHeaderCommand.serviceProvider = package as IServiceProvider;
            InsertConstructorHeaderCommand.environment = InsertConstructorHeaderCommand.serviceProvider.GetService(typeof(DTE)) as DTE2;
            Assumes.Present(InsertConstructorHeaderCommand.environment);

            // This installs our custom command into the environment.
            OleMenuCommandService oleMenuCommandService =
                InsertConstructorHeaderCommand.serviceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (oleMenuCommandService != null)
            {
                oleMenuCommandService.AddCommand(
                    new MenuCommand(InsertConstructorHeaderCommand.ExecuteCommand, new CommandID(DeveloperToolsPackage.CommandSet, CommandId)));
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

            // Don't attempt to insert a header if there's no active document.
            if (InsertConstructorHeaderCommand.environment.ActiveDocument == null)
            {
                return;
            }

            // Get the selected text from the editor environment.
            TextSelection selection = InsertConstructorHeaderCommand.environment.ActiveDocument.Selection as TextSelection;

            // Create an edit point for the current location in the document.  The general idea here is to look backward from the current cursor
            // location until we find a class or structure declaration.  We'll then extract the name of the class (or structure) and look for the
            // constructor.  Once we find it, we'll replace the existing summary with the new boilerplate version that Style COP wants to see.
            EditPoint startPoint = selection.AnchorPoint.CreateEditPoint();

            // The 'FindPattern' will return a collection of found sub-tags in this collection.  The function seems to be a little brain damaged or
            // else it isn't documented clearly.  If you have two sub-tags in your pattern, three text ranges are returned.  It appears that one of
            // the ranges is the entire match, though it shows up in no particular order.
            TextRanges textRanges = null;

            // This starts the process by finding the class or structure declaration.
            if (startPoint.FindPattern(
                InsertConstructorHeaderCommand.ClassDeclarationPattern,
                (int)(vsFindOptions.vsFindOptionsRegularExpression | vsFindOptions.vsFindOptionsBackwards),
                null,
                ref textRanges))
            {
                // Whether the object is a class or a structure will be used when constructing the boilerplate summary comment.  The name of the
                // class is also used to create the summary.  The 'index' to the 'Item' indexer is a little brain damaged as well.  It turns out that
                // the COM class will only work with floating point numbers (go figure) and starts counting at 1.0.
                var objectTypeRange = textRanges.Item(3.0);
                string objectType = objectTypeRange.StartPoint.GetText(objectTypeRange.EndPoint);
                var classNameRange = textRanges.Item(4.0);
                string className = classNameRange.StartPoint.GetText(classNameRange.EndPoint);

                TextRanges modifierRanges = null;

                // This will find the constructor for the class or structure.
                if (startPoint.FindPattern(
                    string.Format(CultureInfo.InvariantCulture, InsertConstructorHeaderCommand.ConstructorDeclarationPattern, className),
                    (int)vsFindOptions.vsFindOptionsRegularExpression,
                    null,
                    ref modifierRanges))
                {
                    var modifierRange = modifierRanges.Item(1.0);
                    string modifier = modifierRange.StartPoint.GetText(modifierRange.EndPoint);
                    bool isStatic = modifier.StartsWith("static", StringComparison.Ordinal);

                    // This will find the summary comment and return the entire range of the comment from the first line to the end of the line
                    // containing the ending tag.  This covers comments that have been compressed onto a single line as well as the more common
                    // multi-line 'summary' comment.
                    EditPoint endPoint = null;
                    if (startPoint.FindPattern(
                        InsertConstructorHeaderCommand.ConstructorSummaryPattern,
                        (int)(vsFindOptions.vsFindOptionsRegularExpression | vsFindOptions.vsFindOptionsBackwards),
                        ref endPoint,
                        ref textRanges))
                    {
                        // Extract the prefix of the comment.  This tells us how many tabs (or spaces) there are in the comment delimiter.
                        var prefixRange = textRanges.Item(1.0);
                        string prefix = prefixRange.StartPoint.GetText(prefixRange.EndPoint);

                        // Static and instance constructors have different formats.  This will chose a header template that is appropriate.
                        string constructorHeader = isStatic ?
                            Resources.StaticConstructorHeaderTemplate :
                            Resources.InstanceConstructorHeaderTemplate;

                        // Replace the existing summary comment with a boilerplate constructed from the information collected about the class.  This
                        // boilerplate comment will pass mustard in Style COP.
                        string replacementText = string.Format(
                            CultureInfo.InvariantCulture,
                            constructorHeader,
                            prefix,
                            className,
                            objectType);
                        startPoint.ReplaceText(
                            endPoint,
                            replacementText,
                            (int)(vsEPReplaceTextOptions.vsEPReplaceTextNormalizeNewlines | vsEPReplaceTextOptions.vsEPReplaceTextTabsSpaces));
                    }
                }
            }
        }
    }
}