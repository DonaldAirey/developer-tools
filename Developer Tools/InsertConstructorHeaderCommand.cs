// <copyright file="InsertConstructorHeaderCommand.cs" company="Gamma Four, Inc.">
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
    /// Command handler.
    /// </summary>
    internal sealed class InsertConstructorHeaderCommand
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
        /// Gets the instance of the command.
        /// </summary>
        private static InsertConstructorHeaderCommand instance;

        /// <summary>
        /// Initializes a new instance of the <see cref="InsertConstructorHeaderCommand"/> class.
        /// </summary>
        /// <param name="package">Owner package.</param>
        /// <param name="commandService">Command service to add command to.</param>
        private InsertConstructorHeaderCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            // Validate the arguments.
            package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            // The environment is needed to examine and modify the active document.
            IServiceProvider serviceProvider = package as IServiceProvider;
            InsertConstructorHeaderCommand.environment = serviceProvider.GetService(typeof(DTE)) as DTE2;
            Assumes.Present(InsertConstructorHeaderCommand.environment);

            // This installs our custom command into the environment.
            commandService.AddCommand(
                new MenuCommand(this.Execute, new CommandID(DeveloperToolsPackage.CommandSet, InsertConstructorHeaderCommand.CommandId)));
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
            InsertConstructorHeaderCommand.instance = new InsertConstructorHeaderCommand(package, commandService);
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
            // Verify the current thread is the UI thread.
            ThreadHelper.ThrowIfNotOnUIThread();

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