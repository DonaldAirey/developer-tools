// <copyright file="FormatCommentCommand.cs" company="Gamma Four, Inc.">
//    Copyright © 2025 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.Design;
    using System.Text;
    using System.Text.RegularExpressions;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft;
    using Microsoft.VisualStudio.Shell;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Command handler.
    /// </summary>
    internal sealed class FormatCommentCommand
    {
        /// <summary>
        /// Command identifier.
        /// </summary>
        private const int CommandId = 0x0001;

        /// <summary>
        /// These element names will cause the elements to be emitted on lines separate from the contents.
        /// </summary>
        private static List<string> breakingElements = new List<string> { "summary" };

        /// <summary>
        /// The current position of the column in the output comment.
        /// </summary>
        private static int columnPosition;

        /// <summary>
        /// Regular expression to separate the comment prefix from the rest of the comment line.
        /// </summary>
        private static Regex commentRegex = new Regex(@"^(?<prefix>\s*//+)\s*(?<comment>.*?)\s*$", RegexOptions.Compiled);

        /// <summary>
        /// The environment for the command.
        /// </summary>
        private static DTE2 environment;

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        private static FormatCommentCommand instance;

        /// <summary>
        /// This provides spacing after the comment delimiter and before the comment.
        /// </summary>
        private static string leftMargin = " ";

        /// <summary>
        /// Recognizes discrete words in the comment stream (XML tags are considered words).
        /// </summary>
        private static Regex wordExpression = new Regex(@"(?<word><[^<>]+>)|(?<word>[^\s<$]+)|(?<word><[^\s]+)");

        /// <summary>
        /// Used to separator one comment word from the next.
        /// </summary>
        private static string wordSeparator = string.Empty;

        /// <summary>
        /// Recognizes a empty XML element.
        /// </summary>
        private static Regex xmlEmptyElement = new Regex("<[^>]+/>$", RegexOptions.Compiled);

        /// <summary>
        /// Recognizes the XML end tag.
        /// </summary>
        private static Regex xmlEndTag = new Regex("</[^>]+>$", RegexOptions.Compiled);

        /// <summary>
        /// Recognizes the parts of an XML element.
        /// </summary>
        private static Regex xmlParts = new Regex("^(?<startTag><[^>]+>)(?<body>.*)(?<endTag></[^>]+>)", RegexOptions.Compiled);

        /// <summary>
        /// Recognizes the XML start tag.
        /// </summary>
        private static Regex xmlStartTag = new Regex("^<[^>]+[^/]>", RegexOptions.Compiled);

        /// <summary>
        /// Initializes a new instance of the <see cref="FormatCommentCommand"/> class.
        /// </summary>
        /// <param name="package">Owner package.</param>
        /// <param name="commandService">Command service to add command to.</param>
        private FormatCommentCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            // Validate the arguments.
            package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            // The environment is needed to examine and modify the active document.
            IServiceProvider serviceProvider = package as IServiceProvider;
            FormatCommentCommand.environment = serviceProvider.GetService(typeof(DTE)) as DTE2;
            Assumes.Present(FormatCommentCommand.environment);

            // This installs our custom command into the environment.
            commandService.AddCommand(
                new MenuCommand(this.Execute, new CommandID(DeveloperToolsPackage.CommandSet, FormatCommentCommand.CommandId)));
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <returns>An awaitable task.</returns>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Verify the current thread is the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            // Instantiate the command.
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            FormatCommentCommand.instance = new FormatCommentCommand(package, commandService);
        }

        /// <summary>
        /// Formats a block of comments and beautifies the XML comments.
        /// </summary>
        /// <param name="inputComment">The block of comment lines to format.</param>
        /// <returns>A right justified and beautified block of comments.</returns>
        private static string FormatCommentstring(string inputComment)
        {
            // These variables control the state of the wrapping.
            FormatCommentCommand.columnPosition = 0;

            // This buffer keeps track of the reconstructed comment block.
            StringBuilder outputComment = new StringBuilder();

            // This breaks the input comment into individual lines.
            string[] commentLines = inputComment.Split(new string[] { "\r\n" }, StringSplitOptions.None | StringSplitOptions.RemoveEmptyEntries);

            // Parse the line into prefixes (the comment start sequence and whitespace) and the actual comment on the line.
            for (int index = 0; index < commentLines.Length; index++)
            {
                // This parses the comment and the prefix out of the current comment line.  The prefix is all characters up to the comment delimiter
                // character.  The comment represents all the characters after the delimiter stripped of the leading and trailing space.
                Match match = FormatCommentCommand.commentRegex.Match(commentLines[index]);
                string prefix = match.Groups["prefix"].Value;
                string comment = match.Groups["comment"].Value;

                // We assume there was some intent if an author left an entire line blank and so we treat it as a paragraph break.  That is, we don't
                // wrap over it.
                if (comment.Length == 0)
                {
                    if (FormatCommentCommand.columnPosition != 0)
                    {
                        outputComment.AppendLine();
                    }

                    outputComment.AppendLine(prefix);
                    FormatCommentCommand.columnPosition = 0;
                    continue;
                }

                // We're going to attempt to format the special comment directives so that they'll wrap nicely.  If a single-line directive is too
                // long, we'll wrap it so the directives end up on separate lines.  This test is used to distinguish between the special directives
                // used for commenting functions, classes and modules from regular block comments.
                bool isCommentDirective = prefix.EndsWith("///", StringComparison.Ordinal);

                // This section will provide formatting for XML tags inside comment blocks.  The tags are examined to see if the text and tags will
                // fit within the margins.  If not, the tags are placed on their own lines and the comment inside the tags is formatted to wrap
                // around the margins.  The algorithm will also eat up partial lines in order to fill out the XML content to the margin.
                if (isCommentDirective && xmlStartTag.IsMatch(comment) && !xmlEmptyElement.IsMatch(comment))
                {
                    while (!xmlEndTag.IsMatch(comment) && index < commentLines.Length)
                    {
                        match = FormatCommentCommand.commentRegex.Match(commentLines[++index]);
                        comment += ' ' + match.Groups["comment"].Value;
                    }

                    FormatCommentCommand.WrapXml(outputComment, comment, prefix);
                    continue;
                }

                // This is used to force a line break on comment lines that meet certain criteria, such as bullet marks and examples.
                bool isBreakingLine = false;

                // Lines that begin with an Asterisk are considered bullet marks and do not wrap and are given an extra margin.
                if (comment.StartsWith("*", StringComparison.Ordinal))
                {
                    // If the previous line was waiting for some wrapping to occur, it's going to be disappointed.  This line is going to start all
                    // on its own.
                    if (FormatCommentCommand.columnPosition != 0)
                    {
                        outputComment.AppendLine();
                    }

                    FormatCommentCommand.columnPosition = 0;

                    // The prefix will be indented for all bullet marks.
                    prefix += "    ";

                    // This will force a line break after the bullet mark is formatted.
                    isBreakingLine = true;
                }

                // Lines that end with colons do not wrap.
                if (comment.EndsWith(":", StringComparison.Ordinal))
                {
                    isBreakingLine = true;
                }

                // This is where all the work is done to right justify the block comment to the margin.
                FormatCommentCommand.WrapLine(outputComment, comment, prefix);

                // This will force a new line for comment lines that don't wrap such as bullet marks and colons.
                if (isBreakingLine)
                {
                    outputComment.AppendLine();
                    FormatCommentCommand.columnPosition = 0;
                }
            }

            // This will finish off any incomplete lines.
            if (FormatCommentCommand.columnPosition > 0)
            {
                outputComment.AppendLine();
            }

            // At this point we've transformed the input block of comments into a right justified block.
            return outputComment.ToString();
        }

        /// <summary>
        /// Indicates whether the current point is on a line with only a comment.
        /// </summary>
        /// <param name="point">The current point of the document.</param>
        /// <returns>True if the line contains only a comment.</returns>
        private static bool IsCommentLine(EditPoint point)
        {
            // Verify the current thread is the UI thread.
            ThreadHelper.ThrowIfNotOnUIThread();

            // This will extract the current line from the given point and determine if it passes the sniff test for a comment.
            EditPoint endOfLine = point.CreateEditPoint();
            endOfLine.EndOfLine();
            return FormatCommentCommand.commentRegex.IsMatch(point.GetText(endOfLine));
        }

        /// <summary>
        /// Right justifies the given comment line against the right margin in the output comment.
        /// </summary>
        /// <param name="outputComment">The output comment.</param>
        /// <param name="inputComment">The comment to be wrapped in the output comment.</param>
        /// <param name="prefix">The current comment line's prefix.</param>
        private static void WrapLine(StringBuilder outputComment, string inputComment, string prefix)
        {
            // This will parse each incoming line breaking up each word into a token.  For the purpose of this tokenizer, complete XML tags are
            // considered to be 'words' even though there may be several embedded spaces in the tag.
            Match match = FormatCommentCommand.wordExpression.Match(inputComment);
            while (match.Success)
            {
                // Take the next word token out of the parsed string.
                string word = match.Groups["word"].Value;

                // This block of code basically figures out what should appear before the next work in the stream.  If we exceed the margin, we're
                // going to put in a newline and the comment prefix.  If we're within the margins, we check to see if we've got the end of a sentence
                // or not.
                if (FormatCommentCommand.columnPosition != 0)
                {
                    // There is no spacing between XML tags and the content of the element.
                    if (FormatCommentCommand.xmlEndTag.IsMatch(word))
                    {
                        FormatCommentCommand.wordSeparator = string.Empty;
                    }

                    // If the separator and the next word fit within the margin, then we'll emit a separator between this word and the previous one.
                    // Otherwise we start a new line if we've exceeded the margin.  This is where the 'wrapping' logic occurs.
                    if (FormatCommentCommand.columnPosition + FormatCommentCommand.wordSeparator.Length + word.Length >= Properties.Settings.Default.WrapMargin)
                    {
                        outputComment.AppendLine();
                        FormatCommentCommand.columnPosition = 0;
                    }
                    else
                    {
                        outputComment.Append(FormatCommentCommand.wordSeparator);
                        FormatCommentCommand.columnPosition += FormatCommentCommand.wordSeparator.Length;
                    }
                }

                // The start of a line will always emit the current prefix and pad it with the left margin.
                if (FormatCommentCommand.columnPosition == 0)
                {
                    outputComment.Append(prefix);
                    outputComment.Append(FormatCommentCommand.leftMargin);
                    FormatCommentCommand.columnPosition += prefix.Length + FormatCommentCommand.leftMargin.Length;
                }

                // This is where all words are emitted into the output comment.
                outputComment.Append(word);
                FormatCommentCommand.columnPosition += word.Length;

                // There are two spaces between the end of a sentence and the start of the next.  There is only one space between a normal word and
                // the next.  Finally, there is no space between an XML start tag and the content of that element.  This rule will determine what
                // kind of gap there is between this word and the next.
                FormatCommentCommand.wordSeparator = word.EndsWith(".", StringComparison.Ordinal) ? "  " :
                    FormatCommentCommand.xmlStartTag.IsMatch(word) ? string.Empty : " ";

                // Pull the next token out of the comment line.
                match = match.NextMatch();
            }
        }

        /// <summary>
        /// Beatifies XML comments.
        /// </summary>
        /// <param name="outputComment">The output comment.</param>
        /// <param name="inputComment">The comment to be wrapped and beautified.</param>
        /// <param name="prefix">The current comment line's prefix.</param>
        private static void WrapXml(StringBuilder outputComment, string inputComment, string prefix)
        {
            // The main idea around formatting XML is to put tags on their own lines when the text can't fit within the margin.  To accomplish this,
            // we'll format the entire tag and then see how long is it.
            StringBuilder commentBody = new StringBuilder();
            FormatCommentCommand.WrapLine(commentBody, inputComment, prefix);

            // This will pull apart the XML into the tags and the content.
            Match match = FormatCommentCommand.xmlParts.Match(inputComment);
            string startTag = match.Groups["startTag"].Value;
            string body = match.Groups["body"].Value;
            string endTag = match.Groups["endTag"].Value;

            // Extract the element name.  The name will be used to determine if the tags appear on their own comment lines or if they can be
            // collapsed onto a single line with their comment.  For example, the ' <summary>' tag is generally emitted on its own line whereas the '
            // <param>' tag is usually combined with the comments.
            Regex xmlElementName = new Regex(@"<(?<name>\w+).*>");
            string elementName = xmlElementName.Match(startTag).Groups["name"].Value;

            // If the tags and the content is too long to fit on a line after formatting, we'll put the tags on their own lines.  We also force a
            // certain class of tags to appear on their own line distinct from the content; it's just easier to read this way.
            if (commentBody.Length >= Properties.Settings.Default.WrapMargin || FormatCommentCommand.breakingElements.Contains(elementName))
            {
                // Write the start tag.
                outputComment.Append(prefix);
                outputComment.Append(' ');
                outputComment.Append(startTag);
                outputComment.AppendLine();

                // This will format the body of the tag to wrap at the right margin.
                FormatCommentCommand.columnPosition = 0;
                FormatCommentCommand.WrapLine(outputComment, body, prefix);
                if (FormatCommentCommand.columnPosition != 0)
                {
                    outputComment.AppendLine();
                }

                // Write the closing tag.
                outputComment.Append(prefix);
                outputComment.Append(' ');
                outputComment.Append(endTag);
                outputComment.AppendLine();
            }
            else
            {
                // If the XML tags and the line fit within the margin, we just write it to the output comment.
                outputComment.Append(commentBody.ToString());
                outputComment.AppendLine();
            }

            // Reset the wrapping parameters for the next comment line in the block.
            FormatCommentCommand.columnPosition = 0;
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

            // This command will only work when there's an active document to examine.
            if (FormatCommentCommand.environment.ActiveDocument == null)
            {
                return;
            }

            // Get the selected text from the environment.
            TextSelection selection = FormatCommentCommand.environment.ActiveDocument.Selection as TextSelection;

            // Get the start end points (round down and up to the start of a line).
            EditPoint startPoint = selection.AnchorPoint.CreateEditPoint();
            startPoint.StartOfLine();

            // The initial endpoint is one line below the start.
            EditPoint endPoint = selection.ActivePoint.CreateEditPoint();
            endPoint.StartOfLine();
            endPoint.LineDown(1);

            // If nothing is selected, then figure out what needs to be formatted by the start point up and the end point down.  As long as we
            // recognize a comment line we'll keep expanding the selection in both directions.
            if (selection.IsEmpty)
            {
                // Find the start of the block.
                while (!startPoint.AtStartOfDocument)
                {
                    if (!FormatCommentCommand.IsCommentLine(startPoint))
                    {
                        startPoint.LineDown(1);
                        break;
                    }

                    startPoint.LineUp(1);
                }

                // Find the end of the block.
                while (!endPoint.AtEndOfDocument)
                {
                    if (!FormatCommentCommand.IsCommentLine(endPoint))
                    {
                        break;
                    }

                    endPoint.LineDown(1);
                }
            }

            // This will swap the old comment for the new right-margin justified and beautified comment.
            startPoint.ReplaceText(
                endPoint,
                FormatCommentCommand.FormatCommentstring(startPoint.GetText(endPoint)),
                (int)(vsEPReplaceTextOptions.vsEPReplaceTextNormalizeNewlines | vsEPReplaceTextOptions.vsEPReplaceTextTabsSpaces));
        }
    }
}