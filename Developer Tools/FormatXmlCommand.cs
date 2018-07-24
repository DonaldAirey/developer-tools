// <copyright file="FormatXmlCommand.cs" company="Gamma Four, Inc.">
//    Copyright © 2018 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.Design;
    using System.Globalization;
    using System.IO;
    using System.Text;
    using System.Xml;
    using System.Xml.Linq;
    using EnvDTE;
    using EnvDTE80;
    using GammaFour.DeveloperTools.Properties;
    using Microsoft;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class FormatXmlCommand
    {
        /// <summary>
        /// Command identifier.
        /// </summary>
        private const int CommandId = 0x0006;

        /// <summary>
        /// The extensions that are recognized as XML type documents.
        /// </summary>
        private static readonly List<string> XmlExtensions = new List<string>
        {
            ".CONFIG",
            ".XML",
            ".XAML",
            ".VSCT",
            ".WXS"
        };

        /// <summary>
        /// The environment for the developer tools.
        /// </summary>
        private static DTE2 environment;

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        private static FormatXmlCommand instance;

        /// <summary>
        /// The service provider for this extension.
        /// </summary>
        private static IServiceProvider serviceProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="FormatXmlCommand"/> class.
        /// </summary>
        /// <param name="package">Owner package.</param>
        /// <param name="commandService">Command service to add command to.</param>
        private FormatXmlCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            // Validate the arguments.
            package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            // The environment is needed to examine and modify the active document.
            FormatXmlCommand.serviceProvider = package as IServiceProvider;
            FormatXmlCommand.environment = serviceProvider.GetService(typeof(DTE)) as DTE2;
            Assumes.Present(FormatXmlCommand.environment);

            // This installs our custom command into the environment.
            commandService.AddCommand(
                new MenuCommand(this.Execute, new CommandID(DeveloperToolsPackage.CommandSet, FormatXmlCommand.CommandId)));
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
            FormatXmlCommand.instance = new FormatXmlCommand(package, commandService);
        }

        /// <summary>
        /// Sorts the attributes of an element and all it's children in alphabetical order.
        /// </summary>
        /// <param name="sourceNode">The original XElement.</param>
        /// <param name="targetContainer">The target XElement where the newly sorted node is placed.</param>
        private static void RecurseIntoDocument(XNode sourceNode, XContainer targetContainer)
        {
            // Convert the generic node to a specific type.
            XElement sourceElement = sourceNode as XElement;
            if (sourceElement != null)
            {
                // Arrange all the attributes in alphabetical order.
                SortedDictionary<string, XAttribute> sortedDictionary = new SortedDictionary<string, XAttribute>();
                foreach (XAttribute attribute in sourceElement.Attributes())
                {
                    sortedDictionary.Add(attribute.Name.LocalName, attribute);
                }

                // Create a new target element from the source and remove the existing attributes.  They'll be replaced with the ordered list.
                XElement element = new XElement(sourceElement);
                element.RemoveAll();

                // Replace the attributes in alphabetical order.
                foreach (KeyValuePair<string, XAttribute> keyValuePair in sortedDictionary)
                {
                    element.Add(keyValuePair.Value);
                }

                // Add it back to the target document.
                targetContainer.Add(element);

                // Recurse into each of the child elements.
                foreach (XNode childNode in sourceElement.Nodes())
                {
                    FormatXmlCommand.RecurseIntoDocument(childNode, element);
                }
            }
            else
            {
                // Pass all other nodes into the output without alterations.
                targetContainer.Add(sourceNode);
            }
        }

        /// <summary>
        /// Orders and beautifies an XML document.
        /// </summary>
        /// <param name="source">The source XML document to order and beautify.</param>
        /// <param name="extension">The extension.  Used to determine tab spacing.</param>
        /// <returns>The text of the scrubbed and beautified document.</returns>
        private static string ScrubDocument(string source, string extension)
        {
            // Used to write the final document in the form of a long string.
            using (StringWriter stringWriter = new StringWriter(CultureInfo.InvariantCulture))
            {
                try
                {
                    // Read the source document.
                    using (StringReader stringReader = new StringReader(source))
                    {
                        // Create an XML document from the source string.
                        XDocument sourceDocument = XDocument.Load(stringReader, LoadOptions.PreserveWhitespace);

                        // Create a new target document.
                        XDocument targetDocument = new XDocument();

                        // Copy the declaration.
                        if (sourceDocument.Declaration != null)
                        {
                            targetDocument.Declaration = new XDeclaration(sourceDocument.Declaration);
                        }

                        // This will order all the attributes in the document in alphabetical order.  Note that the elements can't be similarly
                        // ordered because this could produce forward reference errors.
                        FormatXmlCommand.RecurseIntoDocument(sourceDocument.Root, targetDocument);

                        // This is used to make special consideration for XALM files.
                        bool isXaml = extension == ".XAML";

                        // Beautify and save the target document when it has been ordered.
                        XmlFormattedWriterSettings xmlFormattedWriterSettings = new XmlFormattedWriterSettings
                        {
                            Encoding = Encoding.UTF8,
                            OmitXmlDeclaration = isXaml,
                            TabSize = isXaml ? 4 : 2
                        };
                        XmlWriter xmlWriter = XmlFormattedWriter.Create(stringWriter, xmlFormattedWriterSettings);
                        targetDocument.WriteTo(xmlWriter);
                        xmlWriter.Close();
                    }
                }
                catch (Exception exception)
                {
                    // Show a message box to prove we were here
                    VsShellUtilities.ShowMessageBox(
                        FormatXmlCommand.serviceProvider,
                        exception.Message,
                        Resources.EditorExtensionsTitle,
                        OLEMSGICON.OLEMSGICON_CRITICAL,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }

                // The result of beautifying a long string of XML is another long string that is scrubbed and beautified.  This string will be used
                // to replace the original content of the module.
                return stringWriter.ToString();
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
            string name = FormatXmlCommand.environment.ActiveDocument.FullName;
            string extension = Path.GetExtension(FormatXmlCommand.environment.ActiveDocument.FullName).ToUpperInvariant();
            if (FormatXmlCommand.XmlExtensions.Contains(extension))
            {
                // Get the selected text from the environment and make a note of where we are in the document.
                TextSelection selection = FormatXmlCommand.environment.ActiveDocument.Selection as TextSelection;
                int line = selection.ActivePoint.Line;
                int offset = selection.ActivePoint.LineCharOffset;

                // Get the start end points (round down and up to the start of a line).
                EditPoint startPoint = selection.AnchorPoint.CreateEditPoint();
                startPoint.StartOfDocument();

                // The initial endpoint is one line below the start.
                EditPoint endPoint = selection.ActivePoint.CreateEditPoint();
                endPoint.EndOfDocument();

                // This will replace the XML in the current module with the scrubbed and beautified XML.
                startPoint.ReplaceText(
                    endPoint,
                    FormatXmlCommand.ScrubDocument(startPoint.GetText(endPoint), extension),
                    (int)(vsEPReplaceTextOptions.vsEPReplaceTextNormalizeNewlines | vsEPReplaceTextOptions.vsEPReplaceTextTabsSpaces));

                // There will likely be some dramatic changes to the format of the document.  This will restore the cursor to where it was in the
                // document before we beautified it.
                selection.MoveToLineAndOffset(line, offset);
            }
        }
    }
}