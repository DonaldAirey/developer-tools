// <copyright file="FormatXmlCommand.cs" company="Dark Bond, Inc.">
//    Copyright © 2016-2017 - Dark Bond, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace DarkBond.Tools
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.Design;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.IO;
    using System.Text;
    using System.Windows.Threading;
    using System.Xml;
    using System.Xml.Linq;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using Properties;

    /// <summary>
    /// Beautifies an XML documents by alphabetizing and aligning attributes vertically.
    /// </summary>
    internal static class FormatXmlCommand
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
        /// The service provider for this extension.
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
            FormatXmlCommand.serviceProvider = package as IServiceProvider;
            FormatXmlCommand.environment = FormatXmlCommand.serviceProvider.GetService(typeof(DTE)) as DTE2;
            Assumes.Present(FormatXmlCommand.environment);

            // This installs our custom command into the environment.
            OleMenuCommandService oleMenuCommandService =
                FormatXmlCommand.serviceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (oleMenuCommandService != null)
            {
                oleMenuCommandService.AddCommand(
                    new MenuCommand(FormatXmlCommand.ExecuteCommand, new CommandID(DeveloperToolsPackage.CommandSet, CommandId)));
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
        [SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times", Justification = "Can't make this error go away through restructuring")]
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "General Exception is desired in this scenario.")]
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
    }
}