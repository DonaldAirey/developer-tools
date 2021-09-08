// <copyright file="ScrubXsdCommand.cs" company="Gamma Four, Inc.">
//    Copyright © 2018 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.Design;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Xml.Linq;
    using EnvDTE;
    using EnvDTE80;
    using GammaFour.XmlSchemaDocument;
    using GammaFour.DeveloperTools.Properties;
    using Microsoft;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Creates an executable wrapper around the IDE tool used to generate a middle tier.
    /// </summary>
    internal sealed class ScrubXsdCommand
    {
        /// <summary>
        /// Command identifier.
        /// </summary>
        private const int CommandId = 0x0007;

        // Private Static Fields
        private static XNamespace xs = "http://www.w3.org/2001/XMLSchema";

        /// <summary>
        /// The environment for the developer tools.
        /// </summary>
        private static DTE2 environment;

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        private static ScrubXsdCommand instance;

        /// <summary>
        /// The service provider for this extension.
        /// </summary>
        private static IServiceProvider serviceProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrubXsdCommand"/> class.
        /// </summary>
        /// <param name="package">Owner package.</param>
        /// <param name="commandService">Command service to add command to.</param>
        private ScrubXsdCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            // Validate the arguments.
            package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            // The environment is needed to examine and modify the active document.
            ScrubXsdCommand.serviceProvider = package as IServiceProvider;
            ScrubXsdCommand.environment = serviceProvider.GetService(typeof(DTE)) as DTE2;
            Assumes.Present(ScrubXsdCommand.environment);

            // This installs our custom command into the environment.
            commandService.AddCommand(
                new MenuCommand(this.Execute, new CommandID(DeveloperToolsPackage.CommandSet, ScrubXsdCommand.CommandId)));
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
            ScrubXsdCommand.instance = new ScrubXsdCommand(package, commandService);
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
            using (Utf8StringWriter stringWriter = new Utf8StringWriter())
            {
                try
                {
                    // Read the source document.
                    using (StringReader stringReader = new StringReader(source))
                    {
                        XmlSchemaDocument xmlSchemaDocument = new XmlSchemaDocument(stringReader.ReadToEnd());

                        // Regurgitate the schema as an XDocument
                        XElement schemaElement = new XElement(
                            ScrubXsdCommand.xs + "schema",
                            new XAttribute("id", xmlSchemaDocument.Name),
                            new XAttribute("targetNamespace", xmlSchemaDocument.TargetNamespace));
                        XDocument targetDocument = new XDocument(schemaElement);

                        // This will populate the namespace manager from the namespaces in the root element.
                        foreach (XAttribute xAttribute in xmlSchemaDocument.Root.Attributes())
                        {
                            if (xAttribute.IsNamespaceDeclaration)
                            {
                                // Place the remaining namespace attributes in the root element of the schema description.
                                schemaElement.Add(new XAttribute(xAttribute.Name, xAttribute.Value));
                            }
                        }

                        //  <xs:element name="Domain">
                        XElement dataModelElement = new XElement(ScrubXsdCommand.xs + "element", new XAttribute("name", xmlSchemaDocument.Name));

                        //  This specifies the data domain used by the REST generated code.
                        if (xmlSchemaDocument.Domain != null)
                        {
                            dataModelElement.SetAttributeValue(XmlSchemaDocument.DomainName, xmlSchemaDocument.Domain);
                        }

                        //  This flag indicates that the API uses tokens for authentication.
                        if (xmlSchemaDocument.IsSecure.HasValue)
                        {
                            dataModelElement.SetAttributeValue(XmlSchemaDocument.IsSecureName, xmlSchemaDocument.IsSecure.Value);
                        }

                        //  This flag indicates that the interface is not committed to a peristent store.
                        if (xmlSchemaDocument.IsVolatile.HasValue)
                        {
                            dataModelElement.SetAttributeValue(XmlSchemaDocument.IsVolatileName, xmlSchemaDocument.IsVolatile.Value);
                        }

                        //    <xs:complexType>
                        XElement dataModelComlexTypeElement = new XElement(ScrubXsdCommand.xs + "complexType");

                        //      <xs:choice maxOccurs="unbounded">
                        XElement dataModelChoices = new XElement(ScrubXsdCommand.xs + "choice", new XAttribute("maxOccurs", "unbounded"));

                        // This will scrub and add each of the tables to the schema in alphabetical order.
                        foreach (TableElement tableElement in xmlSchemaDocument.Tables.OrderBy(t => t.Name))
                        {
                            dataModelChoices.Add(ScrubXsdCommand.CreateTable(tableElement));
                        }

                        // The complex types that define the tables.
                        dataModelComlexTypeElement.Add(dataModelChoices);
                        dataModelElement.Add(dataModelComlexTypeElement);

                        // This will order the primary keys.
                        List<UniqueKeyElement> primaryKeyList = new List<UniqueKeyElement>();
                        foreach (TableElement tableElement in xmlSchemaDocument.Tables)
                        {
                            primaryKeyList.AddRange(tableElement.UniqueKeys);
                        }

                        foreach (UniqueKeyElement uniqueConstraintSchema in primaryKeyList.OrderBy(pke => pke.Name))
                        {
                            dataModelElement.Add(CreateUniqueKey(uniqueConstraintSchema));
                        }

                        // This will order the foreign primary keys.
                        List<ForeignKeyElement> foreignKeyList = new List<ForeignKeyElement>();
                        foreach (TableElement tableElement in xmlSchemaDocument.Tables)
                        {
                            foreignKeyList.AddRange(tableElement.ForeignKeys);
                        }

                        foreach (ForeignKeyElement foreignConstraintSchema in foreignKeyList.OrderBy(fke => fke.Name))
                        {
                            dataModelElement.Add(CreateForeignKey(foreignConstraintSchema));
                        }

                        // Add the data model element to the document.
                        schemaElement.Add(dataModelElement);

                        // Save the regurgitated output.
                        targetDocument.Save(stringWriter);
                    }
                }
                catch (Exception exception)
                {
                    // Show a message box to prove we were here
                    VsShellUtilities.ShowMessageBox(
                        ScrubXsdCommand.serviceProvider,
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
        /// Creates the element that describes a table.
        /// </summary>
        /// <param name="tableElement">A description of a table.</param>
        /// <returns>An element that represents a table in an XML Schema document.</returns>
        private static XElement CreateTable(TableElement tableElement)
        {
            //        <xs:element name="Account">
            XElement xElement = new XElement(ScrubXsdCommand.xs + "element", new XAttribute("name", tableElement.Name));

            // gfdata:isVolatile="true"
            if (tableElement.IsVolatile)
            {
                xElement.SetAttributeValue(XmlSchemaDocument.IsVolatileName, true);
            }

            // gfdata:verbs="Delete,Get,Put"
            string verbs = string.Empty;
            foreach (Verb verb in tableElement.Verbs)
            {
                if (!string.IsNullOrEmpty(verbs))
                {
                    verbs += ",";
                }

                verbs += verb.ToString();
            }

            if (!string.IsNullOrEmpty(verbs))
            {
                xElement.Add(new XAttribute(XmlSchemaDocument.VerbsName, verbs));
            }

            //           <xs:complexType>
            XElement complexTypeElement = new XElement(ScrubXsdCommand.xs + "complexType");
            xElement.Add(complexTypeElement);

            //            <xs:sequence>
            XElement sequenceElement = new XElement(ScrubXsdCommand.xs + "sequence");
            complexTypeElement.Add(sequenceElement);

            // Emit each of the columns.
            foreach (ColumnElement columnElement in tableElement.Columns.OrderBy(c => c.Name))
            {
                if (!columnElement.IsRowVersion)
                {
                    sequenceElement.Add(CreateColumn(columnElement));
                }
            }

            // This is a complete XML Schema description of a table.
            return xElement;
        }

        /// <summary>
        /// Creates the element that describes a column in a table.
        /// </summary>
        /// <param name="columnElement">A description of a column.</param>
        /// <returns>An element that can be used in an XML Schema document to describe a column.</returns>
        private static XElement CreateColumn(ColumnElement columnElement)
        {
            string dataType = string.Empty;
            string xmlType = string.Empty;
            switch (columnElement.ColumnType.FullName)
            {
                case "System.Object":

                    xmlType = "xs:anyType";
                    break;

                case "System.Int32":

                    xmlType = "xs:int";
                    break;

                case "System.Int64":

                    xmlType = "xs:long";
                    break;

                case "System.Float":

                    xmlType = "xs:double";
                    break;

                case "System.Double":

                    xmlType = "xs:double";
                    break;

                case "System.Decimal":

                    xmlType = "xs:decimal";
                    break;

                case "System.Boolean":

                    xmlType = "xs:boolean";
                    break;

                case "System.String":

                    xmlType = "xs:string";
                    break;

                case "System.DateTime":

                    xmlType = "xs:dateTime";
                    break;

                case "System.Byte[]":

                    xmlType = "xs:base64Binary";
                    break;

                default:

                    dataType = columnElement.ColumnType.FullName;
                    xmlType = "xs:anyType";
                    break;
            }

            //                          <xs:element name="UserId" msdata:DataType="System.Guid, mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
            //                            msprop:Generator_UserColumnName="UserId" msprop:Generator_ColumnVarNameInTable="columnUserId" msprop:Generator_ColumnPropNameInRow="UserId"
            //                            msprop:Generator_ColumnPropNameInTable="UserIdColumn" type="xs:string" minOccurs="0" />
            XElement xElement = new XElement(ScrubXsdCommand.xs + "element", new XAttribute("name", columnElement.Name));

            // Microsoft uses a custom decoration to describe data types that are not part of the cannon XML Schema datatypes.
            if (dataType != string.Empty)
            {
                xElement.Add(new XAttribute(XmlSchemaDocument.DataTypeName, dataType));
            }

            // This attribute controls whether the column is autoincremented by the database server.
            if (columnElement.IsAutoIncrement)
            {
                xElement.Add(new XAttribute(XmlSchemaDocument.AutoIncrementName, true));
            }

            // Emit the column's type.
            if (!columnElement.HasSimpleType)
            {
                xElement.Add(new XAttribute("type", xmlType));
            }
            else
            {
                if (columnElement.ColumnType.FullName == typeof(string).FullName)
                {
                    //                <xs:simpleType>
                    //                  <xs:restriction base="xs:string">
                    //                    <xs:maxLength value="128" />
                    //                  </xs:restriction>
                    //                </xs:simpleType>
                    XElement restrictionElement = new XElement(
                        ScrubXsdCommand.xs + "restriction",
                        new XAttribute("base", xmlType),
                        new XElement(ScrubXsdCommand.xs + "maxLength", new XAttribute("value", columnElement.MaximumLength)));
                    xElement.Add(new XElement(ScrubXsdCommand.xs + "simpleType", restrictionElement));
                }

                if (columnElement.ColumnType.FullName == typeof(decimal).FullName)
                {
                    //                <xs:simpleType>
                    //                  <xs:restriction base="xs:decimal">
                    //                    <xs:fractionDigits value="6" />
                    //                    <xs:totalDigits value="18" />
                    //                  </xs:restriction>
                    //                </xs:simpleType>
                    XElement restrictionElement = new XElement(
                        ScrubXsdCommand.xs + "restriction",
                        new XAttribute("base", xmlType),
                        new XElement(ScrubXsdCommand.xs + "fractionDigits", new XAttribute("value", columnElement.FractionDigits)),
                        new XElement(ScrubXsdCommand.xs + "totalDigits", new XAttribute("value", columnElement.TotalDigits)));
                    xElement.Add(new XElement(ScrubXsdCommand.xs + "simpleType", restrictionElement));
                }
            }

            // An optional column is identified with a 'minOccurs=0' attribute.
            if (columnElement.ColumnType.IsNullable)
            {
                xElement.Add(new XAttribute("minOccurs", 0));
            }

            // Provide an explicit default value for all column elements.
            if (!columnElement.ColumnType.IsNullable && columnElement.DefaultValue != null)
            {
                xElement.Add(new XAttribute("default", columnElement.DefaultValue.ToString().ToLower()));
            }

            // This describes the column of a table.
            return xElement;
        }

        /// <summary>
        /// Creates the element that describes a unique constraint.
        /// </summary>
        /// <param name="uniqueConstraintSchema">A description of a unique constraint.</param>
        /// <returns>An element that can be used in an XML Schema document to describe a unique constraint.</returns>
        private static XElement CreateUniqueKey(UniqueKeyElement uniqueConstraintSchema)
        {
            //    <xs:unique name="AccessControlKey" msdata:PrimaryKey="true">
            //      <xs:selector xpath=".//mstns:AccessControl" />
            //      <xs:field xpath="mstns:UserId" />
            //      <xs:field xpath="mstns:EntityId" />
            //    </xs:unique>
            XElement uniqueElement = new XElement(
                ScrubXsdCommand.xs + "unique",
                new XAttribute("name", uniqueConstraintSchema.Name));

            if (uniqueConstraintSchema.IsPrimaryKey)
            {
                uniqueElement.Add(new XAttribute(XmlSchemaDocument.IsPrimaryKeyName, true));
            }

            uniqueElement.Add(
                new XElement(
                    ScrubXsdCommand.xs + "selector",
                    new XAttribute("xpath", string.Format(".//mstns:{0}", uniqueConstraintSchema.Table.Name))));

            foreach (ColumnReferenceElement columnReferenceElement in uniqueConstraintSchema.Columns.OrderBy(c => c.Name))
            {
                ColumnElement columnElement = columnReferenceElement.Column;
                uniqueElement.Add(
                    new XElement(
                        ScrubXsdCommand.xs + "field",
                        new XAttribute("xpath", string.Format("mstns:{0}", columnElement.Name))));
            }

            // This describes a unique constraint on a table.
            return uniqueElement;
        }

        /// <summary>
        /// Creates the element that describes a foreign constraint.
        /// </summary>
        /// <param name="foreignKeyElement">A description of a foreign constraint.</param>
        /// <returns>An element that can be used in an XML Schema document to describe a foreign constraint.</returns>
        private static XElement CreateForeignKey(ForeignKeyElement foreignKeyElement)
        {
            //    <xs:keyref name="FK_Entity_AccessControl" refer="EntityKey" msprop:rel_Generator_UserRelationName="FK_Entity_AccessControl"
            //        msprop:rel_Generator_RelationVarName="relationFK_Entity_AccessControl" msprop:rel_Generator_UserChildTable="AccessControl"
            //        msprop:rel_Generator_UserParentTable="Entity" msprop:rel_Generator_ParentPropName="EntityRow"
            //        msprop:rel_Generator_ChildPropName="GetAccessControlRows">
            //      <xs:selector xpath=".//mstns:AccessControl" />
            //      <xs:field xpath="mstns:EntityId" />
            //    </xs:keyref>
            XElement foreignElement = new XElement(
                    ScrubXsdCommand.xs + "keyref",
                    new XAttribute("name", foreignKeyElement.Name),
                    new XAttribute("refer", foreignKeyElement.UniqueKey.Name));

            foreignElement.Add(
                new XElement(
                    ScrubXsdCommand.xs + "selector",
                    new XAttribute("xpath", string.Format(".//mstns:{0}", foreignKeyElement.Table.Name))));

            foreach (ColumnReferenceElement columnReferenceElement in foreignKeyElement.Columns.OrderBy(c => c.Name))
            {
                ColumnElement columnElement = columnReferenceElement.Column;
                foreignElement.Add(
                    new XElement(
                        ScrubXsdCommand.xs + "field",
                        new XAttribute("xpath", string.Format("mstns:{0}", columnElement.Name))));
            }

            // This describes a foreign constraint on a table.
            return foreignElement;
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

            string name = ScrubXsdCommand.environment.ActiveDocument.FullName;
            string extension = Path.GetExtension(ScrubXsdCommand.environment.ActiveDocument.FullName).ToUpperInvariant();
            if (extension == ".XSD")
            {
                // Get the selected text from the environment and make a note of where we are in the document.
                TextSelection selection = ScrubXsdCommand.environment.ActiveDocument.Selection as TextSelection;
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
                    ScrubXsdCommand.ScrubDocument(startPoint.GetText(endPoint), extension),
                    (int)(vsEPReplaceTextOptions.vsEPReplaceTextNormalizeNewlines | vsEPReplaceTextOptions.vsEPReplaceTextTabsSpaces));

                // There will likely be some dramatic changes to the format of the document.  This will restore the cursor to where it was in the
                // document before we beautified it.
                selection.MoveToLineAndOffset(line, offset);
            }
        }

        /// <summary>
        /// A class to force the XML encoding to be UTF-8.
        /// </summary>
        private class Utf8StringWriter : StringWriter
        {
            public override Encoding Encoding
            {
                get { return Encoding.UTF8; }
            }
        }
    }
}