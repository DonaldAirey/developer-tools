// <copyright file="XmlFormattedWriter.cs" company="Gamma Four, Inc.">
//    Copyright © 2025 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// An XML text writer that beautifies the XML stream.
    /// </summary>
    internal class XmlFormattedWriter : XmlWriter
    {
        /// <summary>
        /// A stack of states used to track the current element that is being written.
        /// </summary>
        private Stack<Element> elementStack;

        /// <summary>
        /// A stack of states used to track the state of the writer as nodes are written.
        /// </summary>
        private Stack<WriteState> stateStack;

        /// <summary>
        /// A stream used to write the XML.
        /// </summary>
        private TextWriter textWriter;

        /// <summary>
        /// Settings for writing the XML.
        /// </summary>
        private XmlFormattedWriterSettings xmlFormattedWriterSettings;

        /// <summary>
        /// Initializes a new instance of the <see cref="XmlFormattedWriter"/> class.
        /// </summary>
        /// <param name="textWriter">The <see cref="TextWriter"/> to which you want to write.</param>
        /// <param name="xmlFormattedWriterSettings">
        /// The <see cref="XmlFormattedWriterSettings"/> object used to configure the new <see cref="TextWriter"/> instance.  If this is null, a
        /// <see cref="XmlFormattedWriterSettings"/> with default settings is used.
        /// </param>
        private XmlFormattedWriter(TextWriter textWriter, XmlFormattedWriterSettings xmlFormattedWriterSettings)
        {
            // Initialize the object.
            this.textWriter = textWriter;
            this.elementStack = new Stack<Element>();
            this.stateStack = new Stack<WriteState>();
            this.xmlFormattedWriterSettings = xmlFormattedWriterSettings;
        }

        /// <summary>
        /// Gets the state of the writer.
        /// </summary>
        /// <value>
        /// The state of the writer.
        /// </value>
        public override WriteState WriteState
        {
            get { return this.stateStack.Peek(); }
        }

        /// <summary>
        /// Gets a string that can be used for the current indentation level when writing to the XML stream.
        /// </summary>
        /// <value>
        /// A string that can be used for the current indentation level when writing to the XML stream.
        /// </value>
        private string CurrentIndent
        {
            get
            {
                int totalSpace;
                int tabSize = this.xmlFormattedWriterSettings.TabSize;
                StringBuilder indentText = new StringBuilder();

                // The state determines how the visual alignment is handled.
                switch (this.WriteState)
                {
                case WriteState.Attribute:

                    // This creates a string that will align all the attributes under the first attribute in an element.
                    totalSpace = ((this.elementStack.Count - 1) * tabSize) + this.elementStack.Peek().QualifiedName.Length + 2;
                    for (int tabCount = 0; tabCount < totalSpace / tabSize; tabCount++)
                    {
                        indentText.Append("\t");
                    }

                    for (int spaceCount = 0; spaceCount < totalSpace % tabSize; spaceCount++)
                    {
                        indentText.Append(" ");
                    }

                    break;

                case WriteState.Content:

                    // This lines up the content of an element at the same outline level.
                    totalSpace = this.elementStack.Count * tabSize;
                    for (int tabCount = 0; tabCount < totalSpace / tabSize; tabCount++)
                    {
                        indentText.Append("\t");
                    }

                    break;
                }

                // This string can be used when writing the stream to visually align the nodes in an XML document.
                return indentText.ToString();
            }
        }

        /// <summary>
        /// Closes this stream and the underlying stream.
        /// </summary>
        public override void Close()
        {
            // Close the stream and flushes the contents.
            this.textWriter.Close();
        }

        /// <summary>
        /// Flushes whatever is in the buffer to the underlying streams and also flushes the underlying stream.
        /// </summary>
        public override void Flush()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Returns the closest prefix defined in the current namespace scope for the namespace URI.
        /// </summary>
        /// <param name="ns">The namespace URI whose prefix you want to find.</param>
        /// <returns>The lookup prefix for the given namespace.</returns>
        public override string LookupPrefix(string ns)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Encodes the specified binary bytes as Base64 and writes out the resulting text.
        /// </summary>
        /// <param name="buffer">Byte array to encode.</param>
        /// <param name="index">The position in the buffer indicating the start of the bytes to write.</param>
        /// <param name="count">The number of bytes to write.</param>
        public override void WriteBase64(byte[] buffer, int index, int count)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Writes out a &gt;![CDATA[...]]&lt; block containing the specified text.
        /// </summary>
        /// <param name="text">The text to place inside the CDATA block.</param>
        public override void WriteCData(string text)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Forces the generation of a character entity for the specified Unicode character value.
        /// </summary>
        /// <param name="ch">The Unicode character for which to generate a character entity.</param>
        public override void WriteCharEntity(char ch)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Writes text one buffer at a time.
        /// </summary>
        /// <param name="buffer">Character array containing the text to write.</param>
        /// <param name="index">The position in the buffer indicating the start of the text to write.</param>
        /// <param name="count">The number of characters to write.</param>
        public override void WriteChars(char[] buffer, int index, int count)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Writes out a comment <!--...--> containing the specified text.
        /// </summary>
        /// <param name="text">Text to place inside the comment.</param>
        public override void WriteComment(string text)
        {
            switch (this.WriteState)
            {
            case WriteState.Content:

                // If the writer has already started writing content, then align the comment under the siblings nodes in the output.
                this.textWriter.Write("{0}<!--{1}-->", this.CurrentIndent, text);
                break;

            case WriteState.Element:

                // If the writer hasn't stared writing content yet, then terminate the start element tag and align the comment under the siblings in
                // the output stream.
                this.TerminateStartElement();
                this.textWriter.Write("{0}<!--{1}-->", this.CurrentIndent, text);
                break;
            }
        }

        /// <summary>
        /// Writes the DOCTYPE declaration with the specified name and optional attributes.
        /// </summary>
        /// <param name="name">The name of the DOCTYPE.  This must be non-empty.</param>
        /// <param name="pubid">The internal identifier of this document.</param>
        /// <param name="sysid">The SystemID of this document.</param>
        /// <param name="subset">If non-null it writes where subset is replaced with the value of this argument.</param>
        public override void WriteDocType(string name, string pubid, string sysid, string subset)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// closes the previous WriteStartAttribute call.
        /// </summary>
        public override void WriteEndAttribute()
        {
            // Return the writer to the previous state.
            this.stateStack.Pop();
        }

        /// <summary>
        /// Closes any open elements or attributes and puts the writer back in the Start state.
        /// </summary>
        public override void WriteEndDocument()
        {
            // Return the writer to the original state.
            this.stateStack.Pop();

            if (this.stateStack.Count != 0)
            {
                throw new InvalidOperationException("The state stack was not balanced at the end of writing.");
            }
        }

        /// <summary>
        /// Closes one element and pops the corresponding namespace scope.
        /// </summary>
        public override void WriteEndElement()
        {
            // Pop the element off the stack now that we're done writing it.
            Element element = this.elementStack.Pop();

            // This terminates the element in the output stream.
            this.textWriter.Write("/>");

            // Return the writer to the state it had before we wrote this element.
            this.stateStack.Pop();
        }

        /// <summary>
        /// writes out an entity reference as a name.
        /// </summary>
        /// <param name="name">The name of the entity reference.</param>
        public override void WriteEntityRef(string name)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Closes one element and pops the corresponding namespace scope.
        /// </summary>
        public override void WriteFullEndElement()
        {
            // Pop the element off the stack now that we're done writing it.
            Element element = this.elementStack.Pop();

            // This element had content that couldn't be written in a single tag.  This writes the end tag to the XML output stream.
            if (element.ChildCount == 0)
            {
                this.textWriter.Write("</{1}>", this.CurrentIndent, element.QualifiedName);
            }
            else
            {
                this.textWriter.Write("{0}</{1}>", this.CurrentIndent, element.QualifiedName);
            }

            // Return the writer to the state it had before we wrote this element.
            this.stateStack.Pop();
        }

        /// <summary>
        /// Writes out a processing instruction with a space between the name and text as follows: &gt;?name text?&lt;.
        /// </summary>
        /// <param name="name">The name of the processing instruction.</param>
        /// <param name="text">The text to include in the processing instruction.</param>
        public override void WriteProcessingInstruction(string name, string text)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Writes raw markup manually.
        /// </summary>
        /// <param name="data">string containing the text to write.</param>
        public override void WriteRaw(string data)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Writes raw markup manually.
        /// </summary>
        /// <param name="buffer">Character array containing the text to write.</param>
        /// <param name="index">The position within the buffer indicating the start of the text to write.</param>
        /// <param name="count">The number of characters to write.</param>
        public override void WriteRaw(char[] buffer, int index, int count)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// writes the start of an attribute with the specified prefix, local name, and namespace URI.
        /// </summary>
        /// <param name="prefix">The namespace prefix of the attribute.</param>
        /// <param name="localName">The local name of the attribute.</param>
        /// <param name="namespaceUri">The namespace URI for the attribute.</param>
        public override void WriteStartAttribute(string prefix, string localName, string namespaceUri)
        {
            // This indicates that the writer is processing an attribute.
            this.stateStack.Push(WriteState.Attribute);

            // This is the qualified name for the attribute.
            string qualifiedName = string.IsNullOrEmpty(prefix) ? localName : prefix + ":" + localName;

            // The attributes are aligned visually under the first attribute.
            if (this.elementStack.Peek().AttributeCount == 0)
            {
                this.textWriter.Write(" {0}=", qualifiedName);
            }
            else
            {
                this.textWriter.WriteLine();
                this.textWriter.Write("{0}{1}=", this.CurrentIndent, qualifiedName);
            }

            // This keeps track of the number of elements written.  Currently only the first one is significant.
            this.elementStack.Peek().AttributeCount = this.elementStack.Peek().AttributeCount + 1;
        }

        /// <summary>
        /// Writes the XML declaration.
        /// </summary>
        /// <param name="standalone">If true, it writes "standalone=yes"; if false, it writes "standalone=no".</param>
        public override void WriteStartDocument(bool standalone)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Writes the XML declaration with the version "1.0" and the standalone attribute.
        /// </summary>
        public override void WriteStartDocument()
        {
            // Write the XML declaration unless we've been specifically asked to inhibit it.
            if (!this.xmlFormattedWriterSettings.OmitXmlDeclaration)
            {
                this.textWriter.WriteLine("<?xml version=\"1.0\"?>");
            }

            // This indicates the initial state of the write operation.
            this.stateStack.Push(WriteState.Start);
        }

        /// <summary>
        /// Writes the specified start tag and associates it with the given namespace and prefix.
        /// </summary>
        /// <param name="prefix">The namespace prefix of the element.</param>
        /// <param name="localName">The local name of the element.</param>
        /// <param name="ns">The namespace URI to associate with the element.</param>
        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            // If the start of an element is already being written when the writer came across a child element, then the parent start tag must be
            // terminated properly before the rest of the stream is processed.  Note that the state changes from processing an element start tag to
            // processing the content.
            switch (this.WriteState)
            {
            case WriteState.Element:
                this.TerminateStartElement();
                this.textWriter.WriteLine();
                break;
            }

            // This keeps track of the number of child elements.
            if (this.elementStack.Count != 0)
            {
                this.elementStack.Peek().ChildCount = this.elementStack.Peek().ChildCount + 1;
            }

            // Create a new element and write it out to the stream at the proper indentation level.
            Element element = new Element(localName, prefix);
            this.textWriter.Write("{0}<{1}", this.CurrentIndent, element.QualifiedName);

            // The element will be needed again when all the content has been processed.
            this.elementStack.Push(element);

            // This indicates that we're processing an element start tag.
            this.stateStack.Push(WriteState.Element);
        }

        /// <summary>
        /// Writes the given text content.
        /// </summary>
        /// <param name="text">The text to write.</param>
        public override void WriteString(string text)
        {
            // Validate the 'text' argument
            if (text == null)
            {
                throw new ArgumentNullException("text");
            }

            // Different states have different tokens for writing the text.
            switch (this.WriteState)
            {
            case WriteState.Attribute:

                // Attribute text is written in quotes.
                this.textWriter.Write("\"{0}\"", text);
                break;

            case WriteState.Element:

                this.TerminateStartElement();
                this.textWriter.Write(text.Replace(" ", string.Empty).Replace("\t", string.Empty));
                break;

            default:

                // Text inside of an element needs to terminate the "start element" tag.
                this.textWriter.Write(text.Replace(" ", string.Empty).Replace("\t", string.Empty));
                break;
            }
        }

        /// <summary>
        /// Generates and writes the surrogate character entity for the surrogate character pair.
        /// </summary>
        /// <param name="lowChar">The low surrogate.  This must be a value between 0xDC00 and 0xDFFF.</param>
        /// <param name="highChar">The high surrogate.  This must be a value between 0xD800 and 0xDBFF.</param>
        public override void WriteSurrogateCharEntity(char lowChar, char highChar)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Writes out the given white space.
        /// </summary>
        /// <param name="ws">The string of white space characters.</param>
        public override void WriteWhitespace(string ws)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Creates a new <see cref="XmlFormattedWriter"/> instance using the <see cref="TextWriter"/> and <see cref="XmlFormattedWriterSettings"/>
        /// objects.
        /// </summary>
        /// <param name="textWriter">The <see cref="TextWriter"/> to which you want to write.</param>
        /// <param name="xmlFormattedWriterSettings">
        /// The <see cref="XmlFormattedWriterSettings"/> object used to configure the new <see cref="TextWriter"/> instance.  If this is null, a
        /// <see cref="XmlFormattedWriterSettings"/> with default settings is used.
        /// </param>
        /// <returns>An <see cref="XmlFormattedWriter"/> object.</returns>
        internal static XmlWriter Create(TextWriter textWriter, XmlFormattedWriterSettings xmlFormattedWriterSettings)
        {
            XmlFormattedWriter xmlFormattedTextWriter = new XmlFormattedWriter(textWriter, xmlFormattedWriterSettings);
            return XmlWriter.Create(xmlFormattedTextWriter, xmlFormattedWriterSettings.BaseSettings);
        }

        /// <summary>
        /// Completes the processing of a start element.
        /// </summary>
        private void TerminateStartElement()
        {
            // Terminate the start element tag and indicate that the writer is now writing the content of the element.
            this.textWriter.Write(">");
            this.stateStack.Pop();
            this.stateStack.Push(WriteState.Content);
        }
    }
}