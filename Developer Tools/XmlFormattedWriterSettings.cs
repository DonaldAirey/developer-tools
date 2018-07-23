// <copyright file="XmlFormattedWriterSettings.cs" company="Dark Bond, Inc.">
//    Copyright © 2016-2017 - Dark Bond, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace DarkBond.Tools
{
    using System;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Specified a set of features use to support the XmlFormattedWriter class.
    /// </summary>
    internal class XmlFormattedWriterSettings
    {
        /// <summary>
        /// The size of the tabs.
        /// </summary>
        private int tabSize = 4;

        /// <summary>
        /// The encapsulated XmlWriterSettings object.
        /// </summary>
        private XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();

        /// <summary>
        /// Sets the type of text encoding to use.
        /// </summary>
        /// <value>
        /// The type of text encoding to use.
        /// </value>
        internal Encoding Encoding
        {
            set
            {
                this.xmlWriterSettings.Encoding = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to write an XML declaration.
        /// </summary>
        /// <value>
        /// A value indicating whether to write an XML declaration.
        /// </value>
        internal bool OmitXmlDeclaration
        {
            get
            {
                return this.xmlWriterSettings.OmitXmlDeclaration;
            }

            set
            {
                this.xmlWriterSettings.OmitXmlDeclaration = value;
            }
        }

        /// <summary>
        /// Gets or sets the size of a tab.
        /// </summary>
        /// <value>
        /// The size of a tab.
        /// </value>
        internal int TabSize
        {
            get
            {
                return this.tabSize;
            }

            set
            {
                this.tabSize = value;
            }
        }

        /// <summary>
        /// Gets the base XmlWriterSettings.
        /// </summary>
        /// <value>
        /// The base XmlWriterSettings.
        /// </value>
        internal XmlWriterSettings BaseSettings
        {
            get
            {
                return this.xmlWriterSettings;
            }
        }
    }
}