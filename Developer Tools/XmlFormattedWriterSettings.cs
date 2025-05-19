// <copyright file="XmlFormattedWriterSettings.cs" company="Gamma Four, Inc.">
//    Copyright © 2025 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Specified a set of features use to support the XmlFormattedWriter class.
    /// </summary>
    internal class XmlFormattedWriterSettings
    {
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
                this.BaseSettings.Encoding = value;
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
                return this.BaseSettings.OmitXmlDeclaration;
            }

            set
            {
                this.BaseSettings.OmitXmlDeclaration = value;
            }
        }

        /// <summary>
        /// Gets or sets the size of a tab.
        /// </summary>
        /// <value>
        /// The size of a tab.
        /// </value>
        internal int TabSize { get; set; } = 4;

        /// <summary>
        /// Gets the base XmlWriterSettings.
        /// </summary>
        /// <value>
        /// The base XmlWriterSettings.
        /// </value>
        internal XmlWriterSettings BaseSettings { get; } = new XmlWriterSettings();
    }
}