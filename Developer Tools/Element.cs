// <copyright file="Element.cs" company="Gamma Four, Inc.">
//    Copyright © 2025 - Gamma Four, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace GammaFour.DeveloperTools
{
    /// <summary>
    /// An Element in the output of an XML document.
    /// </summary>
    internal class Element
    {
        /// <summary>
        /// The local name of the element.
        /// </summary>
        private string localName;

        /// <summary>
        /// The prefix that has been assigned to the namespace.
        /// </summary>
        private string prefix;

        /// <summary>
        /// Initializes a new instance of the <see cref="Element"/> class.
        /// </summary>
        /// <param name="localName">The local name of the element.</param>
        /// <param name="prefix">The namespace prefix of the element.</param>
        internal Element(string localName, string prefix)
        {
            // Initialize the object
            this.AttributeCount = 0;
            this.ChildCount = 0;
            this.localName = localName;
            this.prefix = prefix;
        }

        /// <summary>
        /// Gets or sets the number of attributes associated with this element.
        /// </summary>
        /// <value>
        /// The number of attributes associated with this element.
        /// </value>
        internal int AttributeCount { get; set; }

        /// <summary>
        /// Gets or sets the number of child elements associated with this element.
        /// </summary>
        /// <value>
        /// The number of child elements associated with this element.
        /// </value>
        internal int ChildCount { get; set; }

        /// <summary>
        /// Gets the fully qualified name of the element.
        /// </summary>
        /// <value>
        /// The fully qualified name of the element.
        /// </value>
        internal string QualifiedName
        {
            get
            {
                return string.IsNullOrEmpty(this.prefix) ? this.localName : this.prefix + ":" + this.localName;
            }
        }
    }
}