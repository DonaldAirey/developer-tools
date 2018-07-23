// <copyright file="Element.cs" company="Dark Bond, Inc.">
//    Copyright © 2016-2017 - Dark Bond, Inc.  All Rights Reserved.
// </copyright>
// <author>Donald Roy Airey</author>
namespace DarkBond.Tools
{
    using System;

    /// <summary>
    /// An Element in the output of an XML document.
    /// </summary>
    internal class Element
    {
        /// <summary>
        /// The number of attributes associated with this element.
        /// </summary>
        private int attributeCount;

        /// <summary>
        /// The number of child elements associated with this element.
        /// </summary>
        private int childCount;

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
            this.attributeCount = 0;
            this.childCount = 0;
            this.localName = localName;
            this.prefix = prefix;
        }

        /// <summary>
        /// Gets or sets the number of attributes associated with this element.
        /// </summary>
        /// <value>
        /// The number of attributes associated with this element.
        /// </value>
        internal int AttributeCount
        {
            get { return this.attributeCount; }
            set { this.attributeCount = value; }
        }

        /// <summary>
        /// Gets or sets the number of child elements associated with this element.
        /// </summary>
        /// <value>
        /// The number of child elements associated with this element.
        /// </value>
        internal int ChildCount
        {
            get { return this.childCount; }
            set { this.childCount = value; }
        }

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