using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using ObjectEx.Utilities;
using PptxXML.Enums;
using PptxXML.Extensions;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represents an element on a slide.
    /// </summary>
    public abstract class Element
    {
        #region Fields

        public OpenXmlCompositeElement XmlCompositeElement { get; set; } //TODO: remove public setter
        private bool? _isPlaceholder;
        private bool? _hidden;
        private int _id;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets or sets identifier.
        /// </summary>
        public int Id
        {
            get
            {
                if (_id == 0)
                {
                    var (id, hidden) = XmlCompositeElement.GetNvPrValues();
                    _id = id;
                    _hidden = hidden;
                }

                return _id;
            }
        }

        /// <summary>
        /// Determines whether the element is hidden.
        /// </summary>
        public bool Hidden
        {
            get
            {
                if (_hidden == null)
                {
                    var (id, hidden) = XmlCompositeElement.GetNvPrValues();
                    _id = id;
                    _hidden = hidden;
                }

                return (bool)_hidden;
            }

        }

        /// <summary>
        /// Determines whether the element is placeholder.
        /// </summary>
        public bool IsPlaceholder
        {
            get
            {
                if (_isPlaceholder == null)
                {
                    _isPlaceholder = XmlCompositeElement.Descendants<P.PlaceholderShape>().Any();
                }

                return (bool)_isPlaceholder;
            }
        }

        /// <summary>
        /// Gets or sets element type.
        /// </summary>
        public ElementType Type { get; set; } //TODO: remove public modifier for setter

        /// <summary>
        /// Gets or sets the x-coordinate of the upper-left corner of the element in EMUs.
        /// </summary>
        public long X { get; set; }

        /// <summary>
        /// Gets or sets the y-coordinate of the upper-left corner of the element in EMUs.
        /// </summary>
        public long Y { get; set; }

        /// <summary>
        /// Gets or sets width of the element in EMUs.
        /// </summary>
        public long Width { get; set; }

        /// <summary>
        /// Gets or sets height of the element in EMUs.
        /// </summary>
        public long Height { get; set; }

        /// <summary>
        /// Gets or sets tag which can be used for any reason.
        /// </summary>
        [SuppressMessage("ReSharper", "UnusedMember.Global")]
        public object Tag { get; set; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Element"/> class.
        /// </summary>
        protected Element(ElementType et)
        {
            Type = et;
        }

        #endregion Constructors
    }
}