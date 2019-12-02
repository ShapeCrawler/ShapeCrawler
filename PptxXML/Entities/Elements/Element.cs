using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using objectEx.Extensions;
using PptxXML.Enums;
using PptxXML.Extensions;

namespace PptxXML.Entities.Elements
{
    /// <summary>
    /// Represent an element on a slide.
    /// </summary>
    public abstract class Element
    {
        #region Fields

        private readonly OpenXmlCompositeElement _xmlCompositeElement;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets or sets identifier.
        /// </summary>
        public int Id { get; private set; }

        public ElementType Type { get; set; }

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
        /// Determines whether the element is hidden.
        /// </summary>
        public bool Hidden { get; private set; }

        /// <summary>
        /// Gets or sets tag which can be used for any reason.
        /// </summary>
        [SuppressMessage("ReSharper", "UnusedMember.Global")]
        public object Tag { get; set; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initialise instance of <see cref="Element"/> class.
        /// </summary>
        /// <param name="xmlCompositeElement"></param>
        protected Element(OpenXmlCompositeElement xmlCompositeElement)
        {
            xmlCompositeElement.ThrowIfNull(nameof(xmlCompositeElement));
            _xmlCompositeElement = xmlCompositeElement;

            Init();
        }

        #endregion Constructors

        #region Private Methods

        private void Init()
        {
            // Initialise identifier and hidden value
            var (item1, item2) = _xmlCompositeElement.GetNvPrValues();
            Id = item1;
            Hidden = item2;
        }

        #endregion Private Methods
    }
}