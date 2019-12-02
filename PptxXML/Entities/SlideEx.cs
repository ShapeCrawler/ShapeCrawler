using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ObjectEx.Utilities;
using PptxXML.Entities.Elements;
using PptxXML.Services;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Entities
{
    /// <summary>
    /// Represents a slide.
    /// </summary>
    public class SlideEx
    {
        #region Fields

        private List<Element> _elements;

        #endregion Fields

        #region Dependencies

        private readonly SlidePart _xmlSldPart;
        private readonly PresentationDocument _xmlPreDoc;

        #endregion Dependencies

        #region Properties

        /// <summary>
        /// Gets elements.
        /// </summary>
        public IList<Element> Elements
        {
            get
            {
                if (_elements == null)
                {
                    InitElements();
                }

                return _elements;
            }
        }

        /// <summary>
        /// Returns slide number in presentation.
        /// </summary>
        public int Number { get; set; } //TODO: Remove public setter somehow

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initialize a new instance of the <see cref="SlideEx"/> class.
        /// </summary>
        public SlideEx(SlidePart xmlSldPart, PresentationDocument xmlPreDoc, int sldNumber)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            _xmlSldPart = xmlSldPart;
            Check.NotNull(xmlPreDoc, nameof(xmlPreDoc));
            _xmlPreDoc = xmlPreDoc;
            if (sldNumber < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(sldNumber));
            }
            Number = sldNumber;
        }

        #endregion Constructors

        #region Public Methods       

        #endregion Public Methods

        #region Private Methods

        private void InitElements()
        {
            _elements = new List<Element>();
            var elementCreator = new ElementCreator();
            foreach (var xmlCompositeElement in _xmlSldPart.Slide.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>())
            {
                if (xmlCompositeElement is P.GroupShape
                    || xmlCompositeElement is P.Picture
                    || xmlCompositeElement is P.Shape
                    || xmlCompositeElement is P.GraphicFrame)
                {
                    _elements.Add(elementCreator.Create(xmlCompositeElement));
                }
            }
        }

        #endregion Private Methods
    }
}