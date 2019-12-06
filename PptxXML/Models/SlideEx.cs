using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ObjectEx.Utilities;
using PptxXML.Enums;
using PptxXML.Exceptions;
using PptxXML.Models.Elements;
using PptxXML.Services;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Models
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
        private readonly IElementCreator _elCreator;

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
        public SlideEx(SlidePart xmlSldPart, int sldNumber, IElementCreator elCreator)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            Check.IsPositive(sldNumber, nameof(sldNumber));
            Check.NotNull(elCreator, nameof(elCreator));
            _xmlSldPart = xmlSldPart;
            Number = sldNumber;
            _elCreator = elCreator;
        }

        #endregion Constructors

        #region Private Methods

        private void InitElements()
        {
            _elements = new List<Element>();
            var parser = new ShapeTreeParser();
            var candidates = parser.CreateCandidates(_xmlSldPart.Slide.CommonSlideData.ShapeTree);

            foreach (var c in candidates)
            {
                switch (c.ElementType)
                {
                    case ElementType.Shape:
                        {
                            var el = _elCreator.CreateShape(c);
                            _elements.Add(el);
                            break;
                        }
                    case ElementType.Chart:
                        {
                            var el = _elCreator.CreateChart(c);
                            _elements.Add(el);
                            break;
                        }
                    case ElementType.Table:
                        {
                            var el = _elCreator.CreateTable(c);
                            _elements.Add(el);
                            break;
                        }
                    case ElementType.Picture:
                        {
                            var el = _elCreator.CreatePicture(c);
                            _elements.Add(el);
                            break;
                        }
                    default:
                        throw new PptxXMLException(nameof(ElementType));
                }
            }
        }

        #endregion Private Methods
    }
}