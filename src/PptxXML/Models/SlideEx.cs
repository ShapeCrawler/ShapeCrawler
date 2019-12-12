using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ObjectEx.Utilities;
using PptxXML.Enums;
using PptxXML.Exceptions;
using PptxXML.Models.Elements;
using PptxXML.Services;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Models
{
    /// <summary>
    /// Represents a slide.
    /// </summary>
    public class SlideEx
    {
        #region Fields

        private List<Element> _elements;

        #region Dependencies

        private readonly SlidePart _xmlSldPart;
        private readonly IElementFactory _elCreator;
        private readonly IGroupShapeTypeParser _shapeTreeParser;
        private readonly IGroupBuilder _groupBuilder;

        #endregion Dependencies

        #endregion Fields

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
        public SlideEx(SlidePart xmlSldPart, 
                       int sldNumber, 
                       IElementFactory elCreator, 
                       IGroupShapeTypeParser shapeTreeParser,
                       IGroupBuilder groupBuilder)
        {
            _xmlSldPart = xmlSldPart;
            Number = sldNumber;
            _elCreator = elCreator;
            _shapeTreeParser = shapeTreeParser;
            _groupBuilder = groupBuilder;
        }

        #endregion Constructors

        #region Private Methods

        private void InitElements()
        {
            _elements = new List<Element>();
            var shapeTree = _xmlSldPart.Slide.CommonSlideData.ShapeTree;
            var shapeTreeCandidates = _shapeTreeParser.CreateCandidates(shapeTree);

            foreach (var c in shapeTreeCandidates)
            {
                Element el;
                switch (c.ElementType)
                {
                    case ElementType.Shape:
                        {
                            el = _elCreator.CreateShape(c);
                            break;
                        }
                    case ElementType.Chart:
                        {
                            el = _elCreator.CreateChart(c);
                            break;
                        }
                    case ElementType.Table:
                        { 
                            el = _elCreator.CreateTable(c);
                            break;
                        }
                    case ElementType.Picture:
                        {
                            el = _elCreator.CreatePicture(c);
                            break;
                        }
                    case ElementType.Group:
                        {
                            el = _groupBuilder.Build((P.GroupShape)c.CompositeElement);
                            break;
                        }
                    default:
                        throw new PptxXMLException(nameof(ElementType));
                }
                _elements.Add(el);
            }
        }

        #endregion Private Methods
    }
}