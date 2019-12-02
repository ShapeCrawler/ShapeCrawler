using System.Diagnostics.CodeAnalysis;
using PptxXML.Enums;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Entities.Elements
{
    /// <summary>
    /// Represents a shape element on slide.
    /// </summary>
    public class Shape : Element
    {
        #region Fields

        private readonly P.Shape _xmlShape;

        #endregion

        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="Shape"/> class.
        /// </summary>
        /// <param name="xmlShape"></param>
        [SuppressMessage("ReSharper", "SuggestBaseTypeForParameter")]
        public Shape(P.Shape xmlShape) :
            base(xmlShape)
        {
            _xmlShape = xmlShape;

            Init();
        }

        #endregion Constructors

        #region Private Methods

        private void Init()
        {
            Type = ElementType.Shape;
        }

        #endregion
    }
}