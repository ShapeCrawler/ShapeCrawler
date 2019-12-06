using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using PptxXML.Enums;
using PptxXML.Extensions;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Services
{
    /// <summary>
    /// Provides APIs to parse <see cref="P.ShapeTree"/> instance.
    /// </summary>
    public interface IShapeTreeParser
    {
        IEnumerable<ElementCandidate> CreateCandidates(P.ShapeTree shapeTree);
    }


    /// <summary>
    /// Represents a parser of <see cref="P.ShapeTree"/> instance.
    /// </summary>
    public class ShapeTreeParser
    {
        /// <summary>
        /// Creates candidate collection.
        /// </summary>
        /// <returns></returns>
        public IEnumerable<ElementCandidate> CreateCandidates(P.ShapeTree shapeTree)
        {
            var allXmlElements = shapeTree.Elements<OpenXmlCompositeElement>();

            var filterXmlElements = allXmlElements.Where(e => !e.Descendants<P.PlaceholderShape>().Any()); // remove placeholders
            filterXmlElements = filterXmlElements.Where(e => !(e is P.GroupShape)); // remove groups

            // FILTER PICTURES
            var pictureCandidates = filterXmlElements.Where(e => e is P.Picture 
                                                                 || e is P.Shape && e.Descendants<A.BlipFill>().Any()
                                                                 || e is P.GraphicFrame && e.Descendants<P.Picture>().Any());
            var graphicFrameImages = pictureCandidates.Where(e => e is P.GraphicFrame).SelectMany(e => e.Descendants<P.Picture>());
            var picAndShapeImages = pictureCandidates.Where(e => e is P.Picture
                                                                 || e is P.Shape && e.Descendants<A.BlipFill>().Any());

            // Picture candidates
            var xmlPictures = graphicFrameImages.Union(picAndShapeImages);
            var picCandidates = xmlPictures.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.Picture
            });

            // Shape candidates
            var xmlShapes = filterXmlElements.Except(pictureCandidates).Where(e => e is P.Shape);
            var shapeCandidates = xmlShapes.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.Shape
            });

            // Table candidates
            var xmlTables = filterXmlElements.Except(pictureCandidates)
                                             .Except(xmlShapes)
                                             .Where(e => e is P.GraphicFrame grFrame && grFrame.Descendants<A.Table>().Any());
            var tableCandidates = xmlTables.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.Table
            });

            // Chart candidates
            var xmlCharts = filterXmlElements.Except(pictureCandidates)
                                             .Except(xmlShapes)
                                             .Except(xmlTables)
                                             .Where(e => e is P.GraphicFrame grFrame && grFrame.HasChart());
            var chartCandidates = xmlCharts.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.Chart
            });

            return picCandidates.Union(shapeCandidates).Union(tableCandidates).Union(chartCandidates);
        }
    }

    /// <summary>
    /// Represents a parsed candidate element.
    /// </summary>
    public class ElementCandidate
    {
        /// <summary>
        /// Gets or sets corresponding element type.
        /// </summary>
        public ElementType ElementType { get; set; }

        /// <summary>
        /// Gets or sets instance of <see cref="OpenXmlCompositeElement"/>.
        /// </summary>
        public OpenXmlCompositeElement CompositeElement;
    }
}
