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
    /// Represents a parser of <see cref="P.ShapeTree"/> and <see cref="P.GroupShape"/> instances.
    /// </summary>
    /// <remarks>
    /// <see cref="P.ShapeTree"/> and <see cref="P.GroupShape"/> both derived from <see cref="P.GroupShapeType"/> class.
    /// </remarks>
    public class GroupShapeTypeParser : IGroupShapeTypeParser
    {
        #region Public Methods

        /// <summary>
        /// Creates candidate collection.
        /// </summary>
        /// <returns></returns>
        public IEnumerable<ElementCandidate> CreateCandidates(P.GroupShapeType groupTypeShape, bool groupParsed = true)
        {
            // Get all composite element
            var xmlElements = groupTypeShape.Elements<OpenXmlCompositeElement>();

            // Remove placeholders
            xmlElements = xmlElements.Where(e => !e.Descendants<P.PlaceholderShape>().Any());

            // FILTER PICTURES
            var pictureCandidates = xmlElements.Where(e => e is P.Picture
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
            var xmlShapes = xmlElements.Except(pictureCandidates).Where(e => e is P.Shape);
            var shapeCandidates = xmlShapes.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.Shape
            });

            // Table candidates
            var xmlTables = xmlElements.Except(pictureCandidates)
                                       .Except(xmlShapes)
                                       .Where(e => e is P.GraphicFrame grFrame && grFrame.Descendants<A.Table>().Any());
            var tableCandidates = xmlTables.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.Table
            });

            // Chart candidates
            var xmlCharts = xmlElements.Except(pictureCandidates)
                                       .Except(xmlShapes)
                                       .Except(xmlTables)
                                       .Where(e => e is P.GraphicFrame grFrame && grFrame.HasChart());
            var chartCandidates = xmlCharts.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.Chart
            });

            var allCandidates = picCandidates.Union(shapeCandidates).Union(tableCandidates).Union(chartCandidates);

            // Group candidates
            if (groupParsed)
            {
                var xmlGroupCandidates = xmlElements.Where(e => e is P.GroupShape).Select(ce => new ElementCandidate
                {
                    CompositeElement = ce,
                    ElementType = ElementType.Group
                });
                allCandidates = allCandidates.Union(xmlGroupCandidates);
            }

            return allCandidates;
        }

        #endregion Public Methods
    }
}
