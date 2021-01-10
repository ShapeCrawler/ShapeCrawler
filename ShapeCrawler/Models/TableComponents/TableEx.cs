using System;
using System.Linq;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Settings;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
// ReSharper disable All

namespace SlideDotNet.Models.TableComponents
{
    /// <summary>
    /// Represents a table element on a slide.
    /// </summary>
    public class TableEx
    {
        #region Fields

        private readonly Lazy<RowCollection> _rowsCollection;
        private readonly P.GraphicFrame _sdkGrFrame;

        #endregion Fields

        #region Properties

        public RowCollection Rows => _rowsCollection.Value;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="TableEx"/> class.
        /// </summary>
        public TableEx(P.GraphicFrame xmlGrFrame)
        {
            _sdkGrFrame = xmlGrFrame ?? throw new ArgumentNullException(nameof(xmlGrFrame));
            _rowsCollection = new Lazy<RowCollection>(()=>GetRowsCollection());
        }

        #endregion Constructors

        #region Private Methods

        private RowCollection GetRowsCollection()
        {
            var sdkTblRows = _sdkGrFrame.Descendants<A.Table>().First().Elements<A.TableRow>();

            return new RowCollection(sdkTblRows);
        }

        #endregion Private Methods
    }
}