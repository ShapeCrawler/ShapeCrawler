using System.Collections.Generic;
using System.Linq;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.TableComponents;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
// ReSharper disable All

namespace SlideXML.Models.SlideComponents
{
    /// <summary>
    /// Represents a table element on a slide.
    /// </summary>
    public class TableEx
    {
        #region Fields

        private List<RowEx> _rows;
        private readonly P.GraphicFrame _xmlGrFrame;
        private readonly IShapeContext _spContext;

        #endregion Fields

        #region Properties

        public IList<RowEx> Rows
        {
            get
            {
                if (_rows == null)
                {
                    ParseRows();
                }

                return _rows;
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="TableEx"/> class.
        /// </summary>
        public TableEx(P.GraphicFrame xmlGrFrame, IShapeContext spContext)
        {
            _xmlGrFrame = xmlGrFrame;
            _spContext = spContext;
        }

        #endregion Constructors

        #region Private Methods

        private void ParseRows()
        {
            var xmlRows = _xmlGrFrame.Descendants<A.Table>().Single().Elements<A.TableRow>();
            _rows = new List<RowEx>(xmlRows.Count());
            foreach (var r in xmlRows)
            {
                _rows.Add(new RowEx(r, _spContext));
            }
        }

        #endregion Private Methods
    }
}