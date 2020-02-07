using System.Collections.Generic;
using System.Linq;
using SlideXML.Models.Settings;
using SlideXML.Models.TableComponents;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
// ReSharper disable All

namespace SlideXML.Models.SlideComponents
{
    /// <summary>
    /// Represents a table element on a slide.
    /// </summary>
    public class Table
    {
        #region Fields

        private List<RowEx> _rows;
        private readonly P.GraphicFrame _xmlGrFrame;
        private readonly ElementSettings _elSettings;

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
        /// Initialise an instance of <see cref="Table"/> class.
        /// </summary>
        public Table(P.GraphicFrame xmlGrFrame, ElementSettings elSettings)
        {
            _xmlGrFrame = xmlGrFrame;
            _elSettings = elSettings;
        }

        #endregion Constructors

        #region Private Methods

        private void ParseRows()
        {
            var xmlRows = _xmlGrFrame.Descendants<A.Table>().Single().Elements<A.TableRow>();
            _rows = new List<RowEx>(xmlRows.Count());
            foreach (var r in xmlRows)
            {
                _rows.Add(new RowEx(r, _elSettings));
            }
        }

        #endregion Private Methods
    }
}