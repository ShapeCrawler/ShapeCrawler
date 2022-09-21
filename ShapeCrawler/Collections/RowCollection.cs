using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Extensions;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a table rows collection.
    /// </summary>
    public class RowCollection : EditableCollection<SCTableRow> // TODO extract interface and convert to internal
    {
        #region Constructors

        private RowCollection(List<SCTableRow> rowList)
        {
            this.CollectionItems = rowList;
        }

        #endregion Constructors
        
        /// <inheritdoc/>
        public override void Remove(SCTableRow tableRow)
        {
            tableRow.ATableRow.Remove();
            this.CollectionItems.Remove(tableRow);
        }

        /// <summary>
        ///     Removes table row by index.
        /// </summary>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= this.CollectionItems.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var innerRow = this.CollectionItems[index];
            this.Remove(innerRow);
        }
        
        internal static RowCollection Create(SlideTable table, P.GraphicFrame pGraphicFrame)
        {
            IEnumerable<A.TableRow> aTableRows = pGraphicFrame.GetATable().Elements<A.TableRow>();
            var rowList = new List<SCTableRow>(aTableRows.Count());
            int rowIndex = 0;
            rowList.AddRange(aTableRows.Select(aTblRow => new SCTableRow(table, aTblRow, rowIndex++)));

            return new RowCollection(rowList);
        }
    }
}