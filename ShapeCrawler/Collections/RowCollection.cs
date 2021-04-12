using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a table rows collection.
    /// </summary>
    public class RowCollection : EditableCollection<SCTableRow> //TODO extract interface and convert to internal
    {
        #region Constructors

        internal RowCollection(List<SCTableRow> rowList)
        {
            CollectionItems = rowList;
        }

        #endregion Constructors

        internal static RowCollection Create(SlideTable table, P.GraphicFrame pGraphicFrame)
        {
            IEnumerable<A.TableRow> aTableRows = pGraphicFrame.GetATable().Elements<A.TableRow>();
            var rowList = new List<SCTableRow>(aTableRows.Count());
            int rowIndex = 0;
            rowList.AddRange(aTableRows.Select(aTblRow => new SCTableRow(table, aTblRow, rowIndex++)));

            return new RowCollection(rowList);
        }

        #region Public Methods

        /// <summary>
        ///     Removes the specified table row.
        /// </summary>
        /// <param name="scTableRow"></param>
        public override void Remove(SCTableRow scTableRow)
        {
            scTableRow.ATableRow.Remove();
            CollectionItems.Remove(scTableRow);
        }

        /// <summary>
        ///     Removes table row by index.
        /// </summary>
        /// <param name="index"></param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= CollectionItems.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var innerRow = CollectionItems[index];
            Remove(innerRow);
        }

        #endregion Public Methods
    }
}