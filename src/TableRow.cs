using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace exceltowiki
{
    public class TableRow
    {
        Dictionary<int,TableCell> _cells = new Dictionary<int,TableCell>();
        public TableRow()
        {

        }
        public TableRow(Row row, SharedStrings ss = null)
        {
            foreach (Cell cell in row.ChildElements.OfType<Cell>())
            {
                var c = new TableCell(cell, ss);
                _cells[c.ColumnIndex] = c;
            }
        }
        public TableCell this[int i] => _cells.TryGetValue(i, out TableCell c) ? c : TableCell.Null;
        public TableCell this[string i] => this[CellReference.GetColumnIndex(i)];
    }
}
