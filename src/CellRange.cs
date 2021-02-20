using System;
using System.Collections.Generic;
using System.Text;

namespace exceltowiki
{
    public class CellRange
    {
        public CellReference ULCell { get; } = CellReference.Null;
        public CellReference LRCell { get; } = CellReference.Null;
        public CellRange()
        {

        }
        public CellRange(string range)
        {
            if (range == null) throw new ArgumentNullException();
            string[] crefs = range.Split(':');
            if (crefs.Length != 2) throw new ArgumentException("Invalid cell range");
            var cell0 = new CellReference(crefs[0]);
            var cell1 = new CellReference(crefs[1]);
            if(cell0 < cell1)
            {
                ULCell = cell0;
                LRCell = cell1;
            }
            else
            {
                ULCell = cell1;
                LRCell = cell0;
            }
        }
        public override bool Equals(object obj)
        {
            if(obj is CellRange crange)
            {
                return ULCell == crange.ULCell && LRCell == crange.LRCell;
            }
            return base.Equals(obj);
        }
        public override int GetHashCode()
        {
            return ULCell.GetHashCode() ^ LRCell.GetHashCode();
        }
        public override string ToString()
        {
            return ULCell.ToString() + ":" + LRCell.ToString();
        }
        static public bool operator==(CellRange crange0, CellRange crange1)
        {
            return crange0.Equals(crange1);
        }
        static public bool operator!=(CellRange crange0, CellRange crange1)
        {
            return !crange0.Equals(crange1);
        }
        public int GetColumnCount()
        {
            return LRCell.ColumnIndex - ULCell.ColumnIndex + 1;
        }
        public int GetRowCount()
        {
            return LRCell.Row - ULCell.Row + 1;
        }
    }
}
