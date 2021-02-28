using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAdaptor
{
    public class WorksheetAdaptor : IDisposable
    {
        public Worksheet Worksheet { get; private set; }
        public SheetData SheetData { get; private set; }
        public CellRange UsedRange { get; private set; }
        public WorkbookAdaptor WorkbookAdaptor { get; private set; }
        public int RowCount => UsedRange.GetRowCount();
        public int ColumnCount => UsedRange.GetColumnCount();
        private Dictionary<CellReference, TableCell> _cellMap = new Dictionary<CellReference, TableCell>();
        private bool disposedValue;

        public WorksheetAdaptor(WorkbookAdaptor workbookAdaptor, Worksheet worksheet)
        {
            if (workbookAdaptor == null) throw new ArgumentNullException("WorkbookAdaptor cannot be null in WorksheetAdaptor");
            if (worksheet == null) throw new ArgumentNullException("Workseet cannot be null in WorksheetAdaptor");
            WorkbookAdaptor = workbookAdaptor;
            Worksheet = worksheet;
            SheetData = worksheet.GetFirstChild<SheetData>() as SheetData;
            SheetDimension dimension = worksheet.GetFirstChild<SheetDimension>() as SheetDimension;
            UsedRange = new CellRange(dimension.Reference.Value);
            SheetData sheetData = worksheet.GetFirstChild<SheetData>() as SheetData;
            foreach (Row row in sheetData)
            {
                foreach (Cell cell in row.ChildElements.OfType<Cell>())
                {
                    var c = new TableCell(cell, WorkbookAdaptor.SharedStrings);
                    _cellMap[c.Reference] = c;
                }
            }
        }
        public TableCell Cells(CellReference cr)
        {
            if (_cellMap.TryGetValue(cr, out TableCell cell))
                return cell;
            else
                return TableCell.Null;
        }
        public TableCell Cells(string cr) => Cells(new CellReference(cr));
        public TableCell Cells(int row, string column) => Cells(new CellReference(row, column));
        public TableCell Cells(int row, int column) => Cells(new CellReference(row, column));

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Worksheet = null;
                    SheetData = null;
                    UsedRange = null;
                    WorkbookAdaptor = null;
                }
                disposedValue = true;
            }
        }


        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}

