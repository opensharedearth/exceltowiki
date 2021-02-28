using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAdaptor
{
    public class TableCell
    {
        public static TableCell Null = new TableCell();
        public CellReference Reference { get; private set; } = null;
        public Cell Cell { get; private set; } = null;
        public CellValues DataType { get; private set; } = CellValues.Error;
        public CellValue CellValue { get; private set; } = null;
        SharedStrings _sst = null;
        public TableCell()
        {

        }
        public TableCell(Cell cell, SharedStrings sst = null)
        {
            Cell = cell;
            Reference = new CellReference(cell.CellReference.Value);
            if((object)cell.DataType == null)
            {
                DataType = CellValues.Number;
            }
            else
            {
                DataType = cell.DataType;
            }
            CellValue = cell.CellValue;
            _sst = sst;
        }
        public object Value
        {
            get
            {
                if(CellValue != null)
                {
                    switch (DataType)
                    {
                        case CellValues.SharedString:
                            return _sst != null ? _sst[CellValue.Text] : (object)"";
                        default:
                            return CellValue.Text;
                    }
                }
                else
                {
                    return "";
                }
            }
        }
        public int RowIndex => Reference.Row;
        public int ColumnIndex => Reference.ColumnIndex;
        public string Column => Reference.Column;
        public override string ToString()
        {
            return Value.ToString();
        }
        public bool IsNull => Reference == CellReference.Null;
    }
}
