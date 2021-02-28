using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAdaptor
{
    public class WorkbookAdaptor : IDisposable
    {
        private bool disposedValue;

        public SpreadsheetDocument Document { get; private set; } = null;
        public SharedStrings SharedStrings { get; private set; } = null;
        public Sheets Sheets { get; private set; } = null;
        protected WorkbookAdaptor()
        {

        }
        public WorksheetAdaptor GetWorksheet(object index)
        {
            if (index is string s)
                return GetWorksheet(s);
            else if (index is int i)
                return GetWorksheet(i);
            else
                return null;
        }
        public WorksheetAdaptor GetWorksheet(int index)
        {
            Worksheet sheet = GetWorksheet(Document.WorkbookPart, Sheets, index);
            return sheet != null ? new WorksheetAdaptor(this, sheet) : null;
        }
        public WorksheetAdaptor GetWorksheet(string name)
        {
            Worksheet sheet = GetWorksheet(Document.WorkbookPart, Sheets, name);
            return sheet != null ? new WorksheetAdaptor(this, sheet) : null;
        }
        static public WorkbookAdaptor Open(string workbookPath, bool readOnly = true)
        {
            if (workbookPath == null) throw new ArgumentNullException("Path to workbook cannot be null.");
            if (!File.Exists(workbookPath)) throw new ArgumentException("Workbook does not exist");
            try
            {
                SpreadsheetDocument doc = SpreadsheetDocument.Open(workbookPath, !readOnly);
                var wb = new WorkbookAdaptor();
                wb.Document = doc;
                SharedStringTable sst = doc.WorkbookPart.SharedStringTablePart.SharedStringTable;
                wb.SharedStrings = new SharedStrings(sst);
                wb.Sheets = doc.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                return wb;
            }
            catch(Exception ex)
            {
                throw new ApplicationException("Error opening workbook", ex);
            }
        }
        private Worksheet GetWorksheet(WorkbookPart workbookPart, Sheets sheets, int index)
        {
            if (index > 0 && index <= sheets.Count())
            {
                return GetWorksheet(workbookPart, sheets.ChildElements[index - 1] as Sheet);
            }
            return null;
        }
        private Worksheet GetWorksheet(WorkbookPart workbookPart, Sheet sheet)
        {
            if(sheet != null)
            {
                return ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
            }
            return null;
        }
        private Worksheet GetWorksheet(WorkbookPart workbookPart, Sheets sheets, string name)
        {
            foreach (var s in sheets.ChildElements.OfType<Sheet>())
            {
                if (String.Compare(s.Name, name, true) == 0)
                {
                    return GetWorksheet(workbookPart, s);
                }
            }
            return null;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (Document != null) Document.Dispose();
                    Document = null;
                    SharedStrings = null;
                    Sheets = null;
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
