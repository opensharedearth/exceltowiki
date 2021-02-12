using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace exceltowiki
{
    class Program
    {
        public struct WorksheetDef
        {
            public int Index;
            public string Name;
            public object Reference => Name == null ? (object)Index : Name;
            public string Description => Name == null ? Index.ToString() : Name;
        };
        public enum FormatType
        {
            None,
            Date
        };
        public struct ColumnDef
        {
            public string Name;
            public FormatType Format;
            public string FormatParameter;
        }

        static string s_inputFile = null;
        static string s_outputFile = null;
        static ColumnDef[] s_columns = null;
        static WorksheetDef s_worksheet = new WorksheetDef { Index = 1 };
        static bool s_overwrite = false;
        static bool s_headers = false;
        static string s_defaultDateFormat = "g";

        static void Main(string[] args)
        {
            try
            {
                for (int i = 0; i < args.Length; ++i)
                {
                    string arg = args[i];
                    switch (arg)
                    {
                        case "--columns":
                            s_columns = GetColumns(GetArg(args, ++i));
                            break;
                        case "--formats":
                            GetFormats(GetArg(args, ++i), s_columns);
                            break;
                        case "-Y":
                            s_overwrite = true;
                            break;
                        case "--headers":
                            s_headers = true;
                            break;
                        case "--date-format":
                            GetDateFormat(GetArg(args, ++i), s_columns);
                            break;
                        case "--worksheet":
                            s_worksheet = GetWorksheet(GetArg(args, ++i));
                            break;
                        default:
                            if (String.IsNullOrEmpty(arg)) throw new ApplicationException("Null argument seen on command line.");
                            if (arg[0] == '-') throw new ApplicationException($"Unrecognized switch argument {arg} seen.");
                            if (s_inputFile == null) s_inputFile = arg;
                            else if (s_outputFile == null) s_outputFile = arg;
                            else throw new ApplicationException($"Extraneous argument {arg} seen.");
                            break;
                    }
                }
                if (s_inputFile == null) throw new ApplicationException("Missing input file argument.");
                if (s_outputFile == null) throw new ApplicationException("Missing output file argument.");
                if (!File.Exists(s_inputFile)) throw new ApplicationException($"Input file '{s_inputFile}' does not exist.");
                if (File.Exists(s_outputFile) && !s_overwrite) throw new ApplicationException($"Output file '{s_outputFile} exists and overwrite not specified.");
                ConvertExcelToWiki(s_inputFile, s_outputFile, s_worksheet, s_columns, s_headers);
            }
            catch(Exception ex)
            {
                Usage(ex.Message);
            }
        }

        private static string GetDateFormat(string arg, ColumnDef[] cds)
        {
            ValidateDateFormat(arg);
            for(int i = 0; i < cds.Length; ++i)
            {
                if(cds[i].Format == FormatType.Date)
                {
                    cds[i].FormatParameter = arg;
                }
            }
            return arg;
        }

        private static string GetArg(string[] args, int iarg)
        {
            if (iarg >= 0 && iarg < args.Length)
                return args[iarg];
            throw new ApplicationException("Missing command line argument.");
        }

        private static WorksheetDef GetWorksheet(string arg)
        {
            if (int.TryParse(arg, out int v))
                return new WorksheetDef { Index = v };
            else
                return new WorksheetDef { Name = arg };
        }

        private static bool ValidateDateFormat(string arg)
        {
            try
            {
                string s = DateTime.Now.ToString(arg);
                return true;
            }
            catch(Exception ex)
            {
                throw new ApplicationException("Invalid date format argument", ex);
            }
        }

        private static void ConvertExcelToWiki(string s_inputFile, string s_outputFile, WorksheetDef worksheetDef, ColumnDef[] columns, bool headers)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                if (File.Exists(s_outputFile)) File.Delete(s_outputFile);
                using (FileStream outputStream = File.Open(s_outputFile, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    using(StreamWriter writer = new StreamWriter(outputStream))
                    {
                        string path = s_inputFile;
                        if (!Path.IsPathRooted(path)) path = Path.Combine(Directory.GetCurrentDirectory(), s_inputFile);
                        Workbook workbook = app.Workbooks.Open(path);
                        if (workbook != null)
                        {
                            Worksheet worksheet = workbook.Worksheets[worksheetDef.Reference];
                            if (worksheet != null)
                            {
                                writer.WriteLine("{| class=\"wikitable\"");
                                int nrows = worksheet.UsedRange.Rows.Count;
                                if (nrows == 0) throw new ApplicationException("The source table is empty.");
                                int ncolumns = worksheet.UsedRange.Columns.Count;
                                if (columns == null) columns = GetExcelColumns(ncolumns);
                                if (ncolumns < columns.Length) throw new ApplicationException("The source table does not have enough columns");
                                int irow = 1;
                                Console.Write("Writing table...");
                                if(s_headers)
                                {
                                    writer.WriteLine("|+");
                                    foreach(var column in columns)
                                    {
                                        string s = worksheet.Cells[irow, column.Name].Value;
                                        writer.WriteLine("|" + s);
                                    }
                                    irow++;
                                }
                                for (int r = irow; r <= nrows; ++r)
                                {
                                    Console.Write($"\rWriting row {r + 1} of {nrows}.");
                                    writer.WriteLine("|-");
                                    for (int i = 0; i < columns.Length; ++i)
                                    {
                                        string column = columns[i].Name;
                                        string s = worksheet.Cells[r, column].Value;
                                        string s1 = FormatColumn(s, columns[i]);
                                        writer.WriteLine("|" + s1);
                                    }
                                }
                                Console.WriteLine($"\rFinished writing {nrows} rows.");
                                writer.WriteLine("|}");
                            }
                            else
                            {
                                throw new ApplicationException($"Worksheet {worksheetDef.Description} not found.");
                            }
                        }
                        writer.Close();
                    }
                }
            }
            finally
            {
                app.Quit();
            }
        }

        private static ColumnDef[] GetExcelColumns(int ncolumns)
        {
            if (ncolumns > 26 * 26) throw new ApplicationException("Source table has too many columns.");
            List<ColumnDef> columns = new List<ColumnDef>();
            for(int i = 0; i < ncolumns; i++)
            {
                char a = (char)((int)'A' + i % 26);
                if(i < 26)
                {
                    columns.Add(new ColumnDef { Name = new string(a, 1) , Format = FormatType.None });   
                }
                else
                {
                    char b = (char)((int)'A' + i / 26);
                    columns.Add(new ColumnDef { Name = new string(new char[] { b, a }), Format = FormatType.None });
                }
            }
            return columns.ToArray();
        }

        private static string FormatColumn(string s, ColumnDef cd)
        {
            switch (cd.Format)
            {
                case FormatType.Date: return FormatDate(s, cd.FormatParameter);
                default:
                    return s;
            }
        }

        private static string FormatDate(string s, string dateFormat)
        {
            if(DateTime.TryParse(s, out DateTime v))
            {
                return v.ToString(dateFormat);
            }
            else
            {
                return s;
            }
        }

        private static ColumnDef[] GetColumns(string arg)
        {
            string[] columns = arg.Split(',');
            List<ColumnDef> cds = new List<ColumnDef>();
            foreach(var column in columns)
            {
                if (!IsValidExcelColumn(column)) throw new ApplicationException("Invalid Excel column seen; must be A-Z and AA-ZZ");
                cds.Add(new ColumnDef { Name = column });
            }
            return cds.ToArray();
        }

        private static bool IsValidExcelColumn(string column)
        {
            string column1 = column.ToUpper();
            string validColumns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            if (column1.Length < 1 || column1.Length > 2) return false;
            if (column1.Length == 1 && validColumns.Contains(column1.ToUpper())) return true;
            else if (validColumns.Contains(column1[0]) && validColumns.Contains(column[1])) return true;
            else return false;
        }

        private static void GetFormats(string arg, ColumnDef[] cds)
        {
            string[] formats = arg.ToUpper().Split(',');
            for(int i = 0; i < formats.Length; ++i)
            {
                if (i >= cds.Length) throw new ApplicationException("More column formats than columns defined.");
                string format = formats[i];
                switch (format)
                {
                    case "":
                        break;
                    case "DATE":
                        cds[i].Format = FormatType.Date;
                        cds[i].FormatParameter = s_defaultDateFormat;
                        break;
                    default:
                        throw new ApplicationException("Invalid formats argument");
                }
            }
        }

        static void Usage(string error = "")
        {
            if(!String.IsNullOrEmpty(error))
            {
                Console.WriteLine("Error: " + error);
                Console.WriteLine();
            }
            Console.WriteLine("Usage:");
            Console.WriteLine();
            Console.WriteLine("exceltowiki [-Y] input [--columns clist] [--format flist] [--sheet sname] [--dateFormat dformat] [--headers] output");
            Console.WriteLine();
            Console.WriteLine("\tinput       Path to input Excel spreadsheet file (.xls)");
            Console.WriteLine("\toutput      Path to output wiki contents file (.wiki)");
            Console.WriteLine("\t-Y          Overwrite destination file if it exists");
            Console.WriteLine("\t--columns    Comma separated list of column names, e.g. A,C,D,AA,B");
            Console.WriteLine("\t--format     Comma separated list of format conversions, e.g. date,,date");
            Console.WriteLine("\t--headers    Columns have headers");
            Console.WriteLine("\t--date-format Format of date output on date conversion of column data");
            Console.WriteLine();
            Console.WriteLine("Where:");
            Console.WriteLine();
            Console.WriteLine("\tclist       Excel column list, 1 or 2 alphabetic letters in order that the columns will appear in output.");
            Console.WriteLine("\tflist       Format list, one of date or empty string. Position in list corresponds to position in clist.");
            Console.WriteLine("\t            date format indicates column is date/time.  Converted in output according to date-format argument");
            Console.WriteLine("\tdformat     Standard or custom datetime string as described in .NET Documentation. e.g. dd-MMM-yy");

        }
    }
}
