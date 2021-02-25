using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelAdaptor;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

[assembly: InternalsVisibleTo("exceltowiki.Test")]
namespace exceltowiki
{
    public class Program
    {
        internal struct WorksheetDef
        {
            public int Index;
            public string Name;
            public object Reference => Name == null ? (object)Index : Name;
            public string Description => Name == null ? Index.ToString() : Name;
            public bool HasName => Name != null;
        };
        internal enum FormatType
        {
            None,
            Date
        };
        internal struct ColumnDef
        {
            public string Name;
            public FormatType Format;
            public string FormatParameter;
        }


        internal static void Main(string[] args)
        {
            string inputFile = null;
            WorksheetDef worksheet = new WorksheetDef { Index = 1 };
            bool headers = false;
            string[] columnNames = null;
            FormatType[] columnFormats = null;
            string dateFormat = null;
            try
            {
                for (int i = 0; i < args.Length; ++i)
                {
                    string arg = args[i];
                    switch (arg)
                    {
                        case "--columns":
                            columnNames = GetColumnNames(GetArg(args, ++i));
                            break;
                        case "--formats":
                            columnFormats = GetColumnFormats(GetArg(args, ++i));
                            break;
                        case "--headers":
                            headers = true;
                            break;
                        case "--date-format":
                            dateFormat = GetDateFormat(GetArg(args, ++i));
                            break;
                        case "--worksheet":
                            worksheet = GetWorksheet(GetArg(args, ++i));
                            break;
                        case "--help":
                            Usage();
                            return;
                        case "--version":
                            Version();
                            return;
                        default:
                            if (String.IsNullOrEmpty(arg)) throw new ApplicationException("Null argument seen on command line.");
                            if (arg[0] == '-') throw new ApplicationException($"Unrecognized switch argument {arg} seen.");
                            if (inputFile == null) inputFile = arg;
                            else throw new ApplicationException($"Extraneous argument {arg} seen.");
                            break;
                    }
                }
                if (inputFile == null) throw new ApplicationException("Missing input file argument.");
                if (!File.Exists(inputFile)) throw new ApplicationException($"Input file '{inputFile}' does not exist.");
                ConvertExcelToWiki(inputFile, worksheet, columnNames, columnFormats, dateFormat, headers);
            }
            catch (Exception ex)
            {
                Usage(ex.Message);
            }
        }

        internal static string GetDateFormat(string arg)
        {
            string defaultDateFormat = "g";
            if (!String.IsNullOrEmpty(arg))
            {
                ValidateDateFormat(arg);
                return arg;
            }
            return defaultDateFormat;
        }

        internal static string GetArg(string[] args, int iarg)
        {
            if (iarg >= 0 && iarg < args.Length)
                return args[iarg];
            throw new ApplicationException("Missing command line argument.");
        }

        internal static WorksheetDef GetWorksheet(string arg)
        {
            if (int.TryParse(arg, out int v))
                return new WorksheetDef { Index = v };
            else
                return new WorksheetDef { Name = arg };
        }

        internal static bool ValidateDateFormat(string arg)
        {
            try
            {
                string s = DateTime.Now.ToString(arg);
                return true;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Invalid date format argument", ex);
            }
        }

        internal static void ConvertExcelToWiki(string inputFile, WorksheetDef worksheetDef, string[] columnNames, FormatType[] columnFormats, string dateFormat, bool headers)
        {
            try
            {
                TextWriter writer = Console.Out;
                string path = inputFile;
                if (!Path.IsPathRooted(path)) path = Path.Combine(Directory.GetCurrentDirectory(), inputFile);
                using (WorkbookAdaptor workbook = WorkbookAdaptor.Open(path, true))
                {
                    WorksheetAdaptor worksheet = workbook.GetWorksheet(worksheetDef.Reference);
                    if (worksheet != null)
                    {

                        WriteError("Writing table...");
                        writer.WriteLine("{| class=\"wikitable\"");
                        int nrows = worksheet.RowCount;
                        int ncolumns = worksheet.ColumnCount;
                        if (nrows == 0) throw new ApplicationException("The source table is empty.");
                        ColumnDef[] columns = GetColumns(ncolumns, columnNames, columnFormats, dateFormat);
                        for (int rowIndex = 1; rowIndex <= nrows; ++rowIndex)
                        {
                            if (headers && rowIndex == 1)
                            {
                                writer.WriteLine("|+");
                                foreach (var column in columnNames)
                                {
                                    string s = worksheet.Cells(rowIndex, column).ToString();
                                    writer.WriteLine("|" + s);
                                }
                            }
                            else
                            {
                                WriteError($"\rWriting row {rowIndex + 1} of {nrows}.", false);
                                writer.WriteLine("|-");
                                for (int i = 0; i < columns.Length; ++i)
                                {
                                    string column = columns[i].Name;
                                    string s = worksheet.Cells(rowIndex, column).ToString();
                                    string s1 = FormatColumn(s, columns[i]);
                                    writer.WriteLine("|" + s1);
                                }
                            }
                        }
                    }
                    else
                    {
                        throw new ApplicationException("Worksheet not found");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Unable to convert excel to wiki", ex);
            }
        }

        internal static string[] GetExcelColumns(int ncolumns)
        {
            if (ncolumns > 26 * 26) throw new ApplicationException("Source table has too many columns.");
            List<string> columns = new List<string>();
            for (int i = 0; i < ncolumns; i++)
            {
                char a = (char)((int)'A' + i % 26);
                if (i < 26)
                {
                    columns.Add(new string(a, 1));
                }
                else
                {
                    char b = (char)('A' - 1 + i / 26);
                    columns.Add(new string(new char[] { b, a }));
                }
            }
            return columns.ToArray();
        }

        internal static string FormatColumn(string s, ColumnDef cd)
        {
            switch (cd.Format)
            {
                case FormatType.Date: return FormatDate(s, cd.FormatParameter);
                default:
                    return s;
            }
        }

        internal static string FormatDate(string s, string dateFormat)
        {
            if (DateTime.TryParse(s, out DateTime v))
            {
                return v.ToString(dateFormat);
            }
            else
            {
                return s;
            }
        }
        internal static string[] GetColumnNames(string arg)
        {
            string[] columns = arg.Split(',');
            foreach (var column in columns)
            {
                if (!IsValidExcelColumn(column)) throw new ApplicationException("Invalid Excel column seen; must be A-Z and AA-ZZ");
            }
            return columns;
        }

        internal static ColumnDef[] GetColumns(int columns, string[] columnNames, FormatType[] columnFormats, string dateFormat)
        {
            if (columnNames == null) columnNames = GetExcelColumns(columns);
            if (columnFormats == null) columnFormats = new FormatType[0];
            if (columns < columnNames.Length) throw new ApplicationException("The source table does not have enough columns");
            List<ColumnDef> cds = new List<ColumnDef>();
            for (int i = 0; i < columnNames.Length; ++i)
            {
                string name = columnNames[i];
                FormatType type = i < columnFormats.Length ? columnFormats[i] : FormatType.None;
                ColumnDef cd = new ColumnDef { Name = name, Format = type };
                if (type == FormatType.Date) cd.FormatParameter = dateFormat;
                cds.Add(cd);
            }
            return cds.ToArray();
        }

        internal static bool IsValidExcelColumn(string column)
        {
            string column1 = column.ToUpper();
            string validColumns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            if (column1.Length < 1 || column1.Length > 2) return false;
            if (column1.Length == 1 && validColumns.Contains(column1.ToUpper())) return true;
            else if (validColumns.Contains(column1[0]) && validColumns.Contains(column[1])) return true;
            else return false;
        }

        internal static FormatType[] GetColumnFormats(string arg)
        {
            string[] formatNames = arg.ToUpper().Split(',');
            List<FormatType> formats = new List<FormatType>();
            for (int i = 0; i < formatNames.Length; ++i)
            {
                string formatName = formatNames[i];
                switch (formatName)
                {
                    case "":
                        formats.Add(FormatType.None);
                        break;
                    case "DATE":
                        formats.Add(FormatType.Date);
                        break;
                    default:
                        throw new ApplicationException("Invalid formats argument");
                }
            }
            return formats.ToArray();
        }
        private static void WriteError(string[] lines)
        {
            foreach (var line in lines)
            {
                WriteError(line);
            }
        }
        private static void WriteError()
        {
            Console.Error.WriteLine();
        }
        private static void WriteError(string line, bool addLF = true)
        {
            if (addLF)
                Console.Error.WriteLine(line);
            else
                Console.Error.Write(line);
        }
        internal static void Usage(string error = "")
        {
            WriteError();
            if (!String.IsNullOrEmpty(error))
            {
                WriteError("Error: " + error);
                WriteError();
            }
            string[] usage =
            {
                "Usage:",
                "",
                "exceltowiki input [--columns clist] [--format flist] [--sheet sname] [--dateFormat dformat] [--headers]",
                "",
                "\tinput         Path to input Excel spreadsheet file (.xls)",
                "\t--columns     Comma separated list of column names, e.g. A,C,D,AA,B",
                "\t--format      Comma separated list of format conversions, e.g. date,,date",
                "\t--headers     Columns have headers",
                "\t--date-format Format of date output on date conversion of column data",
                "\t--help        display usage",
                "\t--version     display version",
                "",
                "Where:",
                "",
                "\tclist         Excel column list, 1 or 2 alphabetic letters in order that the columns will appear in output.",
                "\tflist         Format list, one of date or empty string. Position in list corresponds to position in clist.",
                "\t              date format indicates column is date/time.  Converted in output according to date-format argument",
                "\tdformat       Standard or custom datetime string as described in .NET Documentation. e.g. \"dd-MMM-yy\".  Default is \"g\""
            };
            WriteError(usage);
        }
        internal static void Version()
        {
            Assembly a = Assembly.GetExecutingAssembly();
            var name = a.GetName();
            WriteError(name.Name + " " + name.Version);
        }
    }
}
