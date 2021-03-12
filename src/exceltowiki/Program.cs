using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelAdaptor;
using WikiAdaptor;
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
            public CellRange Range;
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
            string wikiurl = null;
            string username = null;
            string password = null;
            string pagePrefix = null;
            string tableTitle = null;
            string titleColumn = null;
            string[] pageColumnNames = null;
            FormatType[] pageFormats = null;
            bool overwrite = false;
            bool noTable = false;
            bool noPages = false;
            bool testLogin = false;
            string editSummary = "Modified by exceltowiki";
            CellRange range = null;
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
                        case "--worksheet-range":
                            range = new CellRange(GetArg(args, ++i));
                            break;
                        case "--wiki":
                            wikiurl = GetUrl((GetArg(args, ++i)));
                            break;
                        case "--username":
                            username = GetUsername(GetArg(args, ++i));
                            break;
                        case "--password":
                            password = GetPassword(GetArg(args, ++i));
                            break;
                        case "--table-title":
                            tableTitle = GetTitle(GetArg(args, ++i));
                            break;
                        case "--title-column":
                            titleColumn = GetTitleColumn(GetArg(args, ++i));
                            break;
                        case "--page-prefix":
                            pagePrefix = GetPagePrefix(GetArg(args, ++i));
                            break;
                        case "--page-columns":
                            pageColumnNames = GetColumnNames(GetArg(args, ++i));
                            break;
                        case "--page-format":
                            pageFormats = GetColumnFormats(GetArg(args, ++i));
                            break;
                        case "--overwrite":
                            overwrite = true;
                            break;
                        case "--no-table":
                            noTable = true;
                            break;
                        case "--no-pages":
                            noPages = true;
                            break;
                        case "--test-login":
                            testLogin = true;
                            break;
                        case "--edit-summary":
                            editSummary = GetEditSummary(GetArg(args, ++i));
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
                worksheet.Range = range;
                bool wikiOut = wikiurl != null;
                WikiAdaptor.WikiAdaptor wiki = null;
                if (testLogin)
                {
                    if (!wikiOut) throw new ArgumentException("Test login needs a wiki url specified as well as username and password");
                    if (username == null | password == null) throw new ArgumentException("Test login needs both username and password arguments");
                    wiki = new WikiAdaptor.WikiAdaptor(wikiurl);
                    wiki.Login(username, password);
                    WriteError("Login successful!");
                    return;
                }
                if (inputFile == null) throw new ApplicationException("Missing input file argument.");
                if (!File.Exists(inputFile)) throw new ApplicationException($"Input file '{inputFile}' does not exist.");
                if (wikiOut && (username == null | password == null)) throw new ArgumentException("Both username and password must be specified for a wiki desination");
                if (wikiOut && tableTitle == null) throw new ArgumentException("The table title must be specified for a wiki destination");
                if ((pageColumnNames != null || pageFormats != null) && titleColumn == null) throw new ArgumentException("Column title must be specified if wiki pages are to be created.");
                if (wikiOut)
                {
                    wiki = new WikiAdaptor.WikiAdaptor(wikiurl);
                    wiki.Login(username, password);
                }
                ConvertExcelToWiki(inputFile, wiki, worksheet, columnNames, columnFormats, dateFormat, headers,
                    noTable, noPages, overwrite, tableTitle, titleColumn, pagePrefix, pageColumnNames, pageFormats, editSummary);
                if (wikiOut)
                {
                    wiki.Dispose();
                }
            }
            catch (Exception ex)
            {
                WriteError(ex);
                Usage();
            }
        }

        private static string GetEditSummary(string v)
        {
            return v;
        }

        private static string GetPagePrefix(string v)
        {
            return v;
        }

        private static string GetTitleColumn(string v)
        {
            if (CellReference.IsValidColumn(v)) return v;
            throw new ArgumentException("Invalid argument to page-title.  Must be an Excel column.");
        }

        private static string GetTitle(string v)
        {
            if (!WikiSupport.IsValidTitle(v)) throw new ArgumentException("Invalid argument to table-title.  Must be a legal wiki title.");
            return v;
        }

        private static string GetPassword(string v)
        {
            return v;
        }

        private static string GetUsername(string v)
        {
            return v;
        }

        private static string GetUrl(string v)
        {
            if (Uri.TryCreate(v, UriKind.Absolute, out Uri uri))
            {
                if (v.Last() != '/') v += '/';
                return v;
            }
            throw new ArgumentException("wikiurl argument is invalid.");
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

        internal static void ConvertExcelToWiki(string inputFile, WikiAdaptor.WikiAdaptor wiki, WorksheetDef worksheetDef, string[] columnNames, FormatType[] columnFormats, string dateFormat, bool headers,
             bool noTable, bool noPages, bool overwrite, string tableTitle, string titleColumn, string pagePrefix, string[] pageColumnNames, FormatType[] pageColumnFormats,
             string editSummary)
        {
            try
            {
                bool wikiOut = wiki != null;
                TextWriter writer = Console.Out;
                MemoryStream outputStream = null;
                if (wikiOut)
                {
                    outputStream = new MemoryStream();
                    writer = new StreamWriter(outputStream);
                }
                string path = inputFile;
                if (!Path.IsPathRooted(path)) path = Path.Combine(Directory.GetCurrentDirectory(), inputFile);
                using (WorkbookAdaptor workbook = WorkbookAdaptor.Open(path, true))
                {
                    WorksheetAdaptor worksheet = workbook.GetWorksheet(worksheetDef.Reference);
                    if (worksheet != null)
                    {
                        CellRange range = worksheetDef.Range;
                        int nrows = worksheet.RowCount;
                        int ncolumns = worksheet.ColumnCount;
                        int row1 = 1;
                        int column1 = 1;
                        if (range != null)
                        {
                            nrows = range.GetRowCount();
                            ncolumns = range.GetColumnCount();
                            row1 = range.ULCell.Row;
                            column1 = range.ULCell.ColumnIndex;
                        }
                        if (nrows == 0) throw new ApplicationException("The source table is empty.");
                        ColumnDef[] columns = GetColumns(column1, ncolumns, columnNames, columnFormats, dateFormat);
                        if (!noTable)
                        {
                            WriteError("Writing table...");
                            writer.WriteLine("{| class=\"wikitable\"");
                            for (int rowIndex = row1; rowIndex <= nrows; ++rowIndex)
                            {
                                if (headers && rowIndex == row1)
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
                                    WriteError($"\rWriting table row {rowIndex} of {nrows}.", false);
                                    writer.WriteLine("|-");
                                    for (int i = 0; i < columns.Length; ++i)
                                    {
                                        string column = columns[i].Name;
                                        string s = worksheet.Cells(rowIndex, column).ToString();
                                        string s1 = FormatColumn(s, columns[i]);
                                        if (column == titleColumn)
                                        {
                                            writer.WriteLine("|" + GetPageLink(pagePrefix, s1));
                                        }
                                        else
                                        {
                                            writer.WriteLine("|" + s1);
                                        }
                                    }
                                }
                            }
                            WriteError(" Finished.");
                        }
                        else
                        {
                            WriteError("Skipping table creation");
                        }
                        if (wikiOut)
                        {
                            writer.Flush();
                            outputStream.Position = 0L;
                            ColumnDef[] pageColumns = GetColumns(column1, ncolumns, pageColumnNames, pageColumnFormats, dateFormat);
                            if (!noTable)
                            {
                                try
                                {
                                    wiki.CreatePage(tableTitle, editSummary, outputStream, overwrite).Wait();
                                    WriteError($"Created wiki table {tableTitle}");
                                }
                                catch (Exception ex)
                                {
                                    WriteError("*Warning* Unable to create wiki table: " + ex.Message);
                                }
                            }
                            if (titleColumn != null && !noPages)
                            {
                                Dictionary<string, string> sectionTitles = new Dictionary<string, string>();
                                for (int rowIndex = row1; rowIndex <= nrows; ++rowIndex)
                                {
                                    if (headers && rowIndex == row1)
                                    {
                                        foreach (var column in pageColumnNames)
                                        {
                                            string s = worksheet.Cells(rowIndex, column).ToString();
                                            if (!String.IsNullOrEmpty(s))
                                            {
                                                sectionTitles[column] = s.Trim();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        HashSet<string> titles = new HashSet<string>();
                                        string title = worksheet.Cells(rowIndex, titleColumn).ToString();
                                        if (WikiSupport.IsValidTitle(title))
                                        {
                                            if (!titles.Contains(title))
                                            {
                                                titles.Add(title);
                                                StringBuilder sb = new StringBuilder();
                                                for (int i = 0; i < pageColumns.Length; ++i)
                                                {
                                                    string column = pageColumns[i].Name;
                                                    if (sectionTitles.TryGetValue(column, out string sectionTitle))
                                                    {
                                                        sb.AppendLine($"=={sectionTitle}==");
                                                        sb.AppendLine();
                                                    }
                                                    string s = worksheet.Cells(rowIndex, column).ToString();
                                                    string s1 = FormatColumn(s, pageColumns[i]);
                                                    sb.AppendLine(s1);
                                                }
                                                try
                                                {
                                                    wiki.CreatePage(GetPageTitle(pagePrefix, title), editSummary, sb.ToString(), overwrite).Wait();
                                                    WriteError($"Created page {title}");
                                                }
                                                catch (Exception ex)
                                                {
                                                    WriteError($"Unable to create wiki page from row {rowIndex}: {ex.Message}");
                                                }
                                            }
                                            else
                                            {
                                                WriteError($"Duplicate wiki page title {title} at row {rowIndex}...skipping");
                                            }
                                        }
                                        else
                                        {
                                            WriteError($"Invalid wiki page title {title} at row {rowIndex}...skipping.");
                                        }
                                    }
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

        private static string GetPageLink(string pagePrefix, string s1)
        {
            if (String.IsNullOrEmpty(pagePrefix))
            {
                return "[[" + s1 + "]]";
            }
            else
            {
                return "[[" + pagePrefix + " " + s1 + "|" + s1 + "]]";
            }
        }
        public static string GetPageTitle(string pagePrefix, string title)
        {
            if (String.IsNullOrEmpty(pagePrefix))
            {
                return title;
            }
            else
            {
                return pagePrefix + " " + title;
            }

        }

        internal static string[] GetExcelColumns(int column1, int ncolumns)
        {
            if (column1 + ncolumns - 1 > 26 * 26) throw new ApplicationException("Source table has too many columns.");
            List<string> columns = new List<string>();
            for (int i = column1 - 1; i < column1 + ncolumns - 1; i++)
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
                if (!CellReference.IsValidColumn(column)) throw new ApplicationException("Invalid Excel column seen; must be A-Z and AA-ZZ");
            }
            return columns;
        }

        internal static ColumnDef[] GetColumns(int column1, int columns, string[] columnNames, FormatType[] columnFormats, string dateFormat)
        {
            if (columnNames == null) columnNames = GetExcelColumns(column1, columns);
            if (columnFormats == null) columnFormats = new FormatType[0];
            if (columns < columnNames.Length) throw new ApplicationException("The source table does not have enough columns");
            List<ColumnDef> cds = new List<ColumnDef>();
            for (int i = 0; i < columnNames.Length; ++i)
            {
                string name = columnNames[i];
                int index = CellReference.GetColumnIndex(name);
                if (index < column1 || index > column1 + columns - 1) throw new ArgumentException("Column definition outside range for spreadsheet");
                FormatType type = i < columnFormats.Length ? columnFormats[i] : FormatType.None;
                ColumnDef cd = new ColumnDef { Name = name, Format = type };
                if (type == FormatType.Date) cd.FormatParameter = dateFormat;
                cds.Add(cd);
            }
            return cds.ToArray();
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
        private static void WriteError(Exception ex)
        {
            WriteError(ex.Message);
            if (ex.InnerException != null)
                WriteError(ex.InnerException);
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
                "exceltowiki input [--columns clist] [--format flist] [--worksheet sname] [--date-format dformat] [--headers]",
                "                  [--wiki wikiurl --username username --password password] [--overwrite]",
                "                  [--table-title title | --no-table] [--title-column tcolumn]",
                "                  [--page-prefix prefix] [--no-pages] [--page-columns pclist] [--page-format pflist]",
                "                  [--edit-summary summary]",
                "",
                "exceltowiki --test-login --wiki wikiurl --username username --password password",
                "",
                "\tinput         Path to input Excel spreadsheet file (.xls)",
                "\t--columns     Comma separated list of column names, e.g. A,C,D,AA,B",
                "\t--format      Comma separated list of format conversions, e.g. date,,date",
                "\t--worksheet   The name or index of the worksheet in the Excel workbook to be used as a source",
                "\t--worksheet-range   The range of cells to use for the table.  By default the used range of the worksheet",
                "\t--headers     Columns have headers",
                "\t--date-format Format of date output on date conversion of column data",
                "\t--wiki        Destination wiki",
                "\t--username    Username for wiki login",
                "\t--password    Password for wiki login (bot password)",
                "\t--overwrite   Overwrite wiki pages if they already exist",
                "\t--table-title The title of the wiki page containing the table",
                "\t--no-table    Do not write the table into the wiki",
                "\t--title-column    The column in the Excel spreadsheet that contains the title of the wiki page to be created",
                "\t--no-pages    Pages are not created.  Links are still inserted in table if title-column defined.",
                "\t--page-prefix The prefix that is prepended to the title of each wiki page created except for the table page",
                "\t--page-columns    Comma separated list of Excel columns that become wiki page sections. Section name taken from column header.",
                "\t--page-format Comma separated list of formats for excel columns that become wiki page sections",
                "\t--edit-summary    The description of the page edit.  By default \"Modified by exceltowiki\"",
                "\t--test-login  Test the login credentials for the wiki",
                "\t--help        display usage",
                "\t--version     display version",
                "",
                "Where:",
                "",
                "\tclist         Comma separated Excel column list, 1 or 2 alphabetic letters in order that the columns will appear in output.",
                "\tflist         Comma separated format list, one of date or empty string. Position in list corresponds to position in clist.",
                "\t              date format indicates column is date/time.  Converted in output according to date-format argument",
                "\tsname         Excel worksheet name or index.  Sheet indicies are 1-based.",
                "\tdformat       Standard or custom datetime string as described in .NET Documentation. e.g. \"dd-MMM-yy\".  Default is \"g\"",
                "\twikiurl       The absolute URL of the wiki destination for the pages created",
                "\tusername      The username for the wiki indicated by wikiurl",
                "\tpassword      The password for the wiki indicated by wikiurl; this must be the special bot password",
                "\ttitle         The title of the wiki page that contains the table",
                "\tprefix        The prefix for all of the wiki pages created except for the table page",
                "\tpclist        The comma separated list of spreadsheet column that become sections in the wiki page",
                "\tpflist        The comma separated format list of the spreadsheet column that will become sections of the wiki page",
                "\tsummary       The page edit description"
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
