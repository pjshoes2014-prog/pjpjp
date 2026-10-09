using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace PJShoesSlipRecords
{
    public sealed class StoreException : Exception
    {
        public StoreException(string message) : base(message) { }
        public StoreException(string message, Exception inner)
            : base(message, inner) { }
    }

    /// <summary>
    /// Stores records as Excel 2003 SpreadsheetML. It is a plain, local XML
    /// file that Excel can open without Office automation or extra libraries.
    /// </summary>
    public sealed class SpreadsheetStore
    {
        private const string SpreadsheetNamespace =
            "urn:schemas-microsoft-com:office:spreadsheet";
        private const string OfficeNamespace =
            "urn:schemas-microsoft-com:office:office";
        private const string ExcelNamespace =
            "urn:schemas-microsoft-com:office:excel";
        private const string WorksheetName = "Customer Slips";

        private static readonly XNamespace Ss = SpreadsheetNamespace;
        private static readonly XNamespace O = OfficeNamespace;
        private static readonly XNamespace X = ExcelNamespace;

        private static readonly string[] CurrentHeaders =
        {
            "Slip No", "Customer Name", "Phone Number", "Created At",
            "Entry Date"
        };

        private static readonly string[] LegacyHeaders =
        {
            "Slip No", "Customer Name", "Phone Number", "Created At"
        };

        public string FilePath { get; private set; }

        public SpreadsheetStore(string filePath)
        {
            FilePath = filePath;
            if (!File.Exists(FilePath))
                SaveAll(new List<SlipRecord>());
        }

        public List<SlipRecord> LoadAll()
        {
            XDocument workbook = LoadWorkbook();
            XElement table = FindTable(workbook);
            List<XElement> rows = table.Elements(Ss + "Row").ToList();
            if (rows.Count == 0)
                throw new StoreException("The customer workbook has no header row.");

            List<string> headers = ReadRow(rows[0]);
            bool isLegacy = headers.Count == LegacyHeaders.Length &&
                HeadersMatch(headers, LegacyHeaders);
            if (!isLegacy && !HeadersMatch(headers, CurrentHeaders))
                throw new StoreException(
                    "The customer workbook has an unsupported column layout. " +
                    "Make a backup before changing its headers.");

            List<SlipRecord> records = new List<SlipRecord>();
            HashSet<int> seen = new HashSet<int>();
            for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++)
            {
                List<string> values = ReadRow(rows[rowIndex]);
                if (values.All(delegate(string value) { return IsBlank(value); }))
                    continue;

                int slipNo;
                DateTime createdAt;
                string createdText = ValueAt(values, 3);
                if (!Int32.TryParse(ValueAt(values, 0), NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out slipNo) ||
                    slipNo < 1 ||
                    IsBlank(ValueAt(values, 1)) ||
                    IsBlank(ValueAt(values, 2)) ||
                    !DateTime.TryParse(createdText, CultureInfo.InvariantCulture,
                        DateTimeStyles.AllowWhiteSpaces, out createdAt))
                {
                    throw new StoreException(
                        "Invalid or incomplete data on workbook row " +
                        (rowIndex + 1).ToString(CultureInfo.InvariantCulture) + ".");
                }
                if (!seen.Add(slipNo))
                    throw new StoreException(
                        "The workbook contains duplicate slip number " + slipNo + ".");

                DateTime entryDate = createdAt.Date;
                if (!isLegacy &&
                    !DateTime.TryParse(ValueAt(values, 4), CultureInfo.InvariantCulture,
                        DateTimeStyles.AllowWhiteSpaces, out entryDate))
                {
                    throw new StoreException(
                        "Invalid entry date on workbook row " +
                        (rowIndex + 1).ToString(CultureInfo.InvariantCulture) + ".");
                }

                records.Add(new SlipRecord
                {
                    SlipNo = slipNo,
                    CustomerName = ValueAt(values, 1).Trim(),
                    PhoneNo = ValueAt(values, 2).Trim(),
                    CreatedAt = createdAt,
                    EntryDate = entryDate.Date
                });
            }

            records.Sort(delegate(SlipRecord left, SlipRecord right)
            {
                return left.SlipNo.CompareTo(right.SlipNo);
            });

            // Older SpreadsheetML files are upgraded only after every row has
            // been read and validated. A backup is retained beside the file.
            if (isLegacy || headers.Count != CurrentHeaders.Length)
                SaveAll(records);

            return records;
        }

        public void SaveAll(List<SlipRecord> records)
        {
            ValidateRecords(records);
            XDocument workbook = CreateWorkbook(records);
            string directory = Path.GetDirectoryName(FilePath);
            if (!String.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            string temporaryPath = FilePath + ".tmp";
            string backupPath = FilePath + ".bak";
            try
            {
                workbook.Save(temporaryPath);
                if (File.Exists(FilePath))
                    File.Replace(temporaryPath, FilePath, backupPath);
                else
                    File.Move(temporaryPath, FilePath);
            }
            catch (Exception exception)
            {
                throw new StoreException(
                    "Could not save the customer workbook. Close it in Excel, " +
                    "check that the data folder is writable, and try again.\n\n" +
                    exception.Message,
                    exception);
            }
            finally
            {
                if (File.Exists(temporaryPath))
                {
                    try { File.Delete(temporaryPath); }
                    catch { }
                }
            }
        }

        private static void ValidateRecords(List<SlipRecord> records)
        {
            HashSet<int> seen = new HashSet<int>();
            foreach (SlipRecord record in records)
            {
                if (record == null || record.SlipNo < 1 ||
                    IsBlank(record.CustomerName) ||
                    IsBlank(record.PhoneNo) ||
                    !seen.Add(record.SlipNo))
                    throw new StoreException(
                        "The records contain a blank field or duplicate slip number.");
            }
        }

        private XDocument LoadWorkbook()
        {
            try
            {
                return XDocument.Load(FilePath);
            }
            catch (Exception exception)
            {
                throw new StoreException(
                    "Could not read the customer workbook. Close it in Excel and " +
                    "check that the file is valid:\n\n" + exception.Message,
                    exception);
            }
        }

        private static XElement FindTable(XDocument workbook)
        {
            XElement worksheet = workbook.Root
                .Elements(Ss + "Worksheet")
                .FirstOrDefault(delegate(XElement element)
                {
                    XAttribute name = element.Attribute(Ss + "Name");
                    return name != null && name.Value == WorksheetName;
                });
            if (worksheet == null)
                throw new StoreException(
                    "The workbook does not contain a '" + WorksheetName + "' sheet.");

            XElement table = worksheet.Element(Ss + "Table");
            if (table == null)
                throw new StoreException("The customer-slips sheet has no data table.");
            return table;
        }

        private static bool HeadersMatch(List<string> actual, string[] expected)
        {
            if (actual.Count < expected.Length)
                return false;
            for (int index = 0; index < expected.Length; index++)
            {
                if (!String.Equals(actual[index], expected[index],
                    StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            return true;
        }

        private static List<string> ReadRow(XElement row)
        {
            List<string> values = new List<string>();
            foreach (XElement cell in row.Elements(Ss + "Cell"))
            {
                XElement data = cell.Element(Ss + "Data");
                values.Add(data == null ? String.Empty : data.Value);
            }
            return values;
        }

        private static string ValueAt(List<string> values, int index)
        {
            return index < values.Count && values[index] != null
                ? values[index]
                : String.Empty;
        }

        private static bool IsBlank(string value)
        {
            return value == null || value.Trim().Length == 0;
        }

        private static XDocument CreateWorkbook(List<SlipRecord> records)
        {
            XElement header = new XElement(Ss + "Row");
            foreach (string name in CurrentHeaders)
                header.Add(CreateCell("String", name));

            XElement table = new XElement(Ss + "Table",
                new XElement(Ss + "Column", new XAttribute(Ss + "Width", "70")),
                new XElement(Ss + "Column", new XAttribute(Ss + "Width", "190")),
                new XElement(Ss + "Column", new XAttribute(Ss + "Width", "140")),
                new XElement(Ss + "Column", new XAttribute(Ss + "Width", "145")),
                header);

            foreach (SlipRecord record in records.OrderBy(
                delegate(SlipRecord item) { return item.SlipNo; }))
            {
                table.Add(new XElement(Ss + "Row",
                    CreateCell("Number", record.SlipNo.ToString(
                        CultureInfo.InvariantCulture)),
                    CreateCell("String", record.CustomerName),
                    CreateCell("String", record.PhoneNo),
                    CreateCell("String", record.CreatedAt.ToString(
                        "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)),
                    CreateCell("String", record.EntryDate.ToString(
                        "yyyy-MM-dd", CultureInfo.InvariantCulture))));
            }

            XElement root = new XElement(Ss + "Workbook",
                new XAttribute(XNamespace.Xmlns + "o", O),
                new XAttribute(XNamespace.Xmlns + "x", X),
                new XAttribute(XNamespace.Xmlns + "ss", Ss),
                new XElement(O + "DocumentProperties",
                    new XElement(O + "Author", "PJ Shoes"),
                    new XElement(O + "Title", "PJ Shoes Customer Slips")),
                new XElement(X + "ExcelWorkbook",
                    new XElement(X + "ProtectStructure", "False"),
                    new XElement(X + "ProtectWindows", "False")),
                new XElement(Ss + "Worksheet",
                    new XAttribute(Ss + "Name", WorksheetName),
                    table));

            return new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XProcessingInstruction(
                    "mso-application", "progid=\"Excel.Sheet\""),
                root);
        }

        private static XElement CreateCell(string type, string value)
        {
            return new XElement(Ss + "Cell",
                new XElement(Ss + "Data",
                    new XAttribute(Ss + "Type", type),
                    value ?? String.Empty));
        }
    }
}
