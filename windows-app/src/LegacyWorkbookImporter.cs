using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace PJShoesSlipRecords
{
    /// <summary>
    /// Reads the .xlsx workbook used by the supplied Python application.
    /// Imports the first worksheet named "Lucky Draw" and never modifies it.
    /// </summary>
    public static class LegacyWorkbookImporter
    {
        private static readonly XNamespace Main =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace OfficeRelationships =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PackageRelationships =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        public static List<SlipRecord> Read(string path)
        {
            Package package = null;
            try
            {
                package = Package.Open(path, FileMode.Open, FileAccess.Read);
                XDocument workbook = LoadPart(package, "/xl/workbook.xml");
                XElement sheets = workbook.Root.Element(Main + "sheets");
                if (sheets == null)
                    throw new InvalidDataException("The workbook has no sheets.");

                XElement sheet = sheets.Elements(Main + "sheet")
                    .FirstOrDefault(delegate(XElement item)
                    {
                        XAttribute name = item.Attribute("name");
                        return name != null && name.Value == "Lucky Draw";
                    });
                if (sheet == null)
                    sheet = sheets.Elements(Main + "sheet").FirstOrDefault();
                if (sheet == null)
                    throw new InvalidDataException("The workbook has no worksheets.");

                XAttribute relationshipId =
                    sheet.Attribute(OfficeRelationships + "id");
                if (relationshipId == null)
                    throw new InvalidDataException("A worksheet link is missing.");

                XDocument relationships =
                    LoadPart(package, "/xl/_rels/workbook.xml.rels");
                XElement relation = relationships.Root
                    .Elements(PackageRelationships + "Relationship")
                    .FirstOrDefault(delegate(XElement item)
                    {
                        XAttribute id = item.Attribute("Id");
                        return id != null && id.Value == relationshipId.Value;
                    });
                if (relation == null)
                    throw new InvalidDataException("A worksheet link is invalid.");

                string target = (string)relation.Attribute("Target");
                if (String.IsNullOrEmpty(target))
                    throw new InvalidDataException("A worksheet target is missing.");
                string worksheetPath = ResolveWorksheetPath(target);
                XDocument worksheet = LoadPart(package, worksheetPath);
                List<string> sharedStrings = ReadSharedStrings(package);
                return ReadRows(worksheet, sharedStrings);
            }
            catch (Exception exception)
            {
                StoreException storeException = exception as StoreException;
                if (storeException != null)
                    throw;
                throw new StoreException(
                    "Could not import the selected Excel workbook. Close it in " +
                    "Excel and check that it contains a valid Lucky Draw sheet.\n\n" +
                    exception.Message,
                    exception);
            }
            finally
            {
                if (package != null)
                    package.Close();
            }
        }

        private static XDocument LoadPart(Package package, string partPath)
        {
            Uri requested = PackUriHelper.CreatePartUri(
                new Uri(partPath.TrimStart('/'), UriKind.Relative));
            PackagePart part = package.GetPart(requested);
            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read))
                                return XDocument.Load(System.Xml.XmlReader.Create(stream));
        }

        private static string ResolveWorksheetPath(string target)
        {
            string normalized = target.Replace('\\', '/');
            if (normalized.StartsWith("/", StringComparison.Ordinal))
                return normalized;

            List<string> parts = new List<string> { "xl" };
            foreach (string item in normalized.Split('/'))
            {
                if (item.Length == 0 || item == ".")
                    continue;
                if (item == "..")
                {
                    if (parts.Count > 0)
                        parts.RemoveAt(parts.Count - 1);
                }
                else
                    parts.Add(item);
            }
            return "/" + String.Join("/", parts.ToArray());
        }

        private static List<string> ReadSharedStrings(Package package)
        {
            Uri uri = PackUriHelper.CreatePartUri(
                new Uri("xl/sharedStrings.xml", UriKind.Relative));
            if (!package.PartExists(uri))
                return new List<string>();

            XDocument document = LoadPart(package, "/xl/sharedStrings.xml");
            List<string> strings = new List<string>();
            foreach (XElement item in document.Root.Elements(Main + "si"))
            {
                strings.Add(String.Concat(item.Descendants(Main + "t")
                    .Select(delegate(XElement text) { return text.Value; })
                    .ToArray()));
            }
            return strings;
        }

        private static List<SlipRecord> ReadRows(
            XDocument worksheet, List<string> sharedStrings)
        {
            XElement data = worksheet.Root.Element(Main + "sheetData");
            if (data == null)
                throw new InvalidDataException("The selected sheet has no data.");

            List<SlipRecord> records = new List<SlipRecord>();
            bool headerFound = false;
            foreach (XElement row in data.Elements(Main + "row"))
            {
                Dictionary<int, string> cells = new Dictionary<int, string>();
                foreach (XElement cell in row.Elements(Main + "c"))
                {
                    XAttribute reference = cell.Attribute("r");
                    if (reference == null)
                        continue;
                    int column = ColumnIndex(reference.Value);
                    cells[column] = ReadCell(cell, sharedStrings);
                }

                if (!headerFound)
                {
                    headerFound = true;
                    string first = CellAt(cells, 0).Trim();
                    if (!String.Equals(first, "slip_no",
                            StringComparison.OrdinalIgnoreCase) &&
                        !String.Equals(first, "Slip No",
                            StringComparison.OrdinalIgnoreCase))
                        throw new InvalidDataException(
                            "The Lucky Draw sheet does not have the expected columns.");
                    continue;
                }

                string slipText = CellAt(cells, 0).Trim();
                if (slipText.Length == 0)
                    continue;

                int slipNo;
                if (!Int32.TryParse(slipText, NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out slipNo) || slipNo < 1)
                    throw new InvalidDataException(
                        "A row contains an invalid slip number: " + slipText);

                string customerName = CellAt(cells, 1).Trim();
                string phoneNo = CellAt(cells, 2).Trim();
                if (customerName.Length == 0 || phoneNo.Length == 0)
                    throw new InvalidDataException(
                        "Slip " + slipNo + " has a blank customer name or phone.");

                DateTime createdAt;
                string createdText = CellAt(cells, 3).Trim();
                if (!DateTime.TryParse(createdText, CultureInfo.InvariantCulture,
                        DateTimeStyles.AllowWhiteSpaces, out createdAt))
                {
                    double serial;
                    if (!Double.TryParse(createdText, NumberStyles.Float,
                            CultureInfo.InvariantCulture, out serial))
                        throw new InvalidDataException(
                            "Slip " + slipNo + " has an invalid creation date.");
                    createdAt = DateTime.FromOADate(serial);
                }

                records.Add(new SlipRecord
                {
                    SlipNo = slipNo,
                    CustomerName = customerName,
                    PhoneNo = phoneNo,
                    CreatedAt = createdAt,
                    EntryDate = createdAt.Date
                });
            }

            HashSet<int> seen = new HashSet<int>();
            foreach (SlipRecord record in records)
            {
                if (!seen.Add(record.SlipNo))
                    throw new InvalidDataException(
                        "The workbook contains duplicate slip number " +
                        record.SlipNo + ".");
            }
            return records;
        }

        private static string ReadCell(
            XElement cell, List<string> sharedStrings)
        {
            XAttribute type = cell.Attribute("t");
            if (type != null && type.Value == "inlineStr")
            {
                XElement inlineText = cell.Element(Main + "is");
                return inlineText == null
                    ? String.Empty
                    : String.Concat(inlineText.Descendants(Main + "t")
                        .Select(delegate(XElement text) { return text.Value; })
                        .ToArray());
            }

            XElement value = cell.Element(Main + "v");
            if (value == null)
                return String.Empty;
            if (type != null && type.Value == "s")
            {
                int index;
                if (!Int32.TryParse(value.Value, out index) ||
                    index < 0 || index >= sharedStrings.Count)
                    throw new InvalidDataException(
                        "The workbook contains an invalid shared-string index.");
                return sharedStrings[index];
            }
            return value.Value;
        }

        private static int ColumnIndex(string cellReference)
        {
            int result = 0;
            for (int index = 0; index < cellReference.Length; index++)
            {
                char character = cellReference[index];
                if (character < 'A' || character > 'Z')
                {
                    if (character < 'a' || character > 'z')
                        break;
                    character = Char.ToUpperInvariant(character);
                }
                result = (result * 26) + (character - 'A' + 1);
            }
            return result - 1;
        }

        private static string CellAt(Dictionary<int, string> cells, int index)
        {
            string value;
            return cells.TryGetValue(index, out value) ? value : String.Empty;
        }
    }
}
