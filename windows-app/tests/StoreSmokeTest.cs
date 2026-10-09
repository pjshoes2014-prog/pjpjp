using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using PJShoesSlipRecords;

internal static class StoreSmokeTest
{
    private static int Main(string[] args)
    {
        string directory = Path.Combine(
            Path.GetTempPath(), "pj-shoes-store-smoke-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            TestCurrentStore(directory);
            TestLegacySpreadsheetXml(directory);
            if (args.Length > 0)
                TestLegacyXlsxImport(args[0]);
            Console.WriteLine("All storage smoke checks passed.");
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception.ToString());
            return 1;
        }
        finally
        {
            try { Directory.Delete(directory, true); }
            catch { }
        }
    }

    private static void TestCurrentStore(string directory)
    {
        string path = Path.Combine(directory, "current.xml");
        SpreadsheetStore store = new SpreadsheetStore(path);
        Assert(store.LoadAll().Count == 0, "new store should start empty");

        SlipRecord record = new SlipRecord
        {
            SlipNo = 18,
            CustomerName = "Ayesha & Sons",
            PhoneNo = "03001234567",
            EntryDate = new DateTime(2026, 10, 9),
            CreatedAt = new DateTime(2026, 10, 9, 11, 30, 45)
        };
        store.SaveAll(new List<SlipRecord> { record });
        List<SlipRecord> loaded = store.LoadAll();
        Assert(loaded.Count == 1, "saved row count");
        Assert(loaded[0].CustomerName == record.CustomerName, "name round trip");
        Assert(loaded[0].PhoneNo == record.PhoneNo, "phone string preservation");
        Assert(loaded[0].EntryDate == record.EntryDate, "entry date round trip");
        Assert(loaded[0].CreatedAt == record.CreatedAt, "created date round trip");

        loaded[0].CustomerName = "Updated & Saved";
        store.SaveAll(loaded);
        Assert(File.Exists(path + ".bak"), "prior file backup should exist");
        Assert(store.LoadAll()[0].CustomerName == "Updated & Saved",
            "record update round trip");

        bool rejectedDuplicate = false;
        try
        {
            store.SaveAll(new List<SlipRecord> { record, record.Copy() });
        }
        catch (StoreException)
        {
            rejectedDuplicate = true;
        }
        Assert(rejectedDuplicate, "duplicate slip numbers must be rejected");
    }

    private static void TestLegacySpreadsheetXml(string directory)
    {
        string path = Path.Combine(directory, "legacy.xml");
        string xml =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<ss:Workbook xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">" +
            "<ss:Worksheet ss:Name=\"Customer Slips\"><ss:Table>" +
            "<ss:Row>" +
            Cell("Slip No") + Cell("Customer Name") +
            Cell("Phone Number") + Cell("Created At") +
            "</ss:Row><ss:Row>" +
            Cell("31") + Cell("Legacy Customer") +
            Cell("03005550123") + Cell("2026-08-01 09:10:11") +
            "</ss:Row></ss:Table></ss:Worksheet></ss:Workbook>";
        File.WriteAllText(path, xml);

        SpreadsheetStore store = new SpreadsheetStore(path);
        List<SlipRecord> records = store.LoadAll();
        Assert(records.Count == 1, "legacy XML record count");
        Assert(records[0].SlipNo == 31, "legacy slip number");
        Assert(records[0].EntryDate == new DateTime(2026, 8, 1),
            "legacy entry date defaults to original creation date");
        Assert(File.Exists(path + ".bak"), "legacy source backup should exist");
        Assert(store.LoadAll()[0].CustomerName == "Legacy Customer",
            "migrated record remains readable");
    }

    private static string Cell(string value)
    {
        return "<ss:Cell><ss:Data ss:Type=\"String\">" +
            System.Security.SecurityElement.Escape(value) +
            "</ss:Data></ss:Cell>";
    }

    private static void TestLegacyXlsxImport(string path)
    {
        List<SlipRecord> imported = LegacyWorkbookImporter.Read(path);
        Assert(imported.Count == 1, "legacy XLSX record count");
        Assert(imported[0].SlipNo == 18, "legacy XLSX slip number");
        Assert(imported[0].CustomerName == "Ayesha & Sons",
            "legacy XLSX shared string");
        Assert(imported[0].PhoneNo == "03001234567",
            "legacy XLSX leading-zero phone");
        Assert(imported[0].CreatedAt ==
            DateTime.Parse("2026-09-11 13:25:00", CultureInfo.InvariantCulture),
            "legacy XLSX creation timestamp");
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException("Failed: " + message);
    }
}
