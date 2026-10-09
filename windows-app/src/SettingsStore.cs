using System;
using System.IO;
using System.Xml.Linq;

namespace PJShoesSlipRecords
{
    public sealed class SettingsStore
    {
        private readonly string _filePath;

        public SettingsStore(string filePath)
        {
            _filePath = filePath;
        }

        public string LoadPrinterName()
        {
            if (!File.Exists(_filePath))
                return String.Empty;
            try
            {
                XDocument settings = XDocument.Load(_filePath);
                XElement printer = settings.Root == null
                    ? null
                    : settings.Root.Element("PrinterName");
                return printer == null ? String.Empty : printer.Value;
            }
            catch (Exception exception)
            {
                throw new StoreException(
                    "Could not read printer settings:\n\n" + exception.Message,
                    exception);
            }
        }

        public void SavePrinterName(string printerName)
        {
            string directory = Path.GetDirectoryName(_filePath);
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
            string temporaryPath = _filePath + ".tmp";
            string backupPath = _filePath + ".bak";
            XDocument document = new XDocument(
                new XElement("Settings",
                    new XElement("PrinterName", printerName ?? String.Empty)));
            try
            {
                document.Save(temporaryPath);
                if (File.Exists(_filePath))
                    File.Replace(temporaryPath, _filePath, backupPath);
                else
                    File.Move(temporaryPath, _filePath);
            }
            catch (Exception exception)
            {
                throw new StoreException(
                    "Could not save printer settings:\n\n" + exception.Message,
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
    }
}
