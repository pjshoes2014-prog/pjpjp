using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace PJShoesSlipRecords
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            try
            {
                string dataDirectory = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.ApplicationData),
                    "PJ Shoes\\Slip Records");
                if (!Directory.Exists(dataDirectory))
                    Directory.CreateDirectory(dataDirectory);

                string workbookPath = Path.Combine(
                    dataDirectory, "pj_shoes_slips.xml");
                string settingsPath = Path.Combine(dataDirectory, "settings.xml");
                SpreadsheetStore store = new SpreadsheetStore(workbookPath);
                SettingsStore settings = new SettingsStore(settingsPath);
                List<SlipRecord> records = store.LoadAll();

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm(
                    store, settings, dataDirectory, records));
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    "PJ Shoes Slip Records could not start.\n\n" +
                        exception.Message,
                    "PJ Shoes Slip Records",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
