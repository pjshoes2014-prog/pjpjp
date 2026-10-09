using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace PJShoesSlipRecords
{
    public static class PrinterService
    {
        private const string LogoResourceName =
            "PJShoesSlipRecords.assets.pj-shoes-logo.jpg";

        public static void PrintSlip(SlipRecord record, string printerName)
        {
            if (record == null)
                throw new ArgumentNullException("record");
            PrintDocument document = CreateDocument(printerName);
            document.DocumentName = "PJ Shoes Customer Slip #" + record.SlipNo;
            document.PrintPage += delegate(object sender, PrintPageEventArgs args)
            {
                DrawReceipt(args.Graphics, args.MarginBounds, record, false);
                args.HasMorePages = false;
            };
            Print(document);
        }

        public static void PrintTest(string printerName)
        {
            PrintDocument document = CreateDocument(printerName);
            document.DocumentName = "PJ Shoes Printer Test";
            document.PrintPage += delegate(object sender, PrintPageEventArgs args)
            {
                DrawReceipt(args.Graphics, args.MarginBounds, null, true);
                args.HasMorePages = false;
            };
            Print(document);
        }

        private static PrintDocument CreateDocument(string printerName)
        {
            if (String.IsNullOrEmpty(printerName == null ? null : printerName.Trim()))
                throw new InvalidOperationException(
                    "Choose an installed printer before printing.");

            PrintDocument document = new PrintDocument();
            document.PrinterSettings.PrinterName = printerName;
            if (!document.PrinterSettings.IsValid)
            {
                document.Dispose();
                throw new InvalidOperationException(
                    "The selected printer is no longer installed or available.");
            }
            document.PrintController = new StandardPrintController();
            document.DefaultPageSettings.Margins = new Margins(10, 10, 10, 10);
            return document;
        }

        private static void Print(PrintDocument document)
        {
            try
            {
                document.Print();
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Could not print to '" + document.PrinterSettings.PrinterName +
                    "'. Check that the printer is online and configured with the " +
                    "intended paper size.\n\n" + exception.Message,
                    exception);
            }
            finally
            {
                document.Dispose();
            }
        }

        private static void DrawReceipt(
            Graphics graphics, Rectangle bounds, SlipRecord record, bool testPage)
        {
            graphics.PageUnit = GraphicsUnit.Display;
            float x = bounds.Left;
            float y = bounds.Top;
            float width = bounds.Width;
            using (StringFormat centered = new StringFormat())
            {
                centered.Alignment = StringAlignment.Center;
                centered.LineAlignment = StringAlignment.Near;

                DrawLogo(graphics, bounds, ref y);
                using (Font brand = new Font("Arial", 14.0f, FontStyle.Bold))
                    y = DrawLine(graphics, "PJ SHOES", brand, x, y, width, centered);
                using (Font title = new Font("Arial", 9.0f, FontStyle.Bold))
                    y = DrawLine(graphics, "LUCKY DRAW TOKEN", title,
                        x, y + 2, width, centered);
                y = DrawRule(graphics, x, y + 4, width);

                if (testPage)
                {
                    using (Font label = new Font("Arial", 11.0f, FontStyle.Bold))
                        y = DrawLine(graphics, "TEST PRINT", label,
                            x, y, width, centered);
                    using (Font body = new Font("Arial", 8.0f, FontStyle.Regular))
                        y = DrawLine(graphics,
                            "Printer check only. No customer slip was saved.",
                            body, x, y + 4, width, centered);
                    y = DrawRule(graphics, x, y + 4, width);
                    using (Font small = new Font("Arial", 7.0f, FontStyle.Regular))
                        DrawLine(graphics, DateTime.Now.ToString(
                            "dd MMM yyyy  hh:mm tt", CultureInfo.InvariantCulture),
                            small, x, y, width, centered);
                    return;
                }

                using (Font small = new Font("Arial", 7.0f, FontStyle.Regular))
                    y = DrawLine(graphics, record.EntryDate.ToString(
                        "dd MMM yyyy", CultureInfo.InvariantCulture),
                        small, x, y, width, centered);
                using (Font slip = new Font("Arial", 12.0f, FontStyle.Bold))
                    y = DrawLine(graphics, "SLIP NO: " +
                        record.SlipNo.ToString(CultureInfo.InvariantCulture),
                        slip, x, y + 2, width, centered);
                using (Font issued = new Font("Arial", 7.0f, FontStyle.Regular))
                    y = DrawLine(graphics, "Issued: " + record.CreatedAt.ToString(
                        "dd MMM yyyy  hh:mm tt", CultureInfo.InvariantCulture),
                        issued, x, y, width, centered);
                y = DrawRule(graphics, x, y + 3, width);

                using (Font label = new Font("Arial", 7.0f, FontStyle.Bold))
                    y = DrawLine(graphics, "NAME", label, x, y, width, centered);
                using (Font value = new Font("Arial", 10.0f, FontStyle.Bold))
                    y = DrawLine(graphics, record.CustomerName, value,
                        x, y, width, centered);
                using (Font label = new Font("Arial", 7.0f, FontStyle.Bold))
                    y = DrawLine(graphics, "CELL #", label,
                        x, y + 2, width, centered);
                using (Font value = new Font("Arial", 10.0f, FontStyle.Bold))
                    y = DrawLine(graphics, record.PhoneNo, value,
                        x, y, width, centered);

                y = DrawRule(graphics, x, y + 4, width);
                using (Font tag = new Font("Arial", 8.0f, FontStyle.Regular))
                    y = DrawLine(graphics, "@bill", tag, x, y, width, centered);
                using (Font campaign = new Font("Arial", 10.0f, FontStyle.Bold))
                    y = DrawLine(graphics, "Draw on Dec 2026",
                        campaign, x, y + 2, width, centered);
                y = DrawRule(graphics, x, y + 3, width);
                using (Font note = new Font("Arial", 7.0f, FontStyle.Regular))
                {
                    y = DrawLine(graphics, "Keep this ticket safe for the",
                        note, x, y, width, centered);
                    DrawLine(graphics, "lucky draw! Good luck!",
                        note, x, y, width, centered);
                }
            }
        }

        private static float DrawLine(
            Graphics graphics, string text, Font font, float x, float y,
            float width, StringFormat centered)
        {
            RectangleF layout = new RectangleF(x, y, width, 500);
            SizeF measured = graphics.MeasureString(text, font, (int)width);
            graphics.DrawString(text, font, Brushes.Black, layout, centered);
            return y + measured.Height + 2;
        }

        private static float DrawRule(
            Graphics graphics, float x, float y, float width)
        {
            using (Pen pen = new Pen(Color.Black, 0.7f))
                graphics.DrawLine(pen, x + 3, y, x + width - 3, y);
            return y + 6;
        }

        private static void DrawLogo(Graphics graphics, Rectangle bounds, ref float y)
        {
            Stream stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream(LogoResourceName);
            if (stream == null)
                return;
            using (stream)
            using (Image logo = Image.FromStream(stream))
            {
                float maxWidth = bounds.Width * 0.34f;
                float maxHeight = 50.0f;
                float scale = Math.Min(maxWidth / logo.Width, maxHeight / logo.Height);
                float drawWidth = logo.Width * scale;
                float drawHeight = logo.Height * scale;
                float left = bounds.Left + (bounds.Width - drawWidth) / 2.0f;
                graphics.DrawImage(logo, left, y, drawWidth, drawHeight);
                y += drawHeight + 4;
            }
        }
    }
}
