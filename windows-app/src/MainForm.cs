using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PJShoesSlipRecords
{
    public sealed class MainForm : Form
    {
        private readonly SpreadsheetStore _store;
        private readonly SettingsStore _settings;
        private readonly string _dataDirectory;
        private List<SlipRecord> _records;
        private SlipRecord _loadedRecord;
        private bool _loadingFields;
        private bool _dirty;

        private NumericUpDown _slipNumber;
        private TextBox _customerName;
        private TextBox _phoneNumber;
        private DateTimePicker _entryDate;
        private TextBox _searchBox;
        private DataGridView _slipsGrid;
        private Label _recordCount;
        private Label _entryStatus;
        private ComboBox _printerList;
        private Label _printerStatus;

        public MainForm(
            SpreadsheetStore store, SettingsStore settings,
            string dataDirectory, List<SlipRecord> records)
        {
            _store = store;
            _settings = settings;
            _dataDirectory = dataDirectory;
            _records = records;

            Text = "PJ Shoes | Slip Printing & Phone Records";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(980, 650);
            Size = new Size(1160, 780);
            BackColor = Color.FromArgb(243, 240, 233);
            Font = new Font("Segoe UI", 9.0f, FontStyle.Regular);
            AutoScaleMode = AutoScaleMode.Font;

            BuildInterface();
            BindDirtyTracking();
            RefreshGrid(HighestSlip() + 1);
            StartNewEntry(false);
            RefreshPrinterList();

            FormClosing += MainFormClosing;
        }

        private void BuildInterface()
        {
            TabControl tabs = new TabControl();
            tabs.Dock = DockStyle.Fill;
            tabs.Padding = new Point(18, 5);
            tabs.TabPages.Add(BuildSlipsPage());
            tabs.TabPages.Add(BuildPrinterPage());
            Controls.Add(tabs);
        }

        private TabPage BuildSlipsPage()
        {
            TabPage page = new TabPage("Slips & customers");
            page.BackColor = BackColor;

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.Padding = new Padding(16, 12, 16, 12);
            layout.RowCount = 2;
            layout.ColumnCount = 1;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 66));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.Controls.Add(CreateHeader(
                "PJ SHOES", "Customer entry, printed slips, and saved records"), 0, 0);

            SplitContainer split = new SplitContainer();
            split.Dock = DockStyle.Fill;
            split.Orientation = Orientation.Vertical;
            split.SplitterDistance = 365;
            split.Panel1MinSize = 315;
            split.Panel2MinSize = 500;
            split.BackColor = Color.FromArgb(228, 221, 210);
            split.Panel1.Padding = new Padding(0, 0, 8, 0);
            split.Panel2.Padding = new Padding(8, 0, 0, 0);
            BuildEntryPanel(split.Panel1);
            BuildRecordsPanel(split.Panel2);
            layout.Controls.Add(split, 0, 1);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildPrinterPage()
        {
            TabPage page = new TabPage("Printer settings");
            page.BackColor = BackColor;

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.Padding = new Padding(36);
            layout.ColumnCount = 1;
            layout.RowCount = 6;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 78));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.Controls.Add(CreateHeader(
                "Printer settings", "Choose and save a Windows printer for receipt slips"),
                0, 0);

            Label printerLabel = new Label();
            printerLabel.Text = "Installed printers";
            printerLabel.Dock = DockStyle.Fill;
            printerLabel.TextAlign = ContentAlignment.BottomLeft;
            layout.Controls.Add(printerLabel, 0, 1);

            _printerList = new ComboBox();
            _printerList.Dock = DockStyle.Fill;
            _printerList.DropDownStyle = ComboBoxStyle.DropDownList;
            _printerList.Font = new Font("Segoe UI", 11.0f, FontStyle.Regular);
            layout.Controls.Add(_printerList, 0, 2);

            FlowLayoutPanel actions = new FlowLayoutPanel();
            actions.Dock = DockStyle.Fill;
            actions.Padding = new Padding(0, 12, 0, 0);
            actions.Controls.Add(CreateButton(
                "Save printer", SavePrinter, true, 140, 42));
            actions.Controls.Add(CreateButton(
                "Test print", TestPrint, false, 140, 42));
            actions.Controls.Add(CreateButton(
                "Refresh printers", RefreshPrinterList, false, 160, 42));
            layout.Controls.Add(actions, 0, 3);

            _printerStatus = new Label();
            _printerStatus.Dock = DockStyle.Fill;
            _printerStatus.TextAlign = ContentAlignment.MiddleLeft;
            _printerStatus.ForeColor = Color.FromArgb(75, 69, 59);
            layout.Controls.Add(_printerStatus, 0, 4);

            Label help = new Label();
            help.Text =
                "For an 80mm thermal printer, choose the 80mm paper size in " +
                "the printer's Windows preferences. The test slip is printed " +
                "without creating a customer record.";
            help.Dock = DockStyle.Top;
            help.MaximumSize = new Size(780, 60);
            help.ForeColor = Color.FromArgb(112, 107, 97);
            layout.Controls.Add(help, 0, 5);

            page.Controls.Add(layout);
            return page;
        }

        private void BuildEntryPanel(Control parent)
        {
            GroupBox group = new GroupBox();
            group.Text = "Customer slip";
            group.Dock = DockStyle.Fill;
            group.Padding = new Padding(14);
            group.ForeColor = Color.FromArgb(124, 33, 29);
            parent.Controls.Add(group);

            TableLayoutPanel fields = new TableLayoutPanel();
            fields.Dock = DockStyle.Top;
            fields.AutoSize = true;
            fields.ColumnCount = 2;
            fields.RowCount = 8;
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 16));
            fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            group.Controls.Add(fields);

            Label intro = new Label();
            intro.Text = "Enter a customer and save the slip before printing.";
            intro.Dock = DockStyle.Fill;
            intro.ForeColor = Color.FromArgb(112, 107, 97);
            intro.TextAlign = ContentAlignment.MiddleLeft;
            fields.Controls.Add(intro, 0, 0);
            fields.SetColumnSpan(intro, 2);

            _slipNumber = new NumericUpDown();
            _slipNumber.Minimum = 1;
            _slipNumber.Maximum = 99999999;
            _slipNumber.TextAlign = HorizontalAlignment.Center;
            _slipNumber.Dock = DockStyle.Fill;
            _slipNumber.Font = new Font("Segoe UI", 11.0f, FontStyle.Bold);
            AddField(fields, 1, "Slip / ticket ID", _slipNumber);

            _customerName = new TextBox();
            _customerName.Dock = DockStyle.Fill;
            _customerName.MaxLength = 120;
            _customerName.Font = new Font("Segoe UI", 11.0f, FontStyle.Regular);
            AddField(fields, 2, "Customer name", _customerName);

            _phoneNumber = new TextBox();
            _phoneNumber.Dock = DockStyle.Fill;
            _phoneNumber.MaxLength = 40;
            _phoneNumber.Font = new Font("Segoe UI", 11.0f, FontStyle.Regular);
            AddField(fields, 3, "Phone number", _phoneNumber);

            _entryDate = new DateTimePicker();
            _entryDate.Dock = DockStyle.Fill;
            _entryDate.Format = DateTimePickerFormat.Short;
            _entryDate.Font = new Font("Segoe UI", 11.0f, FontStyle.Regular);
            AddField(fields, 4, "Entry date", _entryDate);

            _entryStatus = new Label();
            _entryStatus.Dock = DockStyle.Fill;
            _entryStatus.ForeColor = Color.FromArgb(112, 107, 97);
            _entryStatus.TextAlign = ContentAlignment.MiddleLeft;
            fields.Controls.Add(_entryStatus, 0, 5);
            fields.SetColumnSpan(_entryStatus, 2);

            Button save = CreateButton("Save slip", SaveSlip, true, 135, 40);
            fields.Controls.Add(save, 0, 6);
            fields.SetColumnSpan(save, 2);

            FlowLayoutPanel actions = new FlowLayoutPanel();
            actions.Dock = DockStyle.Fill;
            actions.WrapContents = false;
            actions.Controls.Add(CreateButton(
                "Save & print", SaveAndPrint, false, 145, 38));
            actions.Controls.Add(CreateButton(
                "New slip", NewSlip, false, 105, 38));
            fields.Controls.Add(actions, 0, 7);
            fields.SetColumnSpan(actions, 2);
        }

        private void BuildRecordsPanel(Control parent)
        {
            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.RowCount = 4;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            parent.Controls.Add(layout);

            TableLayoutPanel search = new TableLayoutPanel();
            search.Dock = DockStyle.Fill;
            search.ColumnCount = 3;
            search.RowCount = 1;
            search.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 58));
            search.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            search.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 66));
            Label searchLabel = new Label();
            searchLabel.Text = "Search";
            searchLabel.TextAlign = ContentAlignment.MiddleLeft;
            searchLabel.Dock = DockStyle.Fill;
            search.Controls.Add(searchLabel, 0, 0);
            _searchBox = new TextBox();
            _searchBox.Dock = DockStyle.Fill;
            _searchBox.Margin = new Padding(0, 7, 8, 7);
            _searchBox.TextChanged += SearchChanged;
            search.Controls.Add(_searchBox, 1, 0);
            search.Controls.Add(CreateButton(
                "Clear", ClearSearch, false, 62, 30), 2, 0);
            layout.Controls.Add(search, 0, 0);

            _slipsGrid = new DataGridView();
            _slipsGrid.Dock = DockStyle.Fill;
            _slipsGrid.ReadOnly = true;
            _slipsGrid.AllowUserToAddRows = false;
            _slipsGrid.AllowUserToDeleteRows = false;
            _slipsGrid.AllowUserToResizeRows = false;
            _slipsGrid.MultiSelect = false;
            _slipsGrid.RowHeadersVisible = false;
            _slipsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _slipsGrid.AutoGenerateColumns = false;
            _slipsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _slipsGrid.BackgroundColor = Color.White;
            _slipsGrid.BorderStyle = BorderStyle.FixedSingle;
            _slipsGrid.ColumnHeadersHeight = 34;
            _slipsGrid.RowTemplate.Height = 31;
            _slipsGrid.EnableHeadersVisualStyles = false;
            _slipsGrid.ColumnHeadersDefaultCellStyle.BackColor =
                Color.FromArgb(124, 33, 29);
            _slipsGrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _slipsGrid.ColumnHeadersDefaultCellStyle.Font =
                new Font("Segoe UI", 9.0f, FontStyle.Bold);
            _slipsGrid.DefaultCellStyle.SelectionBackColor =
                Color.FromArgb(242, 224, 220);
            _slipsGrid.DefaultCellStyle.SelectionForeColor =
                Color.FromArgb(37, 35, 31);
            _slipsGrid.Columns.Add(CreateColumn("Slip", "Slip #", 70));
            _slipsGrid.Columns.Add(CreateColumn("Date", "Entry date", 95));
            _slipsGrid.Columns.Add(CreateColumn("Name", "Customer", 180));
            _slipsGrid.Columns.Add(CreateColumn("Phone", "Phone", 125));
            _slipsGrid.CellClick += GridCellClick;
            _slipsGrid.CellDoubleClick += GridCellDoubleClick;
            layout.Controls.Add(_slipsGrid, 0, 1);

            FlowLayoutPanel actions = new FlowLayoutPanel();
            actions.Dock = DockStyle.Fill;
            actions.WrapContents = false;
            actions.Padding = new Padding(0, 6, 0, 0);
            actions.Controls.Add(CreateButton(
                "Print selected", PrintSelected, false, 130, 34));
            actions.Controls.Add(CreateButton(
                "Delete selected", DeleteSelected, false, 140, 34));
            actions.Controls.Add(CreateButton(
                "Import old workbook", ImportOldWorkbook, false, 160, 34));
            actions.Controls.Add(CreateButton(
                "Open data folder", OpenDataFolder, false, 135, 34));
            layout.Controls.Add(actions, 0, 2);

            _recordCount = new Label();
            _recordCount.Dock = DockStyle.Fill;
            _recordCount.ForeColor = Color.FromArgb(112, 107, 97);
            _recordCount.TextAlign = ContentAlignment.MiddleLeft;
            layout.Controls.Add(_recordCount, 0, 3);
        }

        private static Control CreateHeader(string title, string subtitle)
        {
            TableLayoutPanel header = new TableLayoutPanel();
            header.Dock = DockStyle.Fill;
            header.ColumnCount = 1;
            header.RowCount = 2;
            header.RowStyles.Add(new RowStyle(SizeType.Percent, 62));
            header.RowStyles.Add(new RowStyle(SizeType.Percent, 38));

            Label name = new Label();
            name.Text = title;
            name.Dock = DockStyle.Fill;
            name.Font = new Font("Segoe UI", 19.0f, FontStyle.Bold);
            name.ForeColor = Color.FromArgb(124, 33, 29);
            name.TextAlign = ContentAlignment.MiddleLeft;
            header.Controls.Add(name, 0, 0);

            Label detail = new Label();
            detail.Text = subtitle;
            detail.Dock = DockStyle.Fill;
            detail.Font = new Font("Segoe UI", 9.0f, FontStyle.Regular);
            detail.ForeColor = Color.FromArgb(112, 107, 97);
            detail.TextAlign = ContentAlignment.MiddleLeft;
            header.Controls.Add(detail, 0, 1);
            return header;
        }

        private static void AddField(
            TableLayoutPanel layout, int row, string labelText, Control editor)
        {
            Label label = new Label();
            label.Text = labelText;
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.ForeColor = Color.FromArgb(55, 51, 44);
            layout.Controls.Add(label, 0, row);
            editor.Margin = new Padding(0, 9, 0, 9);
            layout.Controls.Add(editor, 1, row);
        }

        private static DataGridViewTextBoxColumn CreateColumn(
            string name, string header, float fillWeight)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = header,
                FillWeight = fillWeight,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
        }

        private static Button CreateButton(
            string text, EventHandler handler, bool primary, int width, int height)
        {
            Button button = new Button();
            button.Text = text;
            button.Width = width;
            button.Height = height;
            button.Margin = new Padding(0, 0, 8, 0);
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = Color.FromArgb(190, 181, 168);
            button.Font = new Font("Segoe UI", 9.0f, FontStyle.Bold);
            button.Cursor = Cursors.Hand;
            if (primary)
            {
                button.BackColor = Color.FromArgb(124, 33, 29);
                button.ForeColor = Color.White;
                button.FlatAppearance.BorderColor = Color.FromArgb(124, 33, 29);
            }
            else
            {
                button.BackColor = Color.FromArgb(255, 253, 249);
                button.ForeColor = Color.FromArgb(37, 35, 31);
            }
            button.Click += handler;
            return button;
        }

        private void BindDirtyTracking()
        {
            _customerName.TextChanged += InputChanged;
            _phoneNumber.TextChanged += InputChanged;
            _slipNumber.ValueChanged += InputChanged;
            _entryDate.ValueChanged += InputChanged;
        }

        private void InputChanged(object sender, EventArgs eventArgs)
        {
            if (_loadingFields)
                return;
            _dirty = true;
            _entryStatus.Text = "Unsaved changes";
            _entryStatus.ForeColor = Color.FromArgb(121, 83, 20);
        }

        private void SearchChanged(object sender, EventArgs eventArgs)
        {
            RefreshGrid(_loadedRecord == null ? -1 : _loadedRecord.SlipNo);
        }

        private void ClearSearch(object sender, EventArgs eventArgs)
        {
            _searchBox.Clear();
            _searchBox.Focus();
        }

        private void RefreshGrid(int selectedSlip)
        {
            if (_slipsGrid == null)
                return;
            string query = _searchBox == null ? String.Empty : _searchBox.Text.Trim();
            _slipsGrid.Rows.Clear();
            foreach (SlipRecord record in _records.OrderBy(
                delegate(SlipRecord item) { return item.SlipNo; }))
            {
                if (query.Length > 0 &&
                    record.SlipNo.ToString(CultureInfo.InvariantCulture)
                        .IndexOf(query, StringComparison.CurrentCultureIgnoreCase) < 0 &&
                    record.CustomerName.IndexOf(query,
                        StringComparison.CurrentCultureIgnoreCase) < 0 &&
                    record.PhoneNo.IndexOf(query,
                        StringComparison.CurrentCultureIgnoreCase) < 0 &&
                    record.EntryDate.ToString("d", CultureInfo.CurrentCulture)
                        .IndexOf(query, StringComparison.CurrentCultureIgnoreCase) < 0)
                    continue;

                int rowIndex = _slipsGrid.Rows.Add(
                    record.SlipNo.ToString(CultureInfo.InvariantCulture),
                    record.EntryDate.ToString("d", CultureInfo.CurrentCulture),
                    record.CustomerName,
                    record.PhoneNo);
                _slipsGrid.Rows[rowIndex].Tag = record;
                if (record.SlipNo == selectedSlip)
                {
                    _slipsGrid.Rows[rowIndex].Selected = true;
                    _slipsGrid.CurrentCell = _slipsGrid.Rows[rowIndex].Cells[0];
                }
            }
            _recordCount.Text = String.Format(
                CultureInfo.CurrentCulture,
                "{0} saved slips",
                _records.Count);
        }

        private void GridCellClick(object sender, DataGridViewCellEventArgs eventArgs)
        {
            if (eventArgs.RowIndex < 0 ||
                eventArgs.RowIndex >= _slipsGrid.Rows.Count)
                return;
            SlipRecord record = _slipsGrid.Rows[eventArgs.RowIndex].Tag as SlipRecord;
            if (record == null)
                return;
            if (_loadedRecord != null && _loadedRecord.SlipNo == record.SlipNo)
                return;
            if (!ConfirmLeaveCurrentEntry())
            {
                RefreshGrid(_loadedRecord == null ? -1 : _loadedRecord.SlipNo);
                return;
            }
            LoadRecord(record);
        }

        private void GridCellDoubleClick(
            object sender, DataGridViewCellEventArgs eventArgs)
        {
            if (eventArgs.RowIndex >= 0)
                _customerName.Focus();
        }

        private void LoadRecord(SlipRecord record)
        {
            _loadingFields = true;
            try
            {
                _loadedRecord = record.Copy();
                _slipNumber.Value = record.SlipNo;
                _customerName.Text = record.CustomerName;
                _phoneNumber.Text = record.PhoneNo;
                DateTimePicker picker = _entryDate;
                DateTime date = record.EntryDate.Date;
                if (date < picker.MinDate)
                    date = picker.MinDate;
                if (date > picker.MaxDate)
                    date = picker.MaxDate;
                picker.Value = date;
                _dirty = false;
                _entryStatus.Text = "Saved slip";
                _entryStatus.ForeColor = Color.FromArgb(112, 107, 97);
            }
            finally
            {
                _loadingFields = false;
            }
        }

        private void StartNewEntry(bool confirmChanges)
        {
            if (confirmChanges && !ConfirmLeaveCurrentEntry())
                return;
            _loadingFields = true;
            try
            {
                _loadedRecord = null;
                _slipNumber.Value = Math.Min(
                    _slipNumber.Maximum, Math.Max(1, HighestSlip() + 1));
                _customerName.Clear();
                _phoneNumber.Clear();
                _entryDate.Value = DateTime.Today;
                _dirty = false;
                _entryStatus.Text = "New slip ready";
                _entryStatus.ForeColor = Color.FromArgb(112, 107, 97);
                if (_slipsGrid != null)
                    _slipsGrid.ClearSelection();
            }
            finally
            {
                _loadingFields = false;
            }
            _customerName.Focus();
        }

        private void NewSlip(object sender, EventArgs eventArgs)
        {
            StartNewEntry(true);
        }

        private void SaveSlip(object sender, EventArgs eventArgs)
        {
            SlipRecord saved;
            if (SaveCurrent(out saved))
                _entryStatus.Text = "Saved slip #" +
                    saved.SlipNo.ToString(CultureInfo.InvariantCulture);
        }

        private void SaveAndPrint(object sender, EventArgs eventArgs)
        {
            SlipRecord saved;
            if (!SaveCurrent(out saved))
                return;
            try
            {
                PrinterService.PrintSlip(saved, SelectedPrinter());
                _entryStatus.Text = "Slip #" +
                    saved.SlipNo.ToString(CultureInfo.InvariantCulture) +
                    " saved and printed";
                StartNewEntry(false);
            }
            catch (Exception exception)
            {
                MessageBox.Show(this,
                    "Slip #" + saved.SlipNo +
                    " was saved, but it could not be printed.\n\n" +
                    exception.Message +
                    "\n\nSelect the slip and choose Print selected after fixing the printer.",
                    "Slip saved; printer needs attention",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private bool SaveCurrent(out SlipRecord saved)
        {
            saved = null;
            string name = _customerName.Text.Trim();
            string phone = _phoneNumber.Text.Trim();
            int slipNo = Decimal.ToInt32(_slipNumber.Value);
            if (name.Length == 0)
            {
                MessageBox.Show(this, "Enter the customer's name.",
                    "Name required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _customerName.Focus();
                return false;
            }
            if (phone.Length == 0)
            {
                MessageBox.Show(this, "Enter the customer's phone number.",
                    "Phone number required",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _phoneNumber.Focus();
                return false;
            }

            SlipRecord existing = _records.FirstOrDefault(
                delegate(SlipRecord record) { return record.SlipNo == slipNo; });
            if (existing != null &&
                (_loadedRecord == null || _loadedRecord.SlipNo != slipNo))
            {
                MessageBox.Show(this,
                    "Slip #" + slipNo +
                    " already exists. Select it in the saved-slips list to edit it.",
                    "Slip number already used",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            SlipRecord recordToSave = existing == null
                ? new SlipRecord
                {
                    SlipNo = slipNo,
                    CreatedAt = DateTime.Now
                }
                : existing.Copy();
            recordToSave.CustomerName = name;
            recordToSave.PhoneNo = phone;
            recordToSave.EntryDate = _entryDate.Value.Date;

            List<SlipRecord> candidate = CopyRecords(_records);
            int existingIndex = candidate.FindIndex(
                delegate(SlipRecord record) { return record.SlipNo == slipNo; });
            if (existingIndex >= 0)
                candidate[existingIndex] = recordToSave;
            else
                candidate.Add(recordToSave);

            try
            {
                _store.SaveAll(candidate);
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Could not save slip",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            _records = candidate;
            _loadedRecord = recordToSave.Copy();
            _dirty = false;
            saved = recordToSave.Copy();
            _entryStatus.Text = "Saved slip #" +
                slipNo.ToString(CultureInfo.InvariantCulture);
            _entryStatus.ForeColor = Color.FromArgb(53, 91, 50);
            RefreshGrid(slipNo);
            return true;
        }

        private void PrintSelected(object sender, EventArgs eventArgs)
        {
            SlipRecord record = SelectedGridRecord();
            if (record == null)
            {
                MessageBox.Show(this, "Select a saved slip to print.",
                    "No slip selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                PrinterService.PrintSlip(record, SelectedPrinter());
                _entryStatus.Text = "Slip #" +
                    record.SlipNo.ToString(CultureInfo.InvariantCulture) + " printed";
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Could not print slip",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteSelected(object sender, EventArgs eventArgs)
        {
            SlipRecord record = SelectedGridRecord();
            if (record == null)
            {
                MessageBox.Show(this, "Select a saved slip to delete.",
                    "No slip selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (_dirty && !ConfirmLeaveCurrentEntry())
                return;
            DialogResult answer = MessageBox.Show(this,
                "Delete slip #" + record.SlipNo + " for " +
                record.CustomerName + "?\n\nThis cannot be undone.",
                "Delete customer slip",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (answer != DialogResult.Yes)
                return;

            List<SlipRecord> candidate = CopyRecords(_records);
            candidate.RemoveAll(delegate(SlipRecord item)
            {
                return item.SlipNo == record.SlipNo;
            });
            try
            {
                _store.SaveAll(candidate);
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Could not delete slip",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            _records = candidate;
            StartNewEntry(false);
            RefreshGrid(-1);
        }

        private void ImportOldWorkbook(object sender, EventArgs eventArgs)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "Import slips from the previous app";
                dialog.Filter = "Excel workbooks (*.xlsx)|*.xlsx";
                dialog.CheckFileExists = true;
                dialog.Multiselect = false;
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                if (_dirty && !ConfirmLeaveCurrentEntry())
                    return;
                try
                {
                    List<SlipRecord> imported = LegacyWorkbookImporter.Read(dialog.FileName);
                    List<SlipRecord> candidate = CopyRecords(_records);
                    HashSet<int> existingSlips = new HashSet<int>(
                        candidate.Select(delegate(SlipRecord item) { return item.SlipNo; }));
                    int added = 0;
                    int skipped = 0;
                    foreach (SlipRecord record in imported)
                    {
                        if (!existingSlips.Add(record.SlipNo))
                        {
                            skipped++;
                            continue;
                        }
                        candidate.Add(record);
                        added++;
                    }

                    if (added > 0)
                        _store.SaveAll(candidate);
                    _records = candidate;
                    _dirty = false;
                    if (_loadedRecord == null)
                        StartNewEntry(false);
                    RefreshGrid(_loadedRecord == null ? -1 : _loadedRecord.SlipNo);
                    MessageBox.Show(this,
                        added.ToString(CultureInfo.CurrentCulture) +
                        " slips imported. " +
                        skipped.ToString(CultureInfo.CurrentCulture) +
                        " existing slip numbers were skipped.\n\n" +
                        "The selected workbook was not changed.",
                        "Import complete",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception exception)
                {
                    MessageBox.Show(this, exception.Message, "Import failed",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void OpenDataFolder(object sender, EventArgs eventArgs)
        {
            try
            {
                Process.Start("explorer.exe", "\"" + _dataDirectory + "\"");
            }
            catch (Exception exception)
            {
                MessageBox.Show(this,
                    "Could not open the data folder:\n" + exception.Message,
                    "Open data folder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshPrinterList(object sender, EventArgs eventArgs)
        {
            RefreshPrinterList();
        }

        private void RefreshPrinterList()
        {
            if (_printerList == null)
                return;
            string preferred = _printerList.SelectedItem as string;
            if (String.IsNullOrEmpty(preferred))
            {
                try { preferred = _settings.LoadPrinterName(); }
                catch (Exception exception)
                {
                    _printerStatus.Text = exception.Message;
                }
            }

            _printerList.Items.Clear();
            foreach (string printer in PrinterSettings.InstalledPrinters)
                _printerList.Items.Add(printer);

            int selected = -1;
            if (!String.IsNullOrEmpty(preferred))
                selected = _printerList.Items.IndexOf(preferred);
            if (selected < 0 && _printerList.Items.Count > 0)
            {
                string defaultPrinter = new PrinterSettings().PrinterName;
                selected = _printerList.Items.IndexOf(defaultPrinter);
                if (selected < 0)
                    selected = 0;
            }
            if (selected >= 0)
            {
                _printerList.SelectedIndex = selected;
                _printerStatus.Text = "Selected printer: " +
                    _printerList.Items[selected].ToString();
            }
            else
                _printerStatus.Text = "No installed printers were found.";
        }

        private void SavePrinter(object sender, EventArgs eventArgs)
        {
            string printer = SelectedPrinter();
            if (printer.Length == 0)
            {
                MessageBox.Show(this, "Install or connect a printer first.",
                    "No printer available", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            try
            {
                _settings.SavePrinterName(printer);
                _printerStatus.Text = "Saved printer: " + printer;
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Could not save printer",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TestPrint(object sender, EventArgs eventArgs)
        {
            try
            {
                PrinterService.PrintTest(SelectedPrinter());
                _printerStatus.Text =
                    "Test slip sent to " + SelectedPrinter() +
                    ". No customer record was created.";
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.Message, "Test print failed",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string SelectedPrinter()
        {
            return _printerList == null || _printerList.SelectedItem == null
                ? String.Empty
                : _printerList.SelectedItem.ToString();
        }

        private SlipRecord SelectedGridRecord()
        {
            if (_slipsGrid == null || _slipsGrid.SelectedRows.Count == 0)
                return null;
            return _slipsGrid.SelectedRows[0].Tag as SlipRecord;
        }

        private int HighestSlip()
        {
            return _records.Count == 0
                ? 0
                : _records.Max(delegate(SlipRecord record) { return record.SlipNo; });
        }

        private static List<SlipRecord> CopyRecords(List<SlipRecord> records)
        {
            return records.Select(
                delegate(SlipRecord record) { return record.Copy(); }).ToList();
        }

        private bool ConfirmLeaveCurrentEntry()
        {
            if (!_dirty)
                return true;
            DialogResult answer = MessageBox.Show(this,
                "This entry has unsaved changes.\n\n" +
                "Choose Yes to save, No to discard, or Cancel to stay here.",
                "Unsaved changes",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (answer == DialogResult.Yes)
            {
                SlipRecord saved;
                return SaveCurrent(out saved);
            }
            return answer == DialogResult.No;
        }

        private void MainFormClosing(object sender, FormClosingEventArgs eventArgs)
        {
            if (!ConfirmLeaveCurrentEntry())
                eventArgs.Cancel = true;
        }
    }
}
