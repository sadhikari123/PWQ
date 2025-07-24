
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text.Json;
using System.Drawing;
using PWQ.Models;
using CsvHelper;




namespace PWQ.Forms
{
    // ...existing code...
    // ...existing code...
    // Modern Windows 12 color table for menu styling
    public class ModernColorTable : ProfessionalColorTable
    {
        public override Color MenuItemSelected => Color.FromArgb(230, 230, 230);
        public override Color MenuItemSelectedGradientBegin => Color.FromArgb(230, 230, 230);
        public override Color MenuItemSelectedGradientEnd => Color.FromArgb(230, 230, 230);
        public override Color MenuItemPressedGradientBegin => Color.FromArgb(200, 200, 200);
        public override Color MenuItemPressedGradientEnd => Color.FromArgb(200, 200, 200);
        public override Color MenuItemBorder => Color.FromArgb(200, 200, 200);
        public override Color MenuBorder => Color.FromArgb(229, 229, 229);
        public override Color ToolStripDropDownBackground => Color.FromArgb(249, 249, 249);
        public override Color ImageMarginGradientBegin => Color.FromArgb(249, 249, 249);
        public override Color ImageMarginGradientMiddle => Color.FromArgb(249, 249, 249);
        public override Color ImageMarginGradientEnd => Color.FromArgb(249, 249, 249);
    }

    public partial class MainForm : Form
    {
        private Color networkStatusColor = Color.Gray;
        private Dictionary<string, ConfigFile> configFiles = new Dictionary<string, ConfigFile>();
        private TabControl tabControl;
        private MenuStrip menuStrip;
        private ToolStripMenuItem fileMenu, helpMenu;
        private ToolStripMenuItem refreshMenuItem, exitMenuItem, aboutMenuItem, versionMenuItem;
        private Dictionary<string, DataGridView> grids = new Dictionary<string, DataGridView>();
        private DataGridView historyGrid;
        // Removed: unused field
        // Clean Excel-like filter state: configName -> columnName -> allowed values
        private Dictionary<string, Dictionary<string, HashSet<string>>> activeFilters = new Dictionary<string, Dictionary<string, HashSet<string>>>();
        private System.Data.DataTable ToDataTable(List<Dictionary<string, string>> rows, List<string> columns)
        {
            var dt = new System.Data.DataTable();
            foreach (var col in columns)
                dt.Columns.Add(col);
            foreach (var row in rows)
            {
                var dr = dt.NewRow();
                foreach (var col in columns)
                    dr[col] = row.ContainsKey(col) ? row[col] : "";
                dt.Rows.Add(dr);
            }
            return dt;
        }

        // File locking mechanism
        private Dictionary<string, FileStream> fileLocks = new Dictionary<string, FileStream>();
        private Dictionary<string, string> lockOwners = new Dictionary<string, string>();

        public MainForm()
        {
            InitializeComponent();
            SetIcon();
            ApplyModernTheme();
            BuildMenu();
            LoadConfigFiles();
            BuildTabs();
            UpdateNetworkStatus();
            // Optionally, check periodically:
            var timer = new Timer();
            timer.Interval = 10000; // 10 seconds
            timer.Tick += (s, e) => UpdateNetworkStatus();
            timer.Start();
            // Handle form closing to release all file locks
            this.FormClosing += MainForm_FormClosing;

            // (networkStatusPanel logic removed: panel not defined in this context)

        }

        // Utility: Check if user has write access to a file
        private bool HasWriteAccess(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    return false;
                string tempFile = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileName(filePath) + ".writecheck");
                using (FileStream fs = File.Create(tempFile, 1, FileOptions.DeleteOnClose))
                {
                    // If we can create and delete a file, we have write access
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void UpdateNetworkStatus()
        {
            // Check if the network file for Stale EF Config is accessible
            string path = PWQ.Models.Constants.CsvFiles["PWQ Litho Layers"];
            bool connected = System.IO.File.Exists(path);
            networkStatusColor = connected ? Color.LimeGreen : Color.Gray;
            // (networkStatusPanel and networkStatusToolTip logic removed: not defined in this context)
        }

        // Draw a filled circle for the indicator
        private void networkStatusPanel_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            // (networkStatusPanel logic removed: not defined in this context)
            int diameter = 20;
            using (var brush = new SolidBrush(networkStatusColor))
            {
                g.FillEllipse(brush, 2, 2, diameter, diameter);
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            ReleaseAllLocks();
        }

        private void ApplyModernTheme()
        {
            // Set form properties for modern Windows 12 look
            this.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.BackColor = Color.FromArgb(243, 243, 243); // Windows 12 background
            this.ForeColor = Color.FromArgb(32, 32, 32); // Windows 12 text color
            this.Text = "PWQ";
            this.Size = new Size(1400, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1200, 700);
        }

        private void SetIcon()
        {
            try
            {
                // Load the icon from the file (prefer .ico, fallback to .png)
                string iconPath = Path.Combine(Application.StartupPath, "app_icon.ico");
                if (File.Exists(iconPath))
                {
                    this.Icon = new Icon(iconPath);
                }
                else
                {
                    // Try relative path
                    iconPath = "app_icon.ico";
                    if (File.Exists(iconPath))
                    {
                        this.Icon = new Icon(iconPath);
                    }
                    else
                    {
                        // Fallback: try PNG (Windows Forms supports this via Bitmap)
                        string pngPath = Path.Combine(Application.StartupPath, "app_icon_256.png");
                        if (File.Exists(pngPath))
                        {
                            this.Icon = System.Drawing.Icon.FromHandle(new Bitmap(pngPath).GetHicon());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // If icon loading fails, just continue without icon
                System.Diagnostics.Debug.WriteLine($"Could not load icon: {ex.Message}");
            }
        }

        private void BuildMenu()
        {
            menuStrip = new MenuStrip
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                BackColor = Color.FromArgb(249, 249, 249),
                ForeColor = Color.FromArgb(32, 32, 32),
                Renderer = new ToolStripProfessionalRenderer(new ModernColorTable())
            };
            
            fileMenu = new ToolStripMenuItem("File")
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Regular)
            };
            refreshMenuItem = new ToolStripMenuItem("Refresh All", null, (s, e) => RefreshAll())
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Regular)
            };
            exitMenuItem = new ToolStripMenuItem("Exit", null, (s, e) => Close())
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Regular)
            };
            versionMenuItem = new ToolStripMenuItem("Version", null, (s, e) => ShowVersion())
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Regular)
            };
            fileMenu.DropDownItems.AddRange(new ToolStripItem[] { refreshMenuItem, versionMenuItem, new ToolStripSeparator(), exitMenuItem });

            helpMenu = new ToolStripMenuItem("Help")
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Regular)
            };
            // Add placeholder for Wiki link
            var wikiPlaceholder = new ToolStripMenuItem("A link to wiki on PWQ will be added later")
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                Enabled = false
            };
            aboutMenuItem = new ToolStripMenuItem("About", null, (s, e) => ShowAbout())
            {
                Font = new Font("Segoe UI", 10F, FontStyle.Regular)
            };
            helpMenu.DropDownItems.Add(wikiPlaceholder);
            helpMenu.DropDownItems.Add(aboutMenuItem);

            menuStrip.Items.AddRange(new ToolStripMenuItem[] { fileMenu, helpMenu });
            Controls.Add(menuStrip);
            MainMenuStrip = menuStrip;
        }

        private void ShowVersion()
        {
            // Read and parse VERSION_HISTORY.md
            // Try both base directory and current directory
            string[] possiblePaths = {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "VERSION_HISTORY.md"),
                Path.Combine(Directory.GetCurrentDirectory(), "VERSION_HISTORY.md")
            };
            string versionFile = possiblePaths.FirstOrDefault(File.Exists);
            if (versionFile == null)
            {
                MessageBox.Show("VERSION_HISTORY.md not found.\n\nIf you are not connected to the Intel VPN or network, you may not have access to the version history file. Please connect to VPN. All users should have at least read access.", "Version History - Network Issue", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var lines = File.ReadAllLines(versionFile);
            var entries = new List<(string Version, string Date, string Changes)>();
            var regex = new System.Text.RegularExpressions.Regex(@"\|\s*(\d+\.\d+\.\d+)\s*\|\s*(\d{4}-\d{2}-\d{2})\s*\|\s*(.*?)\s*\|");
            foreach (var line in lines)
            {
                var match = regex.Match(line);
                if (match.Success)
                {
                    entries.Add((match.Groups[1].Value, match.Groups[2].Value, match.Groups[3].Value));
                }
            }
            if (entries.Count == 0)
            {
                MessageBox.Show("No version entries found in VERSION_HISTORY.md.", "Version History", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Create a form to show the table
            var dlg = new Form
            {
                Text = "PWQ Version History",
                Size = new Size(900, 400),
                StartPosition = FormStartPosition.CenterParent,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                MinimizeBox = false,
                MaximizeBox = false,
                FormBorderStyle = FormBorderStyle.FixedDialog
            };
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToResizeColumns = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                ColumnHeadersDefaultCellStyle = { Font = new Font("Segoe UI", 10F, FontStyle.Bold) },
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None
            };
            grid.Columns.Add("Version", "Version");
            grid.Columns.Add("Date", "Date");
            grid.Columns.Add("Changes", "Changes");
            foreach (var entry in entries)
            {
                grid.Rows.Add(entry.Version, entry.Date, entry.Changes);
            }
            grid.Columns[0].Width = 90;
            grid.Columns[1].Width = 110;
            grid.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            grid.Columns[2].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            dlg.Controls.Add(grid);
            var btnClose = new Button
            {
                Text = "Close",
                DialogResult = DialogResult.OK,
                Dock = DockStyle.Bottom,
                Height = 40,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnClose.FlatAppearance.BorderSize = 0;
            dlg.AcceptButton = btnClose;
            dlg.CancelButton = btnClose;
            dlg.Controls.Add(btnClose);
            dlg.ShowDialog(this);
        }

        private void LoadConfigFiles()
        {
            configFiles.Clear();
            System.Diagnostics.Debug.WriteLine("Loading config files...");
            foreach (var kvp in Constants.CsvFiles)
            {
                System.Diagnostics.Debug.WriteLine($"Checking file: {kvp.Key} at {kvp.Value}");
                var config = new ConfigFile { Name = kvp.Key, Path = kvp.Value };
                // bool fileLoaded = false;
                
                try
                {
                    if (File.Exists(kvp.Value))
                    {
                        config.LoadFromFile();
                        System.Diagnostics.Debug.WriteLine($"✓ Loaded {kvp.Key}: {config.Rows.Count} rows");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"✗ File not found: {kvp.Value}");
                        config.CreateEmptyStructure();
                        System.Diagnostics.Debug.WriteLine($"Created empty structure for {kvp.Key}");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"✗ Error loading {kvp.Key}: {ex.Message}");
                    config.CreateEmptyStructure();
                }
                
                configFiles[kvp.Key] = config;
            }
        }

        private void BuildTabs()
        {
            // Remove existing tab control if it exists
            if (tabControl != null && Controls.Contains(tabControl))
            {
                Controls.Remove(tabControl);
            }

            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                Padding = new Point(12, 8),
                ItemSize = new Size(120, 40),
                SizeMode = TabSizeMode.Fixed,
                Appearance = TabAppearance.Normal,
                HotTrack = true
            };
            grids.Clear();

            // Custom tab order: LHD LayerList, Stale EF Config, Reticle Comments, ... EFCC_Ignore_Lines, XQual_Ignore_Lines, right before History
            var orderedConfigs = configFiles.Values.ToList();
            orderedConfigs = orderedConfigs.OrderBy(c =>
                c.Name == "LHD LayerList" ? 0 :
                c.Name == "Stale EF Config" ? 1 :
                c.Name == "NPI Lot Config" ? 2 :
                c.Name == "NPI Run Plan" ? 3 :
                c.Name == "OPC Config" ? 4 :
                c.Name == "Reticle Comments" ? 5 :
                c.Name == "EFCC_Ignore_Lines" ? 98 :
                c.Name == "XQual_Ignore_Lines" ? 99 :
                6).ToList();


            foreach (var config in orderedConfigs)
            {
                bool isMAO = config.Name.ToUpper().Contains("MAO");
                bool canWrite = HasWriteAccess(config.Path);
                var tab = new TabPage(config.Name);
                var tableLayout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = 3,
                    Padding = new Padding(10),
                    BackColor = System.Drawing.Color.White
                };
                tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30)); // Notification row
                tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // DataGridView
                tableLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Button row

                // Notification panel
                var notificationPanel = new Panel
                {
                    Dock = DockStyle.Top,
                    Height = 40,
                    BackColor = Color.FromArgb(255, 249, 196),
                    Margin = new Padding(0, 0, 0, 8)
                };
                var notificationLabel = new Label
                {
                    Text = File.Exists(config.Path)
                        ? (canWrite
                            ? "✅ You have write access to this config. Editing is enabled."
                            : "⚠️ You have read-only access to this config. Editing is disabled.")
                        : "❌ Cannot access config file. Please connect to the Intel VPN or network. All users should have at least read access.",
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                    ForeColor = File.Exists(config.Path)
                        ? (canWrite ? Color.FromArgb(16, 137, 62) : Color.FromArgb(102, 77, 3))
                        : Color.FromArgb(196, 43, 28),
                    Padding = new Padding(16, 0, 16, 0),
                    Visible = true
                };
                notificationPanel.Controls.Add(notificationLabel);
                tableLayout.Controls.Add(notificationPanel, 0, 0);

                // DataGridView
                var dataGrid = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    ReadOnly = true,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    MultiSelect = false,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                    BackgroundColor = Color.White,
                    Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                    RowTemplate = { Height = 36 },
                    ColumnHeadersHeight = 44,
                    ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                    EnableHeadersVisualStyles = false,
                    GridColor = Color.FromArgb(200, 200, 200),
                    BorderStyle = BorderStyle.None,
                    CellBorderStyle = DataGridViewCellBorderStyle.Single,
                    RowHeadersVisible = false,
                    AllowUserToResizeRows = false,
                    ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
                    AutoGenerateColumns = true
                };
                dataGrid.DataSource = ToDataTable(config.Rows, config.Columns);
                dataGrid.DataBindingComplete += (s, e) =>
                {
                    dataGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                    dataGrid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(249, 249, 249);
                    dataGrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(32, 32, 32);
                    dataGrid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
                    dataGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    dataGrid.ColumnHeadersDefaultCellStyle.Padding = new Padding(16, 0, 16, 0);
                    dataGrid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 212);
                    dataGrid.DefaultCellStyle.SelectionForeColor = Color.White;
                    dataGrid.DefaultCellStyle.Padding = new Padding(16, 0, 16, 0);
                    dataGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 249, 249);

                    // Set column header text with dark bold triangle for filter dropdown (same as History tab)
                    foreach (DataGridViewColumn col in dataGrid.Columns)
                    {
                        string baseHeader = col.HeaderText;
                        // Remove any existing triangle
                        baseHeader = baseHeader.Replace(" ▼", "").Replace("▼", "").TrimEnd();
                        col.HeaderText = baseHeader + "  ▼"; // Add dark bold triangle
                        col.Name = baseHeader; // Ensure column.Name is always the original column name (no triangle)
                    }
                };
                // Enable double-click to edit row (disable for MAO if no write)
                if (!(isMAO && !canWrite))
                {
                    dataGrid.CellDoubleClick += (s, e) =>
                    {
                        if (e.RowIndex >= 0)
                        {
                            // Get the unique key for the selected row (use all column values as composite key)
                            var grid = (DataGridView)s;
                            var row = grid.Rows[e.RowIndex];
                            var keyDict = new Dictionary<string, string>();
                            foreach (DataGridViewColumn col in grid.Columns)
                            {
                                keyDict[col.Name] = row.Cells[col.Index].Value?.ToString() ?? "";
                            }
                            EditRow(config.Name, keyDict);
                        }
                    };
                }
                // Excel-style column header click for filtering (all tabs, using exact same logic as History tab)
                dataGrid.ColumnHeaderMouseClick += (s, e) =>
                {
                    ShowFilterDropdown(config.Name, e.ColumnIndex, dataGrid, e.Location);
                };
                grids[config.Name] = dataGrid;
                tableLayout.Controls.Add(dataGrid, 0, 1);

                // Button panel
                var btnPanel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = FlowDirection.LeftToRight,
                    AutoSize = true,
                    Padding = new Padding(0, 16, 0, 16),
                    BackColor = Color.White
                };
                var btnAdd = new Button
                {
                    Text = "Add Row",
                    Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    Padding = new Padding(20, 8, 20, 8),
                    Margin = new Padding(0, 0, 12, 0),
                    FlatStyle = FlatStyle.Flat,
                    MinimumSize = new Size(120, 40),
                    UseVisualStyleBackColor = false,
                    Enabled = !(isMAO && !canWrite)
                };
                btnAdd.FlatAppearance.BorderSize = 0;
                if (!(isMAO && !canWrite))
                    btnAdd.Click += (s, e) => AddRow(config.Name);
                var btnEdit = new Button
                {
                    Text = "Edit Row",
                    Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                    BackColor = Color.FromArgb(16, 137, 62),
                    ForeColor = Color.White,
                    Padding = new Padding(20, 8, 20, 8),
                    Margin = new Padding(0, 0, 12, 0),
                    FlatStyle = FlatStyle.Flat,
                    MinimumSize = new Size(120, 40),
                    UseVisualStyleBackColor = false,
                    Enabled = !(isMAO && !canWrite)
                };
                btnEdit.FlatAppearance.BorderSize = 0;
                if (!(isMAO && !canWrite))
                    btnEdit.Click += (s, e) => EditRow(config.Name);
                var btnDelete = new Button
                {
                    Text = "Delete Row",
                    Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                    BackColor = Color.FromArgb(196, 43, 28),
                    ForeColor = Color.White,
                    Padding = new Padding(20, 8, 20, 8),
                    Margin = new Padding(0, 0, 12, 0),
                    FlatStyle = FlatStyle.Flat,
                    MinimumSize = new Size(120, 40),
                    UseVisualStyleBackColor = false,
                    Enabled = !(isMAO && !canWrite)
                };
                btnDelete.FlatAppearance.BorderSize = 0;
                if (!(isMAO && !canWrite))
                    btnDelete.Click += (s, e) => DeleteRow(config.Name);
                btnPanel.Controls.Add(btnAdd);
                btnPanel.Controls.Add(btnEdit);
                btnPanel.Controls.Add(btnDelete);
                tableLayout.Controls.Add(btnPanel, 0, 2);

                tab.Controls.Add(tableLayout);
                tabControl.TabPages.Add(tab);
            }
            // End foreach
            // After all config tabs, add the History tab and bring to front
            BuildHistoryTab();
            Controls.Add(tabControl);
            tabControl.BringToFront();
        }
    // Excel-style filter dropdown for config tabs (not History)
    // Legacy: This method is now replaced by ShowFilterDropdown for all tabs for consistency.
    // Kept for reference or if needed for legacy support.
    // (Intentionally left empty)
    private void ShowConfigFilterDropdown(string configName, int columnIndex, DataGridView grid, Point location)
    {
    }
// ---- BEGIN: Removed legacy filter dropdown code for config tabs ----
// (All Excel-style filter dropdown logic for config tabs is now handled by ShowFilterDropdown for consistency.)
// ---- END ----

    // (Removed: legacy filter method, now replaced by SetConfigFilter everywhere)

        // Clean filter application: always applies all active filters for the config
        private void ApplyConfigFilters(string configName)
        {
            if (!configFiles.ContainsKey(configName) || !grids.ContainsKey(configName))
                return;
            var config = configFiles[configName];
            var grid = grids[configName];
            var filteredRows = config.Rows;
            if (activeFilters.ContainsKey(configName))
            {
                foreach (var filter in activeFilters[configName])
                {
                    var col = filter.Key;
                    var allowed = filter.Value;
                    filteredRows = filteredRows.Where(row => row.ContainsKey(col) && allowed.Contains(row[col])).ToList();
                }
            }
            var dt = ToDataTable(filteredRows, config.Columns);
            grid.DataSource = dt;
            grid.Refresh();
            if (dt.Rows.Count > 0)
            {
                grid.ClearSelection();
                grid.Rows[0].Selected = true;
            }
        }
        // Clean filter setter: sets allowed values for a column, or removes filter if all values selected
        private void SetConfigFilter(string configName, string columnName, List<string> allowedValues, List<string> allPossibleValues)
        {
            if (!activeFilters.ContainsKey(configName))
                activeFilters[configName] = new Dictionary<string, HashSet<string>>();
            // If all values are selected, remove filter for this column
            if (allowedValues.Count == 0 || allowedValues.Count == allPossibleValues.Count)
            {
                activeFilters[configName].Remove(columnName);
            }
            else
            {
                activeFilters[configName][columnName] = new HashSet<string>(allowedValues);
            }
            ApplyConfigFilters(configName);
        }

        // Build the History tab with Refresh and View Details
        private void BuildHistoryTab()
        {
            var historyTab = new TabPage("History");
            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10),
                BackColor = System.Drawing.Color.White
            };
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30)); // Info row
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // DataGridView
            tableLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Button row

            // Modern info panel for History tab
            var infoPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = Color.FromArgb(232, 245, 233) // Light green background
            };
            var infoLabel = new Label
            {
                Text = "📝 Complete history of all row additions, edits, and deletions",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                ForeColor = Color.FromArgb(30, 70, 32), // Dark green text
                Padding = new Padding(16, 0, 16, 0)
            };
            infoPanel.Controls.Add(infoLabel);
            tableLayout.Controls.Add(infoPanel, 0, 0);

            // History DataGridView with modern styling and Excel-style filtering
            historyGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None, // Manual control for better sizing
                BackgroundColor = Color.White,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                RowTemplate = { Height = 36 },
                ColumnHeadersHeight = 44,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                EnableHeadersVisualStyles = false,
                GridColor = Color.FromArgb(200, 200, 200),
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.Single,
                RowHeadersVisible = false,
                AllowUserToResizeRows = false,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            };

            // Excel-style column header click for filtering
            historyGrid.ColumnHeaderMouseClick += (s, e) =>
            {
                ShowFilterDropdown("History", e.ColumnIndex, historyGrid, e.Location);
            };

            // DataBindingComplete event handler for consistent styling and filter icons
            historyGrid.DataBindingComplete += (s, e) =>
            {
                // Modern Windows 12 column header styling for History
                historyGrid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(249, 249, 249);
                historyGrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(32, 32, 32);
                historyGrid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
                historyGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                historyGrid.ColumnHeadersDefaultCellStyle.Padding = new Padding(16, 0, 0, 0);

                // Modern selection styling
                historyGrid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 212);
                historyGrid.DefaultCellStyle.SelectionForeColor = Color.White;
                historyGrid.DefaultCellStyle.Padding = new Padding(16, 0, 0, 0);

                // Modern alternating row colors
                historyGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 249, 249);

                // Set column widths and add filter icons
                if (historyGrid.Columns.Count > 0)
                {
                    if (historyGrid.Columns["Timestamp"] != null)
                    {
                        historyGrid.Columns["Timestamp"].Width = 150;
                        if (!historyGrid.Columns["Timestamp"].HeaderText.EndsWith(" ▼"))
                            historyGrid.Columns["Timestamp"].HeaderText = "Timestamp ▼";
                    }
                    if (historyGrid.Columns["UserID"] != null)
                    {
                        historyGrid.Columns["UserID"].Width = 100;
                        if (!historyGrid.Columns["UserID"].HeaderText.EndsWith(" ▼"))
                            historyGrid.Columns["UserID"].HeaderText = "User ▼";
                    }
                    if (historyGrid.Columns["ConfigFile"] != null)
                    {
                        historyGrid.Columns["ConfigFile"].Width = 120;
                        if (!historyGrid.Columns["ConfigFile"].HeaderText.EndsWith(" ▼"))
                            historyGrid.Columns["ConfigFile"].HeaderText = "Config File ▼";
                    }
                    if (historyGrid.Columns["Operation"] != null)
                    {
                        historyGrid.Columns["Operation"].Width = 80;
                        if (!historyGrid.Columns["Operation"].HeaderText.EndsWith(" ▼"))
                            historyGrid.Columns["Operation"].HeaderText = "Operation ▼";
                    }
                    if (historyGrid.Columns["RowKey"] != null)
                    {
                        historyGrid.Columns["RowKey"].Width = 80;
                        if (!historyGrid.Columns["RowKey"].HeaderText.EndsWith(" ▼"))
                            historyGrid.Columns["RowKey"].HeaderText = "Row Key ▼";
                    }
                    if (historyGrid.Columns["ChangeSummary"] != null)
                    {
                        historyGrid.Columns["ChangeSummary"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        if (!historyGrid.Columns["ChangeSummary"].HeaderText.EndsWith(" ▼"))
                            historyGrid.Columns["ChangeSummary"].HeaderText = "Change Summary ▼";
                    }
                }

                // Color code by operation
                foreach (DataGridViewRow row in historyGrid.Rows)
                {
                    if (row?.Cells != null && row.Cells.Count > 0 && row.Cells["Operation"]?.Value != null)
                    {
                        string operation = row.Cells["Operation"].Value.ToString();
                        switch (operation)
                        {
                            case "ADD":
                                row.DefaultCellStyle.BackColor = Color.FromArgb(232, 245, 233); // Modern light green
                                break;
                            case "EDIT":
                                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 244, 229); // Modern light orange
                                break;
                            case "DELETE":
                                row.DefaultCellStyle.BackColor = Color.FromArgb(253, 237, 237); // Modern light red
                                break;
                        }
                    }
                }
            };

            // Double-click to view details
            historyGrid.CellDoubleClick += (s, e) =>
            {
                if (e.RowIndex >= 0)
                    ShowHistoryDetails();
            };

            tableLayout.Controls.Add(historyGrid, 0, 1);

            // Modern button panel for History tab
            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Padding = new Padding(0, 16, 0, 16),
                BackColor = Color.White
            };
            var btnRefresh = new Button
            {
                Text = "Refresh History",
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                Padding = new Padding(20, 8, 20, 8),
                Margin = new Padding(0, 0, 12, 0),
                FlatStyle = FlatStyle.Flat,
                MinimumSize = new Size(140, 40),
                UseVisualStyleBackColor = false
            };
            btnRefresh.FlatAppearance.BorderSize = 0;
            btnRefresh.Click += (s, e) => RefreshHistoryGrid();

            var btnViewDetails = new Button
            {
                Text = "View Details",
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                BackColor = Color.FromArgb(16, 137, 62),
                ForeColor = Color.White,
                Padding = new Padding(20, 8, 20, 8),
                Margin = new Padding(0, 0, 12, 0),
                FlatStyle = FlatStyle.Flat,
                MinimumSize = new Size(140, 40),
                UseVisualStyleBackColor = false
            };
            btnViewDetails.FlatAppearance.BorderSize = 0;
            btnViewDetails.Click += (s, e) => ShowHistoryDetails();

            btnPanel.Controls.Add(btnRefresh);
            btnPanel.Controls.Add(btnViewDetails);
            tableLayout.Controls.Add(btnPanel, 0, 2);

            historyTab.Controls.Add(tableLayout);
            tabControl.TabPages.Add(historyTab);
            // Removed: unused field assignment
            RefreshHistoryGrid();
        }

        // Refresh the history grid data
        private void RefreshHistoryGrid()
        {
            var historyEntries = HistoryManager.LoadHistoryEntries();
            // Sort by Timestamp descending (latest first)
            historyEntries = historyEntries
                .OrderByDescending(h => {
                    DateTime dtValue;
                    if (DateTime.TryParse(h.Timestamp, out dtValue))
                        return dtValue;
                    return DateTime.MinValue;
                })
                .ToList();
            // Apply Excel-style filters if any
            if (activeFilters.ContainsKey("History"))
            {
                foreach (var filter in activeFilters["History"])
                {
                    var columnName = filter.Key;
                    var allowedValues = filter.Value;
                    historyEntries = historyEntries.Where(entry =>
                    {
                        var value = GetHistoryColumnValue(entry, columnName);
                        return allowedValues.Contains(value);
                    }).ToList();
                }
            }
            var dt = new System.Data.DataTable();
            dt.Columns.Add("Timestamp", typeof(string));
            dt.Columns.Add("UserID", typeof(string));
            dt.Columns.Add("ConfigFile", typeof(string));
            dt.Columns.Add("Operation", typeof(string));
            dt.Columns.Add("RowKey", typeof(string));
            dt.Columns.Add("ChangeSummary", typeof(string));
            foreach (var h in historyEntries)
            {
                string formattedTimestamp = h.Timestamp;
                DateTime dtValue;
                if (DateTime.TryParse(h.Timestamp, out dtValue))
                {
                    formattedTimestamp = dtValue.ToString("MM/dd/yyyy HH:mm:ss");
                }

                // Generate change summary
                string changeSummary = "";
                var oldDict = new Dictionary<string, string>();
                var newDict = new Dictionary<string, string>();
                try { if (!string.IsNullOrWhiteSpace(h.OldValues)) oldDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(h.OldValues); } catch { }
                try { if (!string.IsNullOrWhiteSpace(h.NewValues)) newDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(h.NewValues); } catch { }
                var allKeys = oldDict.Keys.Union(newDict.Keys).OrderBy(k => k).ToList();
                var changes = new List<string>();
                foreach (var key in allKeys)
                {
                    var oldVal = oldDict.ContainsKey(key) ? oldDict[key] : "";
                    var newVal = newDict.ContainsKey(key) ? newDict[key] : "";
                    if (oldVal != newVal)
                    {
                        changes.Add($"{key}: '{oldVal}' → '{newVal}'");
                    }
                }
                if (changes.Count > 0)
                    changeSummary = string.Join(", ", changes);
                else
                    changeSummary = h.ChangeSummary ?? "";

                dt.Rows.Add(formattedTimestamp, h.UserID, h.ConfigFile, h.Operation, h.RowKey, changeSummary);
            }
            historyGrid.DataSource = dt;
        }

        // Show details dialog for selected history entry
        private void ShowHistoryDetails()
        {
            if (historyGrid.SelectedRows.Count == 0)
                return;
            var row = historyGrid.SelectedRows[0];
            // Extract all available cell values from the selected row

            var details = new Dictionary<string, string>();
            foreach (DataGridViewCell cell in row.Cells)
            {
                var colName = cell.OwningColumn?.HeaderText ?? cell.OwningColumn?.Name ?? "";
                details[colName] = cell.Value?.ToString() ?? "";
            }

            // Helper to get value by alternate keys
            string GetDetail(string key, string altKey = null)
            {
                if (details.ContainsKey(key)) return details[key];
                if (altKey != null && details.ContainsKey(altKey)) return details[altKey];
                return "";
            }

            // Try to load the full HistoryEntry from the CSV for OldValues/NewValues
            var historyEntries = PWQ.Models.HistoryManager.LoadHistoryEntries();
            PWQ.Models.HistoryEntry entry = null;
            foreach (var h in historyEntries)
            {
                if ((h.Timestamp ?? "") == GetDetail("Timestamp") &&
                    (h.UserID ?? "") == GetDetail("UserID", "User") &&
                    (h.ConfigFile ?? "") == GetDetail("ConfigFile", "Config File") &&
                    (h.Operation ?? "") == GetDetail("Operation") &&
                    (h.RowKey ?? "") == GetDetail("RowKey", "Row Key"))
                {
                    entry = h;
                    break;
                }
            }
            var oldDict = new Dictionary<string, string>();
            var newDict = new Dictionary<string, string>();
            if (entry != null)
            {
                try { if (!string.IsNullOrWhiteSpace(entry.OldValues)) oldDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(entry.OldValues); } catch { }
                try { if (!string.IsNullOrWhiteSpace(entry.NewValues)) newDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(entry.NewValues); } catch { }
            }
            // Use all keys found in either OldValues or NewValues
            var allKeys = oldDict.Keys.Union(newDict.Keys).OrderBy(k => k).ToList();
            if (allKeys.Count == 0)
            {
                // Fallback: show all available details from the grid row
                allKeys = details.Keys.Where(k => k != "Timestamp" && k != "User" && k != "Config File" && k != "Operation" && k != "Row Key" && k != "Change Summary").ToList();
            }

            var form = new Form {
                Text = $"History Details - {GetDetail("Operation")}",
                Size = new System.Drawing.Size(900, 340),
                StartPosition = FormStartPosition.CenterParent,
                Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular)
            };
            var table = new DataGridView {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular),
                BackgroundColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.None
            };
            table.Columns.Add("Field", "Field");
            foreach (var col in allKeys) table.Columns.Add(col, col);
            table.Rows.Add();
            table.Rows.Add();
            table.Rows[0].Cells[0].Value = "Old Values";
            table.Rows[1].Cells[0].Value = "New Values";
            bool anyChange = false;
            for (int i = 0; i < allKeys.Count; i++) {
                var key = allKeys[i];
                var oldVal = oldDict.ContainsKey(key) ? oldDict[key] : "";
                var newVal = newDict.ContainsKey(key) ? newDict[key] : "";
                // If no Old/NewValues, fallback to grid row
                if (string.IsNullOrEmpty(oldVal) && details.ContainsKey(key)) oldVal = details[key];
                if (string.IsNullOrEmpty(newVal) && details.ContainsKey(key)) newVal = details[key];
                if (oldVal != newVal) {
                    table.Rows[0].Cells[i+1].Value = oldVal;
                    table.Rows[1].Cells[i+1].Value = newVal;
                    table.Rows[1].Cells[i+1].Style.ForeColor = System.Drawing.Color.Red;
                    anyChange = true;
                } else {
                    table.Rows[0].Cells[i+1].Value = "";
                    table.Rows[1].Cells[i+1].Value = "";
                }
            }
            // If no changes, show all available values for both rows
            if (!anyChange) {
                for (int i = 0; i < allKeys.Count; i++) {
                    var key = allKeys[i];
                    var oldVal = oldDict.ContainsKey(key) ? oldDict[key] : (details.ContainsKey(key) ? details[key] : "");
                    var newVal = newDict.ContainsKey(key) ? newDict[key] : (details.ContainsKey(key) ? details[key] : "");
                    table.Rows[0].Cells[i+1].Value = oldVal;
                    table.Rows[1].Cells[i+1].Value = newVal;
                }
            }
            var panel = new Panel { Dock = DockStyle.Top, Height = 110 };
            var meta = new Label {
                Text = $"Timestamp: {GetDetail("Timestamp")}    User: {GetDetail("UserID", "User")}    Config File: {GetDetail("ConfigFile", "Config File")}    Operation: {GetDetail("Operation")}    Row Key: {GetDetail("RowKey", "Row Key")}\nSummary: {GetDetail("ChangeSummary", "Change Summary")}",
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
                AutoSize = false,
                Height = 110
            };
            panel.Controls.Add(meta);
            form.Controls.Add(table);
            form.Controls.Add(panel);
            form.ShowDialog();
        }
        // Helper: Extract value from raw JSON or key-value string if deserialization fails
        private string TryExtractValueFromRaw(string raw, string key)
        {
            if (string.IsNullOrWhiteSpace(raw) || string.IsNullOrWhiteSpace(key))
                return "";
            // Try to find 'key':'value' or "key":"value" in the raw string
            var patterns = new[] {
                $"\\\"{key}\\\"\\s*:\\s*\\\"(.*?)\\\"", // JSON style
                $"'{key}'\\s*:\\s*'(.*?)'", // single-quoted style
                $"{key}\\s*:\\s*'(.*?)'", // key: 'value'
                $"{key}\\s*:\\s*\\\"(.*?)\\\"" // key: "value"
            };
            foreach (var pat in patterns)
            {
                var match = System.Text.RegularExpressions.Regex.Match(raw, pat);
                if (match.Success && match.Groups.Count > 1)
                    return match.Groups[1].Value;
            }
            return "";
        }

        // Helper for Excel-style filtering (History tab only)
        private string GetHistoryColumnValue(PWQ.Models.HistoryEntry entry, string columnName)
        {
            switch (columnName)
            {
                case "Timestamp": return entry.Timestamp ?? "";
                case "UserID": return entry.UserID ?? "";
                case "ConfigFile": return entry.ConfigFile ?? "";
                case "Operation": return entry.Operation ?? "";
                case "RowKey": return entry.RowKey ?? "";
                case "ChangeSummary": return entry.ChangeSummary ?? "";
                default: return "";
            }
        }

        // Excel-style filter dropdown for all tabs (History and config tabs)
        private void ShowFilterDropdown(string configName, int columnIndex, DataGridView grid, Point location)
        {
            if (columnIndex < 0 || columnIndex >= grid.Columns.Count)
                return;
            var column = grid.Columns[columnIndex];
            // Always use the original column name (no triangle)
            var columnName = column.Name;
            var filterForm = new Form
            {
                Text = $"Filter - {column.HeaderText.TrimEnd(' ', '▼')}",
                Size = new Size(250, 400),
                StartPosition = FormStartPosition.Manual,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = false,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                KeyPreview = true
            };
            var gridLocation = grid.PointToScreen(Point.Empty);
            filterForm.Location = new Point(gridLocation.X + column.DisplayIndex * 100, gridLocation.Y + grid.ColumnHeadersHeight);
            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(10)
            };
            tableLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var searchBox = new TextBox
            {
                Dock = DockStyle.Top,
                PlaceholderText = "Search...",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Height = 28
            };
            var selectAllCheckBox = new CheckBox
            {
                Text = "Select All",
                Checked = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Dock = DockStyle.Top
            };
            var valuesList = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                CheckOnClick = true
            };
            // Get unique values for this column, but only from currently filtered rows (excluding this column's filter)
            var uniqueValues = new HashSet<string>();
            if (configName == "History")
            {
                var historyEntries = HistoryManager.LoadHistoryEntries();
                // Apply all filters except the one for this column
                if (activeFilters.ContainsKey(configName))
                {
                    foreach (var filter in activeFilters[configName])
                    {
                        if (filter.Key == columnName) continue;
                        var allowed = filter.Value;
                        historyEntries = historyEntries.Where(entry => allowed.Contains(GetHistoryColumnValue(entry, filter.Key))).ToList();
                    }
                }
                foreach (var entry in historyEntries)
                {
                    string value = GetHistoryColumnValue(entry, columnName);
                    if (!string.IsNullOrEmpty(value))
                        uniqueValues.Add(value);
                }
            }
            else if (configFiles.ContainsKey(configName))
            {
                var config = configFiles[configName];
                // Start with all rows, then apply all filters except this column
                var filteredRows = config.Rows;
                if (activeFilters.ContainsKey(configName))
                {
                    foreach (var filter in activeFilters[configName])
                    {
                        if (filter.Key == columnName) continue;
                        var allowed = filter.Value;
                        filteredRows = filteredRows.Where(row => row.ContainsKey(filter.Key) && allowed.Contains(row[filter.Key])).ToList();
                    }
                }
                foreach (var row in filteredRows)
                {
                    var value = row.ContainsKey(columnName) ? row[columnName] : "";
                    if (!string.IsNullOrEmpty(value))
                        uniqueValues.Add(value);
                }
            }
            var sortedValues = uniqueValues.OrderBy(v => v).ToList();
            HashSet<string> selectedValues = null;
            if (activeFilters.ContainsKey(configName) && activeFilters[configName].ContainsKey(columnName))
            {
                selectedValues = new HashSet<string>(activeFilters[configName][columnName]);
            }
            bool selectAll = selectedValues == null;
            foreach (var value in sortedValues)
            {
                bool isChecked = selectAll || (selectedValues != null && selectedValues.Contains(value));
                valuesList.Items.Add(value, isChecked);
            }
            selectAllCheckBox.Checked = selectAll || (selectedValues != null && selectedValues.Count == sortedValues.Count);
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                Padding = new Padding(0, 10, 0, 0)
            };
            var btnCancel = new Button
            {
                Text = "Cancel",
                Size = new Size(75, 30),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            btnCancel.Click += (s, e) => filterForm.Close();
            var btnOK = new Button
            {
                Text = "OK",
                Size = new Size(75, 30),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnOK.FlatAppearance.BorderSize = 0;
            btnOK.Click += (s, e) =>
            {
                var selectedValues = new List<string>();
                foreach (var item in valuesList.CheckedItems)
                {
                    selectedValues.Add(item.ToString());
                }
                System.Diagnostics.Debug.WriteLine($"[DEBUG] FilterDropdown OK: config={configName}, column={columnName}, selectedValues=[{string.Join(",", selectedValues)}]");
                SetConfigFilter(configName, columnName, selectedValues, sortedValues);
                filterForm.Close();
            };
            buttonPanel.Controls.Add(btnCancel);
            buttonPanel.Controls.Add(btnOK);
            filterForm.AcceptButton = btnOK;
            filterForm.CancelButton = btnCancel;
            filterForm.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter && e.Modifiers == Keys.None)
                {
                    e.Handled = true;
                    btnOK.PerformClick();
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    e.Handled = true;
                    filterForm.Close();
                }
            };
            searchBox.TextChanged += (s, e) =>
            {
                var searchText = searchBox.Text.ToLower();
                valuesList.Items.Clear();
                foreach (var value in sortedValues)
                {
                    if (value.ToLower().Contains(searchText))
                    {
                        bool isChecked = selectAll || (selectedValues != null && selectedValues.Contains(value));
                        valuesList.Items.Add(value, isChecked);
                    }
                }
            };
            searchBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    btnOK.PerformClick();
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    e.Handled = true;
                    filterForm.Close();
                }
            };
            selectAllCheckBox.CheckedChanged += (s, e) =>
            {
                for (int i = 0; i < valuesList.Items.Count; i++)
                {
                    valuesList.SetItemChecked(i, selectAllCheckBox.Checked);
                }
            };
            tableLayout.Controls.Add(searchBox, 0, 0);
            tableLayout.Controls.Add(selectAllCheckBox, 0, 1);
            tableLayout.Controls.Add(valuesList, 0, 2);
            tableLayout.Controls.Add(buttonPanel, 0, 3);
            filterForm.Controls.Add(tableLayout);
            filterForm.ShowDialog();
        }

        // Legacy filter method no longer needed; replaced by SetConfigFilter

        private void CreateTabForFile(string fileName, ConfigFile configFile)
        {
            var tabPage = new TabPage(fileName)
            {
                BackColor = Color.FromArgb(250, 250, 250),
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                Padding = new Padding(10)
            };
            
            var grid = CreateDataGrid(configFile);
            grids[fileName] = grid;
            
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                BackColor = Color.FromArgb(250, 250, 250)
            };
            
            var buttonPanel = CreateButtonPanel(fileName, configFile);
            
            panel.Controls.Add(grid);
            panel.Controls.Add(buttonPanel);
            tabPage.Controls.Add(panel);
            tabControl.TabPages.Add(tabPage);
        }

        private DataGridView CreateDataGrid(ConfigFile configFile)
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                GridColor = Color.FromArgb(229, 229, 229),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToResizeRows = false
            };
            
            // This method is no longer used, but kept for compatibility. Use BuildTabs for new tabs.
            return grid;
        }

        private Panel CreateButtonPanel(string fileName, ConfigFile configFile)
        {
            var buttonPanel = new Panel
            {
                Height = 60,
                Dock = DockStyle.Bottom,
                BackColor = Color.FromArgb(250, 250, 250),
                Padding = new Padding(0, 10, 0, 0)
            };
            
            var addButton = new Button
            {
                Text = "Add",
                Size = new Size(80, 35),
                Location = new Point(10, 10),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular)
            };
            addButton.FlatAppearance.BorderSize = 0;
            addButton.Click += (s, e) => AddRow(fileName);
            
            var editButton = new Button
            {
                Text = "Edit",
                Size = new Size(80, 35),
                Location = new Point(100, 10),
                BackColor = Color.FromArgb(16, 137, 62),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular)
            };
            editButton.FlatAppearance.BorderSize = 0;
            editButton.Click += (s, e) => EditRow(fileName);
            
            var deleteButton = new Button
            {
                Text = "Delete",
                Size = new Size(80, 35),
                Location = new Point(190, 10),
                BackColor = Color.FromArgb(196, 43, 28),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular)
            };
            deleteButton.FlatAppearance.BorderSize = 0;
            deleteButton.Click += (s, e) => DeleteRow(fileName);
            
            buttonPanel.Controls.AddRange(new Control[] { addButton, editButton, deleteButton });
            return buttonPanel;
        }


        private void AddRow(string fileName)
        {
            if (!configFiles.ContainsKey(fileName))
            {
                MessageBox.Show($"Config file not found for {fileName}.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // Debug: Show path being used for save
            var debugConfig = configFiles[fileName];
            System.Diagnostics.Debug.WriteLine($"[DEBUG] AddRow: {fileName} will save to {debugConfig.Path}");
            // Multi-user file lock: try to acquire exclusive lock for add
            if (!TryAcquireFileLock(fileName, out string lockError))
            {
                MessageBox.Show(lockError, "File Locked", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                var config = configFiles[fileName];
                // Create a new row dictionary with empty values for all columns
                var newRow = new Dictionary<string, string>();
                foreach (var col in config.Columns)
                    newRow[col] = string.Empty;

                var dlg = new PWQ.Forms.RowEditForm(config.Columns, newRow, fileName);
                var result = dlg.ShowDialog();
                if (result == DialogResult.OK)
        {
            // Show confirmation dialog with two-column table
            var confirmForm = new Form
            {
                Text = "Confirm Add",
                Size = new Size(420, 420),
                StartPosition = FormStartPosition.CenterParent,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                MinimizeBox = false,
                MaximizeBox = false,
                FormBorderStyle = FormBorderStyle.FixedDialog
            };
            var label = new Label
            {
                Text = "Are you sure you want to add this row?",
                Dock = DockStyle.Top,
                Height = 36,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            var table = new DataGridView
            {
                Dock = DockStyle.Top,
                Height = 260,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToResizeColumns = false,
                RowHeadersVisible = false,
                ColumnHeadersHeight = 32,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None
            };
            table.Columns.Add("Column", "Column");
            table.Columns.Add("Value", "Value");
            table.Columns[1].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            table.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            foreach (var kv in dlg.RowData)
            {
                table.Rows.Add(kv.Key, kv.Value);
            }
            table.Columns[0].Width = 140;
            table.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 60,
                Padding = new Padding(0, 10, 0, 0)
            };
            var btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Size = new Size(100, 36),
                Font = new Font("Segoe UI", 10F, FontStyle.Regular)
            };
            var btnAdd = new Button
            {
                Text = "Add",
                DialogResult = DialogResult.OK,
                Size = new Size(100, 36),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                BackColor = Color.FromArgb(16, 137, 62),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnAdd.FlatAppearance.BorderSize = 0;
            btnPanel.Controls.Add(btnAdd);
            btnPanel.Controls.Add(btnCancel);
            confirmForm.Controls.Add(btnPanel);
            confirmForm.Controls.Add(table);
            confirmForm.Controls.Add(label);
            confirmForm.AcceptButton = btnAdd;
            confirmForm.CancelButton = btnCancel;
            var confirmResult = confirmForm.ShowDialog();
            if (confirmResult == DialogResult.OK)
            {
                var newRowData = new Dictionary<string, string>(dlg.RowData);
                // Log to history BEFORE adding
                PWQ.Models.HistoryManager.LogAdd(fileName, newRowData);
                config.Rows.Add(newRowData);
                config.SaveToFile();
                // Reapply filters so filtered view is preserved after add
                ApplyConfigFilters(fileName);
            }
        }
            }
            finally
            {
                ReleaseFileLock(fileName);
            }
        }

        private void EditRow(string fileName, Dictionary<string, string> keyDict)
        {
            if (!grids.ContainsKey(fileName))
            {
                MessageBox.Show("Please select a row to edit.", "No Row Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // Debug: Show path being used for save
            var debugConfig = configFiles[fileName];
            System.Diagnostics.Debug.WriteLine($"[DEBUG] EditRow: {fileName} will save to {debugConfig.Path}");
            // Multi-user file lock: try to acquire exclusive lock for edit
            if (!TryAcquireFileLock(fileName, out string lockError))
            {
                MessageBox.Show(lockError, "File Locked", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                var config = configFiles[fileName];
                // Find the row in config.Rows that matches all key/value pairs in keyDict
                var rowIndex = config.Rows.FindIndex(r => keyDict.All(kv => r.ContainsKey(kv.Key) && r[kv.Key] == kv.Value));
                if (rowIndex < 0)
                {
                    MessageBox.Show("Could not find the selected row in the data.", "Row Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                var rowDict = new Dictionary<string, string>(config.Rows[rowIndex]);
                var dlg = new PWQ.Forms.RowEditForm(config.Columns, rowDict, fileName);
                var result = dlg.ShowDialog();
                if (result == DialogResult.OK)
                {
                    // Only show the nice table confirmation dialog (remove any earlier/extra confirmation)
                    var confirmForm = new Form
                    {
                        Text = "Confirm Save",
                        Size = new Size(420, 420),
                        StartPosition = FormStartPosition.CenterParent,
                        Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                        MinimizeBox = false,
                        MaximizeBox = false,
                        FormBorderStyle = FormBorderStyle.FixedDialog
                    };
                    var label = new Label
                    {
                        Text = "Are you sure you want to save these changes?",
                        Dock = DockStyle.Top,
                        Height = 36,
                        TextAlign = ContentAlignment.MiddleLeft,
                        Font = new Font("Segoe UI", 10F, FontStyle.Bold)
                    };
                    var table = new DataGridView
                    {
                        Dock = DockStyle.Top,
                        Height = 260,
                        ReadOnly = true,
                        AllowUserToAddRows = false,
                        AllowUserToDeleteRows = false,
                        AllowUserToResizeRows = false,
                        AllowUserToResizeColumns = false,
                        RowHeadersVisible = false,
                        ColumnHeadersHeight = 32,
                        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                        Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                        BackgroundColor = Color.White,
                        BorderStyle = BorderStyle.None
                    };
                    table.Columns.Add("Column", "Column");
                    table.Columns.Add("Value", "Value");
                    table.Columns[1].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    table.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    foreach (var kv in dlg.RowData)
                    {
                        table.Rows.Add(kv.Key, kv.Value);
                    }
                    table.Columns[0].Width = 140;
                    table.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    var btnPanel = new FlowLayoutPanel
                    {
                        Dock = DockStyle.Bottom,
                        FlowDirection = FlowDirection.RightToLeft,
                        Height = 60,
                        Padding = new Padding(0, 10, 0, 0)
                    };
                    var btnCancel = new Button
                    {
                        Text = "Cancel",
                        DialogResult = DialogResult.Cancel,
                        Size = new Size(100, 36),
                        Font = new Font("Segoe UI", 10F, FontStyle.Regular)
                    };
                    var btnSave = new Button
                    {
                        Text = "Save",
                        DialogResult = DialogResult.OK,
                        Size = new Size(100, 36),
                        Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                        BackColor = Color.FromArgb(16, 137, 62),
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat
                    };
                    btnSave.FlatAppearance.BorderSize = 0;
                    btnPanel.Controls.Add(btnSave);
                    btnPanel.Controls.Add(btnCancel);
                    confirmForm.Controls.Add(btnPanel);
                    confirmForm.Controls.Add(table);
                    confirmForm.Controls.Add(label);
                    confirmForm.AcceptButton = btnSave;
                    confirmForm.CancelButton = btnCancel;
                    var confirmResult = confirmForm.ShowDialog();
                    if (confirmResult == DialogResult.OK)
                    {
                        // Log edit to history BEFORE updating the row
                        var oldRow = new Dictionary<string, string>(config.Rows[rowIndex]);
                        var newRow = new Dictionary<string, string>(dlg.RowData);
                        HistoryManager.LogEdit(fileName, oldRow, newRow);
                        config.Rows[rowIndex] = dlg.RowData;
                        config.SaveToFile();
                        ApplyConfigFilters(fileName);
                    }
                }
            }
            finally
            {
                ReleaseFileLock(fileName);
            }
        }

        // Overload for legacy button click usage (edit selected row)
        private void EditRow(string fileName)
        {
            var grid = grids.ContainsKey(fileName) ? grids[fileName] : null;
            if (grid == null || grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a row to edit.", "No Row Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var row = grid.SelectedRows[0];
            var keyDict = new Dictionary<string, string>();
            foreach (DataGridViewColumn col in grid.Columns)
            {
                keyDict[col.Name] = row.Cells[col.Index].Value?.ToString() ?? "";
            }
            EditRow(fileName, keyDict);
        }

        private void DeleteRow(string fileName)
        {
            // Multi-user file lock: try to acquire exclusive lock for delete
            // Debug: Show path being used for save
            if (configFiles.ContainsKey(fileName))
            {
                var debugConfig = configFiles[fileName];
                System.Diagnostics.Debug.WriteLine($"[DEBUG] DeleteRow: {fileName} will save to {debugConfig.Path}");
            }
            if (!TryAcquireFileLock(fileName, out string lockError))
            {
                MessageBox.Show(lockError, "File Locked", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                // Delete row logic: persist deletion and log to history
                if (grids.ContainsKey(fileName) && grids[fileName].SelectedRows.Count > 0)
                {
                    var grid = grids[fileName];
                    var selectedIndex = grid.SelectedRows[0].Index;
                    if (selectedIndex >= 0 && selectedIndex < configFiles[fileName].Rows.Count)
                    {
                        var rowDict = configFiles[fileName].Rows[selectedIndex];
                        // Show a two-column table for confirmation
                        var confirmForm = new Form
                        {
                            Text = "Confirm Delete",
                            Size = new Size(420, 420),
                            StartPosition = FormStartPosition.CenterParent,
                            Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                            MinimizeBox = false,
                            MaximizeBox = false,
                            FormBorderStyle = FormBorderStyle.FixedDialog
                        };
                        var label = new Label
                        {
                            Text = "Are you sure you want to delete this row?",
                            Dock = DockStyle.Top,
                            Height = 36,
                            TextAlign = ContentAlignment.MiddleLeft,
                            Font = new Font("Segoe UI", 10F, FontStyle.Bold)
                        };
                        var table = new DataGridView
                        {
                            Dock = DockStyle.Top,
                            Height = 260,
                            ReadOnly = true,
                            AllowUserToAddRows = false,
                            AllowUserToDeleteRows = false,
                            AllowUserToResizeRows = false,
                            AllowUserToResizeColumns = false,
                            RowHeadersVisible = false,
                            ColumnHeadersHeight = 32,
                            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                            Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                            BackgroundColor = Color.White,
                            BorderStyle = BorderStyle.None
                        };
                        table.Columns.Add("Column", "Column");
                        table.Columns.Add("Value", "Value");
                        table.Columns[1].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        table.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        foreach (var kv in rowDict)
                        {
                            table.Rows.Add(kv.Key, kv.Value);
                        }
                        table.Columns[0].Width = 140;
                        table.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        var btnPanel = new FlowLayoutPanel
                        {
                            Dock = DockStyle.Bottom,
                            FlowDirection = FlowDirection.RightToLeft,
                            Height = 60,
                            Padding = new Padding(0, 10, 0, 0)
                        };
                        var btnCancel = new Button
                        {
                            Text = "Cancel",
                            DialogResult = DialogResult.Cancel,
                            Size = new Size(100, 36),
                            Font = new Font("Segoe UI", 10F, FontStyle.Regular)
                        };
                        var btnDelete = new Button
                        {
                            Text = "Delete",
                            DialogResult = DialogResult.OK,
                            Size = new Size(100, 36),
                            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                            BackColor = Color.FromArgb(196, 43, 28),
                            ForeColor = Color.White,
                            FlatStyle = FlatStyle.Flat
                        };
                        btnDelete.FlatAppearance.BorderSize = 0;
                        btnPanel.Controls.Add(btnDelete);
                        btnPanel.Controls.Add(btnCancel);
                        confirmForm.Controls.Add(btnPanel);
                        confirmForm.Controls.Add(table);
                        confirmForm.Controls.Add(label);
                        confirmForm.AcceptButton = btnDelete;
                        confirmForm.CancelButton = btnCancel;
                        var result = confirmForm.ShowDialog();
                        if (result == DialogResult.OK)
                        {
                            // Capture the row to be deleted for history
                            var deletedRow = new Dictionary<string, string>(rowDict);
                            // Remove from data
                            configFiles[fileName].Rows.RemoveAt(selectedIndex);
                            // Persist to CSV
                            configFiles[fileName].SaveToFile();
                            // Log to history
                            HistoryManager.LogDelete(fileName, deletedRow);
                            // Reapply filters so filtered view is preserved after delete
                            ApplyConfigFilters(fileName);
                            // If a config file was deleted, refresh the history tab if visible
                            var selectedTab = tabControl.SelectedTab;
                            if (selectedTab != null && selectedTab.Text == "History")
                            {
                                RefreshHistoryGrid();
                            }
                        }
                    }
                }
            }
            finally
            {
                ReleaseFileLock(fileName);
            }
        }
        // Multi-user file locking: Only for add/edit/delete actions. Everyone can view.
        // Uses a lock file (e.g., config.csv.lock) to coordinate across machines.
        private bool TryAcquireFileLock(string fileName, out string errorMessage)
        {
            errorMessage = null;
            if (!configFiles.ContainsKey(fileName))
            {
                errorMessage = $"Config file not found for {fileName}.";
                return false;
            }
            var config = configFiles[fileName];
            string lockFilePath = config.Path + ".lock";
            string user = Environment.UserName;
            string machine = Environment.MachineName;
            string lockInfo = $"{user}@{machine}\n{DateTime.Now:yyyy-MM-dd HH:mm:ss}";
            try
            {
                // Try to create the lock file exclusively
                var fs = new FileStream(lockFilePath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                using (var sw = new StreamWriter(fs))
                {
                    sw.Write(lockInfo);
                }
                // Keep a handle so we can delete it later
                fileLocks[fileName] = new FileStream(lockFilePath, FileMode.Open, FileAccess.Read, FileShare.None);
                lockOwners[fileName] = user + "@" + machine;
                return true;
            }
            catch (IOException)
            {
                // Lock file exists, read who owns it
                try
                {
                    string owner = "another user";
                    string time = "unknown time";
                    if (File.Exists(lockFilePath))
                    {
                        var lines = File.ReadAllLines(lockFilePath);
                        if (lines.Length > 0) owner = lines[0];
                        if (lines.Length > 1) time = lines[1];
                    }
                    errorMessage = $"This file is currently being edited by {owner} (since {time}). Please try again later.";
                }
                catch
                {
                    errorMessage = "This file is currently locked by another user.";
                }
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = $"Could not acquire lock: {ex.Message}";
                return false;
            }
        }

        private void ReleaseFileLock(string fileName)
        {
            if (!configFiles.ContainsKey(fileName)) return;
            var config = configFiles[fileName];
            string lockFilePath = config.Path + ".lock";
            if (fileLocks.ContainsKey(fileName))
            {
                try
                {
                    fileLocks[fileName]?.Close();
                    fileLocks[fileName]?.Dispose();
                }
                catch { }
                fileLocks.Remove(fileName);
            }
            if (File.Exists(lockFilePath))
            {
                try { File.Delete(lockFilePath); } catch { }
            }
            if (lockOwners.ContainsKey(fileName))
                lockOwners.Remove(fileName);
        }

        private void RefreshAll()
        {
            // Clear all filters before reloading
            activeFilters.Clear();
            LoadConfigFiles();
            BuildTabs();
        }

        private void ShowAbout()
        {
            // Custom About dialog with hyperlink for LHD
            var aboutForm = new Form
            {
                Text = "About",
                Size = new Size(520, 270),
                StartPosition = FormStartPosition.CenterParent,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                MinimizeBox = false,
                MaximizeBox = false,
                FormBorderStyle = FormBorderStyle.FixedDialog
            };
            var lblTitle = new Label
            {
                Text = "PWQ",
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                Dock = DockStyle.Top,
                Height = 36,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(16, 12, 0, 0)
            };
            var lblDesc = new Label
            {
                Text = "A tool for litho layer owner to log litho's process window after CD measurement on a PWQ lot. The output of this report will be read by DefMet to build the final report. Please contact sanjay.adhikari@intel.com for any issues/feedback.",
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = 110,
                Width = 480,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                Padding = new Padding(16, 16, 16, 0)
            };
            var btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Dock = DockStyle.Bottom,
                Height = 40,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnOK.FlatAppearance.BorderSize = 0;
            aboutForm.AcceptButton = btnOK;
            aboutForm.CancelButton = btnOK;
            aboutForm.Controls.Add(btnOK);
            aboutForm.Controls.Add(lblDesc);
            aboutForm.Controls.Add(lblTitle);
            aboutForm.ShowDialog(this);
        }
        private void ReleaseAllLocks()
        {
            foreach (var lockStream in fileLocks.Values)
            {
                try
                {
                    lockStream?.Close();
                    lockStream?.Dispose();
                }
                catch { }
            }
            fileLocks.Clear();
            lockOwners.Clear();
        }
    }
}
