
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using System.IO;

namespace PWQ.Forms
{
    public partial class RowEditForm : Form
    {
        // Allowed CLUSTER_NAME values (from LayerList.csv)
        private static readonly HashSet<string> AllowedClusterNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "NIKONWET", "LIDRY193", "LIDRY248", "ASMLWET", "ASMLEUV"
        };
        public Dictionary<string, string> RowData { get; private set; }
        private List<string> columns;
        private Dictionary<string, TextBox> textBoxes = new Dictionary<string, TextBox>();
        private Dictionary<string, ComboBox> comboBoxes = new Dictionary<string, ComboBox>();
        private Dictionary<string, Label> errorLabels = new Dictionary<string, Label>();

        public string TabName { get; set; } // Add this property to set from MainForm

        public RowEditForm()
        {
            InitializeComponent();
        }

        public RowEditForm(List<string> columns, Dictionary<string, string> initial = null, string tabName = null)
        {
            // InitializeComponent(); // Not needed, UI is built in code
            this.columns = columns;
            RowData = new Dictionary<string, string>();
            TabName = tabName; // Set TabName before creating controls
            SetIcon();
            CreateFormControls(initial);
        }


        private void SetIcon()
        {
            try
            {
                // Load the icon from the file
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
                }
            }
            catch (Exception ex)
            {
                // If icon loading fails, just continue without icon
                Console.WriteLine($"Could not load icon: {ex.Message}");
            }
        }

        private void CreateFormControls(Dictionary<string, string> initial)
        {
            // Clear existing controls
            Controls.Clear();
            textBoxes.Clear();
            comboBoxes.Clear();
            errorLabels.Clear();

            int y = 10;
            int maxLabelWidth = 120;
            int maxInputWidth = 300;
            int spacingY = 10;
            int controlHeight = 28;
            int errorLabelHeight = 20;
            int fieldPadding = 8;
            var controlsList = new List<Control>();
            foreach (var col in columns)
            {
                var lbl = new Label { Text = col, Left = 10, Top = y, Width = maxLabelWidth, Height = controlHeight };
                Control inputControl;
                if (col.Equals("userID", StringComparison.OrdinalIgnoreCase))
                {
                    var txt = new TextBox { Left = 10 + maxLabelWidth + fieldPadding, Top = y, Width = maxInputWidth, Height = controlHeight, ReadOnly = true, BackColor = Color.LightGray, ForeColor = Color.DarkGray };
                    txt.Text = Environment.UserName;
                    inputControl = txt;
                    textBoxes[col] = txt;
                }
                else if (col.Equals("timestamp", StringComparison.OrdinalIgnoreCase))
                {
                    var txt = new TextBox { Left = 10 + maxLabelWidth + fieldPadding, Top = y, Width = maxInputWidth, Height = controlHeight, ReadOnly = true, BackColor = Color.LightGray, ForeColor = Color.DarkGray };
                    txt.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    inputControl = txt;
                    textBoxes[col] = txt;
                }
                else if (col.Equals("FE/BE", StringComparison.OrdinalIgnoreCase))
                {
                    var cmb = new ComboBox { Left = 10 + maxLabelWidth + fieldPadding, Top = y, Width = maxInputWidth, Height = controlHeight, DropDownStyle = ComboBoxStyle.DropDownList };
                    cmb.Items.AddRange(new[] { "FE", "BE" });
                    if (initial != null && initial.ContainsKey(col))
                        cmb.SelectedItem = initial[col];
                    cmb.SelectedIndexChanged += (s, e) => ValidateField(col, cmb.SelectedItem?.ToString());
                    inputControl = cmb;
                    comboBoxes[col] = cmb;
                }
                else if (col.Equals("KEY", StringComparison.OrdinalIgnoreCase) && TabName != null && TabName.Equals("LayerList", StringComparison.OrdinalIgnoreCase))
                {
                    // Editable KEY for LayerList: KEY = PROCESS_LAYER_CLUSTER_NAME, but user can type
                    var txt = new TextBox { Left = 10 + maxLabelWidth + fieldPadding, Top = y, Width = maxInputWidth, Height = controlHeight, ReadOnly = false, BackColor = Color.White, ForeColor = Color.Black };
                    string process = initial != null && initial.ContainsKey("PROCESS") ? initial["PROCESS"] : "";
                    string layer = initial != null && initial.ContainsKey("LAYER") ? initial["LAYER"] : "";
                    string cluster = initial != null && initial.ContainsKey("CLUSTER_NAME") ? initial["CLUSTER_NAME"] : "";
                    string autoKey = $"{process}_{layer}_{cluster}";
                    txt.Text = autoKey;
                    textBoxes[col] = txt;
                    inputControl = txt;

                    // Listen for changes in PROCESS, LAYER, CLUSTER_NAME to update KEY suggestion
                    void updateKey()
                    {
                        string p = textBoxes.ContainsKey("PROCESS") ? textBoxes["PROCESS"].Text.Trim() : "";
                        string l = textBoxes.ContainsKey("LAYER") ? textBoxes["LAYER"].Text.Trim() : "";
                        string c = textBoxes.ContainsKey("CLUSTER_NAME") ? textBoxes["CLUSTER_NAME"].Text.Trim() : "";
                        string autoKeyNow = $"{p}_{l}_{c}";
                        // Only update if user hasn't changed it manually
                        if (txt.Focused == false || txt.Text == autoKey || string.IsNullOrWhiteSpace(txt.Text))
                        {
                            txt.Text = autoKeyNow;
                        }
                        autoKey = autoKeyNow;
                        ValidateField("KEY", txt.Text); // Validate after update
                    }
                    txt.TextChanged += (s, e) => { ValidateField("KEY", txt.Text); };
                    txt.KeyUp += (s, e) => { ValidateField("KEY", txt.Text); };
                    // Attach event handlers after all controls are created (see below)
                }
                else
                {
                    var txt = new TextBox { Left = 10 + maxLabelWidth + fieldPadding, Top = y, Width = maxInputWidth, Height = controlHeight, Multiline = true, ScrollBars = ScrollBars.Vertical };
                    string initialValue = "";
                    if (initial != null && initial.ContainsKey(col))
                        initialValue = initial[col] ?? "";
                    if (TabName != null && TabName.Equals("Run Plan", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(initialValue) && (col.Equals("PRODUCT", StringComparison.OrdinalIgnoreCase) || col.Equals("ROUTE_TYPE", StringComparison.OrdinalIgnoreCase)))
                    {
                        txt.Text = "*";
                        this.Text = $"PREFILL SUCCESS: Set {col} = '*' for tab '{TabName}'";
                    }
                    else
                    {
                        txt.Text = initialValue;
                        this.Text = $"SET VALUE: {col} = '{initialValue}' (Tab: '{TabName}')";
                    }
                    txt.TextChanged += (s, e) => {
                        AdjustTextBoxHeight(txt);
                        ValidateField(col, txt.Text);
                        RepositionControls();
                    };
                    txt.KeyUp += (s, e) => {
                        ValidateField(col, txt.Text);
                    };
                    AdjustTextBoxHeight(txt);
                    inputControl = txt;
                    textBoxes[col] = txt;
                }
                controlsList.Add(lbl);
                controlsList.Add(inputControl);
                Controls.Add(lbl);
                Controls.Add(inputControl);
                // Error label
                var errorLbl = new Label
                {
                    Left = 10 + maxLabelWidth + fieldPadding,
                    Top = y + controlHeight + 2,
                    Width = maxInputWidth,
                    Height = errorLabelHeight,
                    ForeColor = Color.Red,
                    Text = "",
                    Visible = false,
                    Font = new Font(this.Font, FontStyle.Italic)
                };
                Controls.Add(errorLbl);
                errorLabels[col] = errorLbl;
                // Calculate next y
                y += controlHeight + errorLabelHeight + spacingY;
            }

            // Attach event handlers for KEY auto-update (after all controls are created)
            if (TabName != null && TabName.Equals("LayerList", StringComparison.OrdinalIgnoreCase) && textBoxes.ContainsKey("KEY"))
            {
                var txtKey = textBoxes["KEY"];
                string autoKey = txtKey.Text;
                void updateKey()
                {
                    string p = textBoxes.ContainsKey("PROCESS") ? textBoxes["PROCESS"].Text.Trim() : "";
                    string l = textBoxes.ContainsKey("LAYER") ? textBoxes["LAYER"].Text.Trim() : "";
                    string c = textBoxes.ContainsKey("CLUSTER_NAME") ? textBoxes["CLUSTER_NAME"].Text.Trim() : "";
                    string autoKeyNow = $"{p}_{l}_{c}";
                    // Only update if user hasn't changed it manually
                    if (txtKey.Focused == false || txtKey.Text == autoKey || string.IsNullOrWhiteSpace(txtKey.Text))
                    {
                        txtKey.Text = autoKeyNow;
                    }
                    autoKey = autoKeyNow;
                    ValidateField("KEY", txtKey.Text); // Validate after update
                }
                if (textBoxes.ContainsKey("PROCESS"))
                    textBoxes["PROCESS"].TextChanged += (s, e) => updateKey();
                if (textBoxes.ContainsKey("LAYER"))
                    textBoxes["LAYER"].TextChanged += (s, e) => updateKey();
                if (textBoxes.ContainsKey("CLUSTER_NAME"))
                    textBoxes["CLUSTER_NAME"].TextChanged += (s, e) => updateKey();
            }
            // Add Save button at the bottom, centered
            var btnSave = new Button
            {
                Text = "Save",
                Width = 120,
                Height = 38,
                Left = (10 + maxLabelWidth + fieldPadding + maxInputWidth) / 2 - 60,
                Top = y + 10,
                Font = new Font(this.Font.FontFamily, 11F, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += (s, e) => { SaveRow(); };
            Controls.Add(btnSave);
            // Resize form to fit all controls
            this.Height = btnSave.Top + btnSave.Height + 40;
            this.Width = 10 + maxLabelWidth + fieldPadding + maxInputWidth + 40;
        }

        // Real-time field validation for editing
        private void ValidateField(string col, string value)
        {
            string colUpper = col.ToUpperInvariant();
            string error = null;
            // CLUSTER_NAME validation: only allow 5 specific values
            if (colUpper == "CLUSTER_NAME")
            {
                if (!AllowedClusterNames.Contains(value))
                {
                    error = $"CLUSTER_NAME must be one of: {string.Join(", ", AllowedClusterNames)}.";
                }
            }
            if (string.IsNullOrWhiteSpace(value))
            {
                error = $"{col} cannot be empty.";
            }
            else if (colUpper == "KEY" && !string.IsNullOrEmpty(TabName) && TabName.Equals("LayerList", System.StringComparison.OrdinalIgnoreCase))
            {
                string p = textBoxes != null && textBoxes.ContainsKey("PROCESS") ? textBoxes["PROCESS"].Text.Trim() : "";
                string l = textBoxes != null && textBoxes.ContainsKey("LAYER") ? textBoxes["LAYER"].Text.Trim() : "";
                string c = textBoxes != null && textBoxes.ContainsKey("CLUSTER_NAME") ? textBoxes["CLUSTER_NAME"].Text.Trim() : "";
                string expected = $"{p}_{l}_{c}";
                if ((!string.IsNullOrWhiteSpace(p) || !string.IsNullOrWhiteSpace(l) || !string.IsNullOrWhiteSpace(c)) && value != expected)
                {
                    error = $"KEY must match '{{PROCESS}}_{{LAYER}}_{{CLUSTER_NAME}}' ({expected})";
                }
            }
            else if (colUpper == "RETICLE")
            {
                if ((value.Length != 13 && value.Length != 14) || !System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Za-z0-9]{13,14}$"))
                    error = $"{col} must be 13 or 14 alphanumeric characters.";
            }
            else if (colUpper.Contains("COUNT") || colUpper.Contains("DAYS") || colUpper.Contains("XQUALS"))
            {
                if (!int.TryParse(value, out int intVal) || intVal < 0)
                    error = $"{col} must be a positive integer.";
            }
            else if (colUpper == "FE/BE")
            {
                if (value != "FE" && value != "BE")
                    error = $"{col} must be either 'FE' or 'BE'.";
            }
            else if (colUpper == "OPERATION" || colUpper == "OPERATIONS")
            {
                var operations = value.Split(',');
                foreach (var op in operations)
                {
                    var opTrimmed = op.Trim();
                    if (string.IsNullOrEmpty(opTrimmed))
                        continue;
                    if (opTrimmed.Length != 6 || !System.Text.RegularExpressions.Regex.IsMatch(opTrimmed, @"^\d{6}$"))
                    {
                        error = $"Each operation must be exactly 6 digits (e.g. 123456).";
                        break;
                    }
                }
            }
            else if (colUpper == "PROCESS")
            {
                if (value.Length != 4 || !System.Text.RegularExpressions.Regex.IsMatch(value, @"^\d{4}$"))
                {
                    error = $"{col} must be exactly 4 numeric characters (e.g. 1234).";
                }
            }
            else if (colUpper == "ENTITY" || colUpper == "TOOL" || colUpper == "TOOLS" || colUpper == "ALLOWED_TOOLS")
            {
                var items = value.Split(',');
                foreach (var item in items)
                {
                    var trimmed = item.Trim();
                    if (string.IsNullOrEmpty(trimmed))
                        continue;
                    if (trimmed.Length != 6 || !System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^[A-Za-z0-9]{6}$"))
                    {
                        error = $"Each {(colUpper == "ENTITY" ? "entity" : "tool")} must be exactly 6 alphanumeric characters (e.g. ABC123).";
                        break;
                    }
                }
            }
            else if (colUpper == "LAYER")
            {
                if (value.Length != 3 || !System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Za-z0-9]{3}$"))
                {
                    error = $"{col} must be exactly 3 alphanumeric characters.";
                }
            }
            else if (colUpper == "TOOLMATRIX")
            {
                // ToolMatrix: comma-separated list of 6-char alphanumeric Entity codes
                var items = value.Split(',');
                foreach (var item in items)
                {
                    var trimmed = item.Trim();
                    if (string.IsNullOrEmpty(trimmed))
                        continue;
                    if (trimmed.Length != 6 || !System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^[A-Za-z0-9]{6}$"))
                    {
                        error = $"Each ToolMatrix entry must be a 6-character alphanumeric Entity code (e.g. ABC123), separated by commas.";
                        break;
                    }
                }
            }
            else if (colUpper == "POR_METROLOGY")
            {
                if (value != "OCD" && value != "CDSEM")
                {
                    error = $"{col} must be either 'OCD' or 'CDSEM'.";
                }
            }

            // Show/hide error label and color
            if (errorLabels.ContainsKey(col))
            {
                var label = errorLabels[col];
                if (error != null)
                {
                    // Highlight whitespace at start/end if present
                    string valueForWhitespace = textBoxes.ContainsKey(col) ? textBoxes[col].Text : comboBoxes.ContainsKey(col) ? comboBoxes[col].Text : null;
                    bool hasWhitespace = !string.IsNullOrEmpty(valueForWhitespace) && (valueForWhitespace.StartsWith(" ") || valueForWhitespace.EndsWith(" "));
                    if (hasWhitespace)
                    {
                        // Show whitespace visually in error label
                        string visibleWhitespace = valueForWhitespace.Replace(" ", "·");
                        label.Text = error + $" (Whitespace: '{visibleWhitespace}')";
                        label.Font = new Font(label.Font, FontStyle.Bold | FontStyle.Italic);
                        label.Visible = true;
                        if (textBoxes.ContainsKey(col))
                        {
                            textBoxes[col].BackColor = Color.LightPink;
                            // Add red underline for whitespace
                            textBoxes[col].BorderStyle = BorderStyle.FixedSingle;
                            textBoxes[col].Padding = new Padding(0, 0, 0, 2);
                            textBoxes[col].Invalidate();
                        }
                        else if (comboBoxes.ContainsKey(col))
                        {
                            comboBoxes[col].BackColor = Color.LightPink;
                        }
                    }
                    else
                    {
                        label.Text = error;
                        label.Font = new Font(label.Font, FontStyle.Italic);
                        label.Visible = true;
                        if (textBoxes.ContainsKey(col))
                        {
                            textBoxes[col].BackColor = Color.LightPink;
                            textBoxes[col].BorderStyle = BorderStyle.Fixed3D;
                        }
                        else if (comboBoxes.ContainsKey(col))
                            comboBoxes[col].BackColor = Color.LightPink;
                    }
                }
                else
                {
                    label.Visible = false;
                    if (textBoxes.ContainsKey(col))
                        textBoxes[col].BackColor = Color.White;
                    else if (comboBoxes.ContainsKey(col))
                        comboBoxes[col].BackColor = Color.White;
                }
            }
        }
        private void RepositionControls()
        {
            int y = 10;
            foreach (var col in columns)
            {
                // Find the label for this column
                var lbl = Controls.OfType<Label>().FirstOrDefault(l => l.Text == col);
                
                if (lbl != null)
                {
                    lbl.Top = y;
                    
                    // Check if it's a textbox or combobox
                    if (textBoxes.ContainsKey(col))
                    {
                        var txt = textBoxes[col];
                        txt.Top = y;
                        
                        // Position error label below the textbox
                        if (errorLabels.ContainsKey(col))
                        {
                            errorLabels[col].Top = y + txt.Height + 2;
                        }
                        
                        y += txt.Height + 30; // Extra space for error label
                    }
                    else if (comboBoxes.ContainsKey(col))
                    {
                        var cmb = comboBoxes[col];
                        cmb.Top = y;
                        
                        // Position error label below the combobox
                        if (errorLabels.ContainsKey(col))
                        {
                            errorLabels[col].Top = y + 25;
                        }
                        
                        y += 50; // Fixed height for combobox + error label
                    }
                }
            }

            // Reposition save button
            var btnSave = Controls.OfType<Button>().FirstOrDefault(b => b.Text == "Save");
            if (btnSave != null)
            {
                btnSave.Top = y + 10;
            }

            // Adjust form height
            this.Height = y + 80;
        }

        private void AdjustTextBoxHeight(TextBox txt)
        {
            // Calculate needed height for all lines
            using (Graphics g = txt.CreateGraphics())
            {
                int border = txt.Height - txt.ClientSize.Height;
                SizeF size = g.MeasureString(txt.Text + "\n", txt.Font, txt.Width);
                int neededHeight = (int)Math.Ceiling(size.Height) + border + 8;
                txt.Height = Math.Max(32, neededHeight);
            }
        }

        private void SaveRow()
        {
            // Enhanced: Check for empty fields and type/format based on column name (enforced for all tables)
            foreach (var col in columns)
            {
                string value = null;
                if (textBoxes.ContainsKey(col))
                    value = textBoxes[col].Text?.Trim() ?? "";
                else if (comboBoxes.ContainsKey(col))
                    value = comboBoxes[col].Text?.Trim() ?? "";
                else
                    value = RowData.ContainsKey(col) ? RowData[col]?.Trim() ?? "" : "";

                // Final check for errors (should be caught by real-time validation)
                string colUpper = col.ToUpperInvariant();
                if (string.IsNullOrWhiteSpace(value))
                {
                    MessageBox.Show($"{col} cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (colUpper == "RETICLE" && ((value.Length != 13 && value.Length != 14) || !System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Za-z0-9]{13,14}$")))
                {
                    MessageBox.Show($"{col} must be 13 or 14 alphanumeric characters.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if ((colUpper.Contains("COUNT") || colUpper.Contains("DAYS") || colUpper.Contains("XQUALS")) && (!int.TryParse(value, out int intVal) || intVal < 0))
                {
                    MessageBox.Show($"{col} must be a positive integer.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (colUpper == "FE/BE" && (value != "FE" && value != "BE"))
                {
                    MessageBox.Show($"{col} must be either 'FE' or 'BE'.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (colUpper == "LAYER" && (value.Length != 3 || !System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Za-z0-9]{3}$")))
                {
                    MessageBox.Show($"{col} must be exactly 3 alphanumeric characters.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                // Enforce Operation field to be 6-digit numeric (for all tabs)
                if (colUpper == "OPERATION" || colUpper == "OPERATIONS")
                {
                    var operations = value.Split(',');
                    foreach (var op in operations)
                    {
                        var opTrimmed = op.Trim();
                        if (string.IsNullOrEmpty(opTrimmed))
                            continue;
                        if (opTrimmed.Length != 6 || !System.Text.RegularExpressions.Regex.IsMatch(opTrimmed, "^\\d{6}$"))
                        {
                            MessageBox.Show($"'{opTrimmed}' is invalid. Each operation must be exactly 6 digits. Example: 123456", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }

            // Check for real-time validation errors first
            if (HasValidationErrors())
            {
                MessageBox.Show("Please fix the validation errors shown in red before saving.", "Validation Errors", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (var col in columns)
            {
                if (col.Equals("userID", StringComparison.OrdinalIgnoreCase))
                    RowData[col] = Environment.UserName;
                else if (col.Equals("timestamp", StringComparison.OrdinalIgnoreCase))
                    RowData[col] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                else if (comboBoxes.ContainsKey(col))
                    RowData[col] = comboBoxes[col].SelectedItem?.ToString() ?? "";
                else if (textBoxes.ContainsKey(col))
                    RowData[col] = textBoxes[col].Text;
            }


            // Enhanced: Check for empty fields and type/format based on column name (enforced for all tables)
            foreach (var col in columns)
            {
                string value = RowData[col]?.Trim() ?? "";
                if (string.IsNullOrEmpty(value))
                {
                    MessageBox.Show($"{col} cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (textBoxes.ContainsKey(col))
                        textBoxes[col].BackColor = Color.LightPink;
                    else if (comboBoxes.ContainsKey(col))
                        comboBoxes[col].BackColor = Color.LightPink;
                    return;
                }

                string colUpper = col.ToUpperInvariant();
                // Integer fields: COUNT, DAYS, XQUALS
                if (colUpper.Contains("COUNT") || colUpper.Contains("DAYS") || colUpper.Contains("XQUALS"))
                {
                    if (!int.TryParse(value, out int intVal) || intVal < 0)
                    {
                        MessageBox.Show($"{col} must be a positive integer.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (textBoxes.ContainsKey(col))
                            textBoxes[col].BackColor = Color.LightPink;
                        return;
                    }
                }

                // FE/BE must be FE or BE
                if (colUpper == "FE/BE")
                {
                    if (value != "FE" && value != "BE")
                    {
                        MessageBox.Show($"{col} must be either 'FE' or 'BE'.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (comboBoxes.ContainsKey(col))
                            comboBoxes[col].BackColor = Color.LightPink;
                        return;
                    }
                }

                // LAYER: 3-char alphanumeric
                if (colUpper == "LAYER")
                {
                    if (value.Length != 3 || !System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Za-z0-9]{3}$"))
                    {
                        MessageBox.Show($"{col} must be exactly 3 alphanumeric characters.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (textBoxes.ContainsKey(col))
                            textBoxes[col].BackColor = Color.LightPink;
                        return;
                    }
                }
            }

            // ...existing tab-specific validation and confirmation logic...
            // OPC Config tab validation
            if (TabName == "OPC Config")
            {
                // FE/BE validation - must be selected
                if (RowData.ContainsKey("FE/BE"))
                {
                    string febeValue = RowData["FE/BE"]?.Trim();
                    if (string.IsNullOrEmpty(febeValue) || (febeValue != "FE" && febeValue != "BE"))
                    {
                        MessageBox.Show("FE/BE must be selected as either 'FE' or 'BE'.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // Layer: 3 character string
                if (RowData.ContainsKey("Layer") || RowData.ContainsKey("layer") || RowData.ContainsKey("Layers") || RowData.ContainsKey("layers"))
                {
                    string layerValue = RowData.ContainsKey("Layer") ? RowData["Layer"]?.Trim() :
                                       RowData.ContainsKey("layer") ? RowData["layer"]?.Trim() :
                                       RowData.ContainsKey("Layers") ? RowData["Layers"]?.Trim() :
                                       RowData["layers"]?.Trim();
                    if (string.IsNullOrEmpty(layerValue) || layerValue.Length != 3 || !System.Text.RegularExpressions.Regex.IsMatch(layerValue, "^[A-Za-z0-9]{3}$"))
                    {
                        MessageBox.Show("Layer must be exactly 3 characters (letters and numbers only).", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // OPC: 7 character string with 4 numbers and 3 dots (format: 0.0.0.0)
                if (RowData.ContainsKey("OPC") || RowData.ContainsKey("OPC rev") || RowData.ContainsKey("OPC Rev"))
                {
                    string opcValue = RowData.ContainsKey("OPC") ? RowData["OPC"]?.Trim() : 
                                     RowData.ContainsKey("OPC rev") ? RowData["OPC rev"]?.Trim() : 
                                     RowData["OPC Rev"]?.Trim();
                    if (string.IsNullOrEmpty(opcValue))
                    {
                        MessageBox.Show("OPC Rev field cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    // Must be exactly 7 characters and match the pattern: digit.digit.digit.digit (e.g. 1.2.3.4)
                    if (!System.Text.RegularExpressions.Regex.IsMatch(opcValue, "^\\d\\.\\d\\.\\d\\.\\d$"))
                    {
                        MessageBox.Show("OPC Rev must be exactly 7 characters in the format 0.0.0.0 (digits and periods only). Example: 1.2.3.4", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // Operation: comma and space separated 6-character number strings
                if (RowData.ContainsKey("Operation") || RowData.ContainsKey("Operations"))
                {
                    string operationValue = RowData.ContainsKey("Operation") ? RowData["Operation"]?.Trim() : RowData["Operations"]?.Trim();
                    
                    if (string.IsNullOrEmpty(operationValue))
                    {
                        MessageBox.Show("Operation field cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Check for invalid characters (anything other than digits, commas, and spaces)
                    if (System.Text.RegularExpressions.Regex.IsMatch(operationValue, "[^0-9,\\s]"))
                    {
                        MessageBox.Show("Operation field contains invalid characters. Only digits (0-9), commas, and spaces are allowed.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Split by comma and validate each operation code
                    var operations = operationValue.Split(',');
                    var validOperations = new List<string>();

                    foreach (var op in operations)
                    {
                        var opTrimmed = op.Trim();
                        
                        // Skip empty parts
                        if (string.IsNullOrEmpty(opTrimmed))
                            continue;

                        // Must be exactly 6 digits
                        if (opTrimmed.Length != 6 || !System.Text.RegularExpressions.Regex.IsMatch(opTrimmed, "^\\d{6}$"))
                        {
                            MessageBox.Show($"'{opTrimmed}' is invalid. Each operation must be exactly 6 digits. Example: 123456", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        validOperations.Add(opTrimmed);
                    }

                    if (validOperations.Count == 0)
                    {
                        MessageBox.Show("At least one valid operation code is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Reformat to ensure proper comma+space formatting
                    RowData["Operation"] = string.Join(", ", validOperations);
                }

                // Comments field is optional in OPC Config, so no validation needed
                // (No validation block for Comments)
            }

            // Remove confirmation here, just close and let MainForm handle confirmation
            DialogResult = DialogResult.OK;
            Close();
        }


        private bool HasValidationErrors()
        {
            return errorLabels.Values.Any(label => label.Visible);
        }

        private bool ValidateOPCConfigField(string columnName, string value, out string errorMessage)
        {
            errorMessage = "";
            bool hasError = false;
            value = value?.Trim() ?? "";

            // Handle case variations for column names
            switch (columnName.ToUpper())
            {
                case "FE/BE":
                    if (string.IsNullOrEmpty(value) || (value != "FE" && value != "BE"))
                    {
                        errorMessage = "Must select 'FE' or 'BE'";
                        hasError = true;
                    }
                    break;

                case "LAYER":
                case "LAYERS":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "Layer cannot be empty";
                        hasError = true;
                    }
                    else if (value.Length != 3)
                    {
                        errorMessage = "Must be exactly 3 characters";
                        hasError = true;
                    }
                    else if (!System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Za-z0-9]{3}$"))
                    {
                        errorMessage = "Only letters and numbers allowed";
                        hasError = true;
                    }
                    break;

                case "OPC":
                case "OPC REV":
                case "OPC_REV":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "OPC Rev cannot be empty";
                        hasError = true;
                    }
                    else if (!System.Text.RegularExpressions.Regex.IsMatch(value, "^\\d\\.\\d\\.\\d\\.\\d$"))
                    {
                        errorMessage = "OPC Rev must be exactly 7 characters in the format 0.0.0.0 (digits and periods only). Example: 1.2.3.4";
                        hasError = true;
                    }
                    break;

                case "OPERATION":
                case "OPERATIONS":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "Operation cannot be empty";
                        hasError = true;
                    }
                    else if (System.Text.RegularExpressions.Regex.IsMatch(value, "[^0-9,\\s]"))
                    {
                        errorMessage = "Only digits, commas, and spaces allowed";
                        hasError = true;
                    }
                    else
                    {
                        // Be very lenient during real-time typing - only validate completed parts
                        var operations = value.Split(',');
                        
                        for (int i = 0; i < operations.Length; i++)
                        {
                            var opTrimmed = operations[i].Trim();
                            
                            // Skip empty parts (like when user types "123456," and there's an empty part after comma)
                            if (string.IsNullOrEmpty(opTrimmed))
                                continue;
                            
                            // If this is the last part and the original value doesn't end with comma,
                            // or if this is not the last part, validate it
                            bool isLastPart = (i == operations.Length - 1);
                            bool shouldValidate = !isLastPart || !value.EndsWith(",");
                            
                            if (shouldValidate)
                            {
                                if (opTrimmed.Length != 6 || !System.Text.RegularExpressions.Regex.IsMatch(opTrimmed, "^\\d{6}$"))
                                {
                                    errorMessage = "Each operation must be exactly 6 digits";
                                    hasError = true;
                                    break;
                                }
                            }
                            else
                            {
                                // For the last part when user is still typing (ends with comma)
                                // Just check if it's numeric if not empty
                                if (opTrimmed.Length > 0)
                                {
                                    if (!System.Text.RegularExpressions.Regex.IsMatch(opTrimmed, "^\\d+$"))
                                    {
                                        errorMessage = "Each operation must be exactly 6 digits";
                                        hasError = true;
                                        break;
                                    }
                                    else if (opTrimmed.Length > 6)
                                    {
                                        errorMessage = "Each operation must be exactly 6 digits";
                                        hasError = true;
                                        break;
                                    }
                                }
                            }
                        }
                        
                        // Show helpful message if user is in the middle of typing
                        if (!hasError && value.EndsWith(","))
                        {
                            errorMessage = "Type next 6-digit operation...";
                            // This is informational, not an error
                        }
                    }
                    break;

                case "PROCESS":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "PROCESS cannot be empty";
                        hasError = true;
                    }
                    else if (value.Length != 4)
                    {
                        errorMessage = "PROCESS must be exactly 4 characters";
                        hasError = true;
                    }
                    else if (!System.Text.RegularExpressions.Regex.IsMatch(value, "^[0-9]{4}$"))
                    {
                        errorMessage = "PROCESS must be 4 numeric characters only";
                        hasError = true;
                    }
                    break;

                default:
                    // For any other fields in OPC Config, check if they should be validated
                    if (columnName.Equals("userID", StringComparison.OrdinalIgnoreCase) ||
                        columnName.Equals("timestamp", StringComparison.OrdinalIgnoreCase) ||
                        columnName.Equals("KEY", StringComparison.OrdinalIgnoreCase))
                    {
                        // These fields are auto-filled, don't validate
                        break;
                    }
                    else
                    {
                        // All other fields must not be empty
                        if (string.IsNullOrEmpty(value))
                        {
                            errorMessage = $"{columnName} cannot be empty";
                            hasError = true;
                        }
                    }
                    break;
            }

            return hasError;
        }

        private bool ValidateNPILotConfigField(string columnName, string value, out string errorMessage)
        {
            errorMessage = "";
            bool hasError = false;
            value = value?.Trim() ?? "";

            switch (columnName)
            {
                case "LOT":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "LOT cannot be empty";
                        hasError = true;
                    }
                    else if (value.Length < 7)
                    {
                        errorMessage = "LOT must be at least 7 characters";
                        hasError = true;
                    }
                    break;

                case "LOT_TYPE":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "LOT_TYPE cannot be empty";
                        hasError = true;
                    }
                    // Allow any non-empty text for LOT_TYPE
                    break;

                case "RETICLE_PRODUCT":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "RETICLE_PRODUCT cannot be empty";
                        hasError = true;
                    }
                    else if (value.Length != 6 || !System.Text.RegularExpressions.Regex.IsMatch(value, @"^[a-zA-Z0-9]{6}$"))
                    {
                        errorMessage = "RETICLE_PRODUCT must be exactly 6 alphanumeric characters";
                        hasError = true;
                    }
                    break;

                case "ROUTE_TYPE":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "ROUTE_TYPE cannot be empty";
                        hasError = true;
                    }
                    // Allow any non-empty text for ROUTE_TYPE
                    break;

                case "PROD_ORDER":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "PROD_ORDER cannot be empty";
                        hasError = true;
                    }
                    else if (!int.TryParse(value, out int prodOrder) || prodOrder <= 0)
                    {
                        errorMessage = "PROD_ORDER must be a positive number";
                        hasError = true;
                    }
                    break;

                case "LOT_ORDER":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "LOT_ORDER cannot be empty";
                        hasError = true;
                    }
                    else if (!int.TryParse(value, out int order) || order <= 0)
                    {
                        errorMessage = "LOT_ORDER must be a positive number";
                        hasError = true;
                    }
                    break;
            }

            return hasError;
        }

        private bool ValidateRunPlanField(string columnName, string value, out string errorMessage)
        {
            errorMessage = "";
            bool hasError = false;
            value = value?.Trim() ?? "";

            switch (columnName)
            {
                case "PROCESS":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "PROCESS cannot be empty";
                        hasError = true;
                    }
                    else if (value.Length != 4)
                    {
                        errorMessage = "PROCESS must be exactly 4 characters";
                        hasError = true;
                    }
                    else if (!System.Text.RegularExpressions.Regex.IsMatch(value, "^[0-9]{4}$"))
                    {
                        errorMessage = "PROCESS must be 4 numeric characters only";
                        hasError = true;
                    }
                    break;

                case "RETICLE_PRODUCT":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "RETICLE_PRODUCT cannot be empty";
                        hasError = true;
                    }
                    else if (value.Length != 6 || !System.Text.RegularExpressions.Regex.IsMatch(value, @"^[a-zA-Z0-9]{6}$"))
                    {
                        errorMessage = "RETICLE_PRODUCT must be exactly 6 alphanumeric characters";
                        hasError = true;
                    }
                    break;

                case "PRODUCT":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "PRODUCT cannot be empty";
                        hasError = true;
                    }
                    break;

                case "ROUTE_TYPE":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "ROUTE_TYPE cannot be empty";
                        hasError = true;
                    }
                    // Allow any non-empty text for ROUTE_TYPE
                    break;

                case "LAYER_GROUP":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "LAYER_GROUP cannot be empty";
                        hasError = true;
                    }
                    break;

                case "LAYER":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "LAYER cannot be empty";
                        hasError = true;
                    }
                    else if (value.Length != 3)
                    {
                        errorMessage = "LAYER must be exactly 3 characters";
                        hasError = true;
                    }
                    else if (!System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Za-z0-9]{3}$"))
                    {
                        errorMessage = "LAYER: Only letters and numbers allowed";
                        hasError = true;
                    }
                    break;

                case "RUN_PLAN":
                    if (string.IsNullOrEmpty(value))
                    {
                        errorMessage = "RUN_PLAN cannot be empty";
                        hasError = true;
                    }
                    break;
            }

            return hasError;
        }

        private string FormatOperationString(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;
            
            // Split by commas, trim each part, and rejoin with ", "
            var operations = value.Split(',');
            var formattedOps = new List<string>();
            
            foreach (var op in operations)
            {
                var trimmed = op.Trim();
                if (!string.IsNullOrEmpty(trimmed))
                {
                    formattedOps.Add(trimmed);
                }
            }
            
            return string.Join(", ", formattedOps);
        }
    }
}
