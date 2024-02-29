namespace Trilogen
{
    partial class ImportForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ImportForm));
            this.gbLogin = new System.Windows.Forms.GroupBox();
            this.btnDisconnect = new System.Windows.Forms.Button();
            this.btnConnect = new System.Windows.Forms.Button();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.txtSiteUrl = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblSiteUrl = new System.Windows.Forms.Label();
            this.btn_Help = new System.Windows.Forms.Button();
            this.gbValidate = new System.Windows.Forms.GroupBox();
            this.btnValidateFileExcel = new System.Windows.Forms.Button();
            this.lblNumRecords = new System.Windows.Forms.Label();
            this.lblRecords = new System.Windows.Forms.Label();
            this.btnValidateFileCSV = new System.Windows.Forms.Button();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.txtImportFilename = new System.Windows.Forms.TextBox();
            this.lblFilename = new System.Windows.Forms.Label();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.gbImport = new System.Windows.Forms.GroupBox();
            this.btn_TEST = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.dgvMappings = new System.Windows.Forms.DataGridView();
            this.Include = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Column = new System.Windows.Forms.DataGridViewLinkColumn();
            this.Mapping = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.cbSelectAll = new System.Windows.Forms.CheckBox();
            this.btnImport = new System.Windows.Forms.Button();
            this.lblMappings = new System.Windows.Forms.Label();
            this.cbListname = new System.Windows.Forms.ComboBox();
            this.lblListname = new System.Windows.Forms.Label();
            this.gbLogin.SuspendLayout();
            this.gbValidate.SuspendLayout();
            this.gbImport.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMappings)).BeginInit();
            this.SuspendLayout();
            // 
            // gbLogin
            // 
            this.gbLogin.Controls.Add(this.btnDisconnect);
            this.gbLogin.Controls.Add(this.btnConnect);
            this.gbLogin.Controls.Add(this.txtPassword);
            this.gbLogin.Controls.Add(this.txtUser);
            this.gbLogin.Controls.Add(this.txtSiteUrl);
            this.gbLogin.Controls.Add(this.label2);
            this.gbLogin.Controls.Add(this.label1);
            this.gbLogin.Controls.Add(this.lblSiteUrl);
            this.gbLogin.Location = new System.Drawing.Point(12, 30);
            this.gbLogin.Name = "gbLogin";
            this.gbLogin.Size = new System.Drawing.Size(1071, 142);
            this.gbLogin.TabIndex = 0;
            this.gbLogin.TabStop = false;
            this.gbLogin.Text = "Sharepoint Details";
            // 
            // btnDisconnect
            // 
            this.btnDisconnect.Enabled = false;
            this.btnDisconnect.Location = new System.Drawing.Point(895, 120);
            this.btnDisconnect.Name = "btnDisconnect";
            this.btnDisconnect.Size = new System.Drawing.Size(89, 23);
            this.btnDisconnect.TabIndex = 6;
            this.btnDisconnect.Text = "Disconnect";
            this.btnDisconnect.UseVisualStyleBackColor = true;
            this.btnDisconnect.Click += new System.EventHandler(this.btnDisconnect_Click);
            // 
            // btnConnect
            // 
            this.btnConnect.Location = new System.Drawing.Point(990, 120);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(75, 23);
            this.btnConnect.TabIndex = 5;
            this.btnConnect.Text = "Connect";
            this.btnConnect.UseVisualStyleBackColor = true;
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // txtPassword
            // 
            this.txtPassword.BackColor = System.Drawing.Color.LightSteelBlue;
            this.txtPassword.Location = new System.Drawing.Point(132, 85);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(306, 27);
            this.txtPassword.TabIndex = 1;
            // 
            // txtUser
            // 
            this.txtUser.BackColor = System.Drawing.Color.LightSteelBlue;
            this.txtUser.Location = new System.Drawing.Point(132, 52);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(306, 27);
            this.txtUser.TabIndex = 1;
            // 
            // txtSiteUrl
            // 
            this.txtSiteUrl.BackColor = System.Drawing.Color.LightSteelBlue;
            this.txtSiteUrl.Location = new System.Drawing.Point(132, 19);
            this.txtSiteUrl.Name = "txtSiteUrl";
            this.txtSiteUrl.Size = new System.Drawing.Size(933, 27);
            this.txtSiteUrl.TabIndex = 1;
            this.txtSiteUrl.Text = "https://testsp2:33782/";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 85);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 21);
            this.label2.TabIndex = 0;
            this.label2.Text = "User password";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 21);
            this.label1.TabIndex = 0;
            this.label1.Text = "User login";
            // 
            // lblSiteUrl
            // 
            this.lblSiteUrl.AutoSize = true;
            this.lblSiteUrl.Location = new System.Drawing.Point(6, 22);
            this.lblSiteUrl.Name = "lblSiteUrl";
            this.lblSiteUrl.Size = new System.Drawing.Size(79, 21);
            this.lblSiteUrl.TabIndex = 0;
            this.lblSiteUrl.Text = "Site URL:";
            // 
            // btn_Help
            // 
            this.btn_Help.Location = new System.Drawing.Point(994, 12);
            this.btn_Help.Name = "btn_Help";
            this.btn_Help.Size = new System.Drawing.Size(89, 23);
            this.btn_Help.TabIndex = 7;
            this.btn_Help.Text = "Help";
            this.btn_Help.UseVisualStyleBackColor = true;
            this.btn_Help.Click += new System.EventHandler(this.ShowHelpForm_Click);
            // 
            // gbValidate
            // 
            this.gbValidate.Controls.Add(this.btnValidateFileExcel);
            this.gbValidate.Controls.Add(this.lblNumRecords);
            this.gbValidate.Controls.Add(this.lblRecords);
            this.gbValidate.Controls.Add(this.btnValidateFileCSV);
            this.gbValidate.Controls.Add(this.btnBrowse);
            this.gbValidate.Controls.Add(this.txtImportFilename);
            this.gbValidate.Controls.Add(this.lblFilename);
            this.gbValidate.Location = new System.Drawing.Point(12, 178);
            this.gbValidate.Name = "gbValidate";
            this.gbValidate.Size = new System.Drawing.Size(1071, 111);
            this.gbValidate.TabIndex = 0;
            this.gbValidate.TabStop = false;
            this.gbValidate.Text = "Prepare for import";
            // 
            // btnValidateFileExcel
            // 
            this.btnValidateFileExcel.Location = new System.Drawing.Point(866, 71);
            this.btnValidateFileExcel.Name = "btnValidateFileExcel";
            this.btnValidateFileExcel.Size = new System.Drawing.Size(186, 23);
            this.btnValidateFileExcel.TabIndex = 9;
            this.btnValidateFileExcel.Text = "Validate Exel file";
            this.btnValidateFileExcel.UseVisualStyleBackColor = true;
            this.btnValidateFileExcel.Click += new System.EventHandler(this.btnValidateFileExcel_Click);
            // 
            // lblNumRecords
            // 
            this.lblNumRecords.AutoSize = true;
            this.lblNumRecords.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNumRecords.Location = new System.Drawing.Point(96, 51);
            this.lblNumRecords.Name = "lblNumRecords";
            this.lblNumRecords.Padding = new System.Windows.Forms.Padding(0, 2, 0, 2);
            this.lblNumRecords.Size = new System.Drawing.Size(81, 27);
            this.lblNumRecords.TabIndex = 0;
            this.lblNumRecords.Text = "0 records";
            this.lblNumRecords.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblRecords
            // 
            this.lblRecords.AutoSize = true;
            this.lblRecords.Location = new System.Drawing.Point(6, 53);
            this.lblRecords.Name = "lblRecords";
            this.lblRecords.Size = new System.Drawing.Size(76, 21);
            this.lblRecords.TabIndex = 0;
            this.lblRecords.Text = "Records:";
            // 
            // btnValidateFileCSV
            // 
            this.btnValidateFileCSV.Location = new System.Drawing.Point(673, 71);
            this.btnValidateFileCSV.Name = "btnValidateFileCSV";
            this.btnValidateFileCSV.Size = new System.Drawing.Size(186, 23);
            this.btnValidateFileCSV.TabIndex = 8;
            this.btnValidateFileCSV.Text = "Validate CSV file";
            this.btnValidateFileCSV.UseVisualStyleBackColor = true;
            this.btnValidateFileCSV.Click += new System.EventHandler(this.btnValidateFile_Click);
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(990, 20);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 7;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // txtImportFilename
            // 
            this.txtImportFilename.BackColor = System.Drawing.Color.LightSteelBlue;
            this.txtImportFilename.Location = new System.Drawing.Point(95, 20);
            this.txtImportFilename.Name = "txtImportFilename";
            this.txtImportFilename.Size = new System.Drawing.Size(880, 27);
            this.txtImportFilename.TabIndex = 6;
            // 
            // lblFilename
            // 
            this.lblFilename.AutoSize = true;
            this.lblFilename.Location = new System.Drawing.Point(6, 23);
            this.lblFilename.Name = "lblFilename";
            this.lblFilename.Size = new System.Drawing.Size(88, 21);
            this.lblFilename.TabIndex = 0;
            this.lblFilename.Text = "File name:";
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "CSV Files|*.csv|xlsx Files|*.xlsx";
            this.openFileDialog.Title = "Import CSV file into Sharepoint";
            this.openFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog_FileOk);
            // 
            // gbImport
            // 
            this.gbImport.Controls.Add(this.btn_TEST);
            this.gbImport.Controls.Add(this.btnReset);
            this.gbImport.Controls.Add(this.dgvMappings);
            this.gbImport.Controls.Add(this.cbSelectAll);
            this.gbImport.Controls.Add(this.btnImport);
            this.gbImport.Controls.Add(this.lblMappings);
            this.gbImport.Controls.Add(this.cbListname);
            this.gbImport.Controls.Add(this.lblListname);
            this.gbImport.Location = new System.Drawing.Point(12, 295);
            this.gbImport.Name = "gbImport";
            this.gbImport.Size = new System.Drawing.Size(1071, 235);
            this.gbImport.TabIndex = 0;
            this.gbImport.TabStop = false;
            this.gbImport.Text = "Import into Sharepoint";
            // 
            // btn_TEST
            // 
            this.btn_TEST.Location = new System.Drawing.Point(95, 206);
            this.btn_TEST.Name = "btn_TEST";
            this.btn_TEST.Size = new System.Drawing.Size(75, 23);
            this.btn_TEST.TabIndex = 14;
            this.btn_TEST.Text = "test";
            this.btn_TEST.UseVisualStyleBackColor = true;
            this.btn_TEST.Visible = false;
            this.btn_TEST.Click += new System.EventHandler(this.Test_Click);
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(885, 206);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(75, 23);
            this.btnReset.TabIndex = 13;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // dgvMappings
            // 
            this.dgvMappings.AllowUserToAddRows = false;
            this.dgvMappings.AllowUserToDeleteRows = false;
            this.dgvMappings.AllowUserToResizeRows = false;
            this.dgvMappings.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader;
            this.dgvMappings.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvMappings.BackgroundColor = System.Drawing.Color.LightSteelBlue;
            this.dgvMappings.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dgvMappings.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.Raised;
            this.dgvMappings.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.dgvMappings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvMappings.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Include,
            this.Column,
            this.Mapping});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.LightGray;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.GradientActiveCaption;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvMappings.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgvMappings.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgvMappings.Location = new System.Drawing.Point(95, 46);
            this.dgvMappings.MultiSelect = false;
            this.dgvMappings.Name = "dgvMappings";
            this.dgvMappings.RowHeadersVisible = false;
            this.dgvMappings.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dgvMappings.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvMappings.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvMappings.ShowCellToolTips = false;
            this.dgvMappings.ShowEditingIcon = false;
            this.dgvMappings.ShowRowErrors = false;
            this.dgvMappings.Size = new System.Drawing.Size(970, 130);
            this.dgvMappings.TabIndex = 10;
            this.dgvMappings.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvMappings_CellValueChanged);
            // 
            // Include
            // 
            this.Include.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.Include.HeaderText = "*";
            this.Include.Name = "Include";
            this.Include.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Include.ToolTipText = "Import this column";
            this.Include.Width = 23;
            // 
            // Column
            // 
            this.Column.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column.HeaderText = "Column";
            this.Column.Name = "Column";
            this.Column.ReadOnly = true;
            // 
            // Mapping
            // 
            this.Mapping.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.Mapping.HeaderText = "Mapping";
            this.Mapping.Name = "Mapping";
            this.Mapping.Width = 76;
            // 
            // cbSelectAll
            // 
            this.cbSelectAll.AutoSize = true;
            this.cbSelectAll.Location = new System.Drawing.Point(966, 182);
            this.cbSelectAll.Name = "cbSelectAll";
            this.cbSelectAll.Size = new System.Drawing.Size(99, 25);
            this.cbSelectAll.TabIndex = 11;
            this.cbSelectAll.Text = "Select all";
            this.cbSelectAll.UseVisualStyleBackColor = true;
            this.cbSelectAll.CheckedChanged += new System.EventHandler(this.cbSelectAll_CheckedChanged);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(966, 206);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(86, 23);
            this.btnImport.TabIndex = 12;
            this.btnImport.Text = "Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // lblMappings
            // 
            this.lblMappings.AutoSize = true;
            this.lblMappings.Location = new System.Drawing.Point(6, 46);
            this.lblMappings.Name = "lblMappings";
            this.lblMappings.Size = new System.Drawing.Size(78, 21);
            this.lblMappings.TabIndex = 0;
            this.lblMappings.Text = "Mapping:";
            // 
            // cbListname
            // 
            this.cbListname.BackColor = System.Drawing.Color.LightSteelBlue;
            this.cbListname.FormattingEnabled = true;
            this.cbListname.Location = new System.Drawing.Point(95, 16);
            this.cbListname.Name = "cbListname";
            this.cbListname.Size = new System.Drawing.Size(970, 27);
            this.cbListname.TabIndex = 0;
            this.cbListname.SelectedIndexChanged += new System.EventHandler(this.cbListname_SelectedIndexChanged);
            this.cbListname.TextUpdate += new System.EventHandler(this.cbListname_TextUpdate);
            this.cbListname.SelectedValueChanged += new System.EventHandler(this.cbListname_SelectedValueChanged);
            // 
            // lblListname
            // 
            this.lblListname.AutoSize = true;
            this.lblListname.Location = new System.Drawing.Point(6, 19);
            this.lblListname.Name = "lblListname";
            this.lblListname.Size = new System.Drawing.Size(42, 21);
            this.lblListname.TabIndex = 0;
            this.lblListname.Text = "List:";
            // 
            // ImportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1095, 541);
            this.Controls.Add(this.btn_Help);
            this.Controls.Add(this.gbImport);
            this.Controls.Add(this.gbValidate);
            this.Controls.Add(this.gbLogin);
            this.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "ImportForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import CSV into SharePoint";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ImportForm_FormClosed);
            this.gbLogin.ResumeLayout(false);
            this.gbLogin.PerformLayout();
            this.gbValidate.ResumeLayout(false);
            this.gbValidate.PerformLayout();
            this.gbImport.ResumeLayout(false);
            this.gbImport.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMappings)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gbLogin;
        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.TextBox txtSiteUrl;
        private System.Windows.Forms.Label lblSiteUrl;
        private System.Windows.Forms.GroupBox gbValidate;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtImportFilename;
        private System.Windows.Forms.Label lblFilename;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button btnValidateFileCSV;
        private System.Windows.Forms.GroupBox gbImport;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Label lblMappings;
        private System.Windows.Forms.ComboBox cbListname;
        private System.Windows.Forms.Label lblListname;
        private System.Windows.Forms.CheckBox cbSelectAll;
        private System.Windows.Forms.Label lblNumRecords;
        private System.Windows.Forms.Label lblRecords;
        private System.Windows.Forms.DataGridView dgvMappings;
        private System.Windows.Forms.Button btnDisconnect;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Include;
        private System.Windows.Forms.DataGridViewLinkColumn Column;
        private System.Windows.Forms.DataGridViewComboBoxColumn Mapping;
        private System.Windows.Forms.Button btn_TEST;
        private System.Windows.Forms.Button btn_Help;
        private System.Windows.Forms.TextBox txtUser;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnValidateFileExcel;
    }
}