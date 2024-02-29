using CsvHelper;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Trilogen.Helpers;

namespace Trilogen
{
    public partial class ImportForm : Form
    {
        SharepointManager spManager;

        string[] csvHeaders = null;

        List<string[]> csvRecords = null;

        Thread importThread = null;

        StatusForm statusForm = null;

        // delegates
        private delegate void ResetListNameCallback();

        private delegate void EnableBrowseCallback();

        private delegate void EnableImportCallback();

        public ImportForm()
        {
            // Handle the ApplicationExit event to know when the application is exiting.
            Application.ApplicationExit += new EventHandler(OnApplicationExit);

            // init
            InitializeComponent();

            // lock other groups
            gbValidate.Enabled = false;
            gbImport.Enabled = false;
            btnValidateFileCSV.Enabled = false;

        }

        public bool Connect(string siteUrl, string user, string password)
        {
            bool connected = false;

            try
            {
                // sp manager
                spManager = new SharepointManager(siteUrl, user, password);

                // events
                spManager.ImportCompleted += new ImportHandler(importCompleted);
                spManager.ImportFailed += new ImportHandler(importFailed);
                spManager.ImportUpdate += new ImportHandler(importUpdated);

                // sp lists
                var spLists = spManager.GetLists();

                // lists exists
                if (spLists != null)
                {
                    // get num items
                    int numListItems = spLists.Count;

                    // has items
                    if (numListItems > 0)
                    {
                        for (int i = 0; i < numListItems; i++)
                        {
                            // list item ref
                            var listItem = spLists[i];

                            // new item
                            ComboBoxItem item = new ComboBoxItem
                            {
                                Text = listItem.Title,
                                Value = listItem.Id,
                                Fields = listItem.Fields
                            };

                            // add to list
                            cbListname.Items.Add(item);
                        }
                    }

                    // success
                    connected = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error connecting to SharePoint site", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connected = false;
            }

            return connected;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            DialogResult openFileResult = openFileDialog.ShowDialog();
        }

        private void openFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            txtImportFilename.Text = openFileDialog.FileName;

            // import disabled until file is validated
            gbImport.Enabled = false;
            btnValidateFileCSV.Enabled = true;

            // reset label
            lblNumRecords.Text = "0 records";
        }


        private string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            if (cell.CellValue != null)
            {
                string value = cell.CellValue.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    // in older version e.g. 2.0, you can use GetItem instead of ElementAt
                    return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.ElementAt(int.Parse(value)).InnerText;
                }
                else
                {
                    return value;
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// Чтение Excel
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        private bool ValidateExcelFile(string fileName)
        {
            // clear existing records
            dgvMappings.Rows.Clear();

            csvHeaders = new string[0];
            csvRecords = new List<string[]>();

            // file is not empty
            if (!string.IsNullOrEmpty(fileName) && (File.Exists(fileName)))
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    var rows = sheetData.Elements<Row>();
                    int rowCount = rows.Count();

                    if (rowCount < 2)
                    {
                        return false;
                    }

                    var first = true;
                    foreach (var row in rows)
                    {
                        if (first)
                        {
                            first = false;
                            Row headRow = row;

                            if (headRow == null)
                            {
                                return false;
                            }

                            List<string> headers = new List<string>();
                            var cells = headRow.Elements<Cell>();
                            foreach (Cell c in cells)
                            {
                                headers.Add(GetCellValue(spreadsheetDocument, c));
                            }

                            if (headers.Count == 0)
                            {
                                return false;
                            }

                            csvHeaders = headers.ToArray();

                            // loop through headers
                            foreach (var header in csvHeaders)
                            {
                                if (!string.IsNullOrEmpty(header))
                                {
                                    // add datagrid row
                                    DataGridViewRow newRow = new DataGridViewRow();
                                    // checkbox
                                    DataGridViewCheckBoxCell cbCell = new DataGridViewCheckBoxCell();
                                    // text box
                                    DataGridViewTextBoxCell txtHeader = new DataGridViewTextBoxCell
                                    {
                                        Value = header
                                    };
                                    DataGridViewComboBoxCell cbMappings = new DataGridViewComboBoxCell();
                                    // add cells
                                    newRow.Cells.Add(cbCell);
                                    newRow.Cells.Add(txtHeader);
                                    newRow.Cells.Add(cbMappings);
                                    // add row
                                    dgvMappings.Rows.Add(newRow);
                                    // set read only after row is added
                                    cbMappings.ReadOnly = true;
                                }
                            }
                        }
                        else
                        {
                            List<string> cellValues = new List<string>();
                            var cells = row.Elements<Cell>();
                            foreach (Cell c in cells)
                            {
                                cellValues.Add(GetCellValue(spreadsheetDocument, c));
                            }
                            string[] record = cellValues.ToArray();
                            csvRecords.Add(record);
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Чтение CSV
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        private bool ValidateCSVFile(string filename)
        {
            // file is not empty
            if (!string.IsNullOrEmpty(filename) && (File.Exists(filename)))
            {

                // file stream
                using (StreamReader csvStreamReader = File.OpenText(filename))
                {
                    // csv parser
                    CsvParser csvParser = new CsvParser(csvStreamReader);

                    // read headers
                    csvHeaders = csvParser.Read();

                    if (csvHeaders == null)
                        return false;

                    // clear existing records
                    dgvMappings.Rows.Clear();

                    // loop through headers
                    foreach (var header in csvHeaders)
                    {
                        if (!string.IsNullOrEmpty(header))
                        {
                            // add datagrid row
                            DataGridViewRow newRow = new DataGridViewRow();

                            // checkbox
                            DataGridViewCheckBoxCell cbCell = new DataGridViewCheckBoxCell();

                            // text box
                            DataGridViewTextBoxCell txtHeader = new DataGridViewTextBoxCell
                            {
                                Value = header
                            };

                            DataGridViewComboBoxCell cbMappings = new DataGridViewComboBoxCell();

                            // add cells
                            newRow.Cells.Add(cbCell);
                            newRow.Cells.Add(txtHeader);
                            newRow.Cells.Add(cbMappings);

                            // add row
                            dgvMappings.Rows.Add(newRow);

                            // set read only after row is added
                            cbMappings.ReadOnly = true;
                        }
                    }

                    // new records obj
                    csvRecords = new List<string[]>();

                    // loop records
                    while (true)
                    {
                        // record
                        string[] record = csvParser.Read();

                        // record does not exist
                        if (record == null)
                        {
                            break;
                        }

                        csvRecords.Add(record);
                    }

                    return true;
                }

            }

            return false;
        }


        private void btnValidateFileExcel_Click(object sender, EventArgs e)
        {
            // filename
            string csvFile = txtImportFilename.Text;

            // file path set
            if (!string.IsNullOrEmpty(csvFile) && csvFile.Length > 0 && (File.Exists(csvFile)))
            {
                // get csv headers
                bool testExcel = ValidateExcelFile(csvFile);

                // valid file
                if (!testExcel)
                {
                    MessageBox.Show("Unable to validate Excel file", "Error parsing Excel file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // has records
                if (csvRecords != null && csvRecords.Count > 0)
                {
                    // enable sharepoint import
                    gbImport.Enabled = true;
                    lblNumRecords.Text = csvRecords.Count + " records";
                }
                else
                {
                    MessageBox.Show("Excel file has no records to import!", "No records to import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnValidateFile_Click(object sender, EventArgs e)
        {
            // filename
            string csvFile = txtImportFilename.Text;

            // file path set
            if (!string.IsNullOrEmpty(csvFile) && csvFile.Length > 0 && (File.Exists(csvFile)))
            {
                // get csv headers
                bool testCSV = ValidateCSVFile(csvFile);

                // valid file
                if (!testCSV)
                {
                    MessageBox.Show("Unable to validate CSV file", "Error parsing CSV file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // has records
                if (csvRecords != null && csvRecords.Count > 0)
                {
                    // enable sharepoint import
                    gbImport.Enabled = true;
                    lblNumRecords.Text = csvRecords.Count + " records";
                }
                else
                {
                    MessageBox.Show("CSV file has no records to import!", "No records to import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            string siteUrl = txtSiteUrl.Text;
            string user = txtUser.Text;
            string password = txtPassword.Text;

            bool result = Connect(siteUrl, user, password);

            if (result)
            {
                // button flags
                btnConnect.Enabled = false;
                btnDisconnect.Enabled = true;

                // enable csv processing
                gbValidate.Enabled = true;
            }
        }

        private void UpdateMappings()
        {
            // we have a selection
            if (cbListname.SelectedIndex > -1)
            {
                // populate columns
                ComboBoxItem selectedItem = (ComboBoxItem)cbListname.Items[cbListname.SelectedIndex];

                // sp fields
                var fields = selectedItem.Fields;

                // update datagrid
                for (int i = 0; i < dgvMappings.Rows.Count; i++)
                {
                    // column
                    string columnName = Convert.ToString(dgvMappings.Rows[i].Cells[1].Value);

                    // get combobox from each row
                    DataGridViewComboBoxCell cbCurrentMapping = (DataGridViewComboBoxCell)dgvMappings.Rows[i].Cells[2];

                    // clear current mapping items
                    cbCurrentMapping.Value = null;

                    // enable combo box for mappings
                    cbCurrentMapping.ReadOnly = false;

                    // clear items in box
                    cbCurrentMapping.Items.Clear();

                    int fieldIndex = 0;
                    string matchedProperty = string.Empty;

                    // add list fields to drop list items
                    foreach (var field in fields)
                    {
                        // not read only
                        if (!field.ReadOnlyField)
                        {
                            cbCurrentMapping.Items.Add(field.InternalName);

                            if (columnName == field.InternalName)
                            {
                                matchedProperty = field.InternalName;
                            }
                        }

                        fieldIndex++;
                    }

                    // found a field match
                    if (!string.IsNullOrEmpty(matchedProperty))
                    {
                        cbCurrentMapping.Value = matchedProperty;
                    }
                }

                // refresh grid
                dgvMappings.Update();
                dgvMappings.Refresh();
            }
        }

        private void cbListname_SelectedIndexChanged(object sender, EventArgs e)
        {
            // update mappings
            UpdateMappings();
        }

        bool HasSelectedMappingColumn()
        {
            for (int i = 0; i < dgvMappings.Rows.Count; i++)
            {
                DataGridViewRow row = dgvMappings.Rows[i];
                DataGridViewCheckBoxCell cbCell = (DataGridViewCheckBoxCell)row.Cells[0];
                bool isChecked = (Convert.ToBoolean(cbCell.Value));

                if (isChecked)
                    return true;
            }

            return false;
        }
               
        private void importUpdated(ImportEvent e)
        {
            if (e != null)
            {
                // round progress
                int progressValue = (int)Math.Round(e.ImportProgressValue, 0);

                if (statusForm != null)
                {
                    statusForm.SetProgressValue(progressValue);
                    statusForm.SetStatusText("Imported " + e.RecordsImported + " of " + e.TotalRecords + " records");
                }
            }
        }

        private void importCompleted(ImportEvent e)
        {
            // close status window
            if (statusForm != null)
            {
                statusForm.CloseWindow();
            }

            // success message
            MessageBox.Show("Successfully imported " + e.RecordsImported + " records into SharePoint!", "Import completed", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // enable browse
            enableBrowse();

            // enable import
            enableImport();
        }

        private void importFailed(ImportEvent e)
        {
            // close status window
            if (statusForm != null)
            {
                statusForm.CloseWindow();
            }

            // error msg
            MessageBox.Show("Error importing CSV file into SharePoint.\n\n" + e.ErrorMessage, "Error importing CSV", MessageBoxButtons.OK, MessageBoxIcon.Error);

            // enable browse
            enableBrowse();

            // enable import
            enableImport();
        }

        private void importCancelled()
        {
            // abort import thread
            if (importThread != null)
            {
                if (importThread.IsAlive)
                {
                    importThread.Abort();
                    importThread = null;
                }
            }

            // enable browse
            enableBrowse();

            // enable import
            enableImport();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            // clear status window
            if (statusForm != null)
            {
                statusForm.CloseWindow();
                statusForm = null;
            }

            // validate column selection
            if (!HasSelectedMappingColumn())
            {
                MessageBox.Show("You must select at least 1 column to import", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
            }

            // we are connected to sp site
            if (spManager != null)
            {
                // vars
                Guid listId = Guid.Empty;

                // has list been selected
                if (cbListname.SelectedIndex > -1)
                {
                    ComboBoxItem selectedItem = (ComboBoxItem)cbListname.Items[cbListname.SelectedIndex];

                    // select item exists
                    if (selectedItem != null)
                    {
                        listId = (Guid)selectedItem.Value;
                    }
                }
                else
                {
                    MessageBox.Show("Select list", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // selected list
                if (listId != Guid.Empty)
                {
                    // start import thread
                    importThread = new Thread(
                        () => spManager.Import(listId, csvHeaders, csvRecords, dgvMappings)
                    );

                    // start thread
                    importThread.Start();

                    // show new status window
                    statusForm = new StatusForm();
                    statusForm.SetTitle("Importing CSV records");
                    statusForm.SetStatusText("Importing " + csvRecords.Count + " records");
                    statusForm.Show();

                    // bind to cancel event
                    statusForm.CancelClicked += importCancelled;

                    // disable buttons
                    btnBrowse.Enabled = false;
                    btnValidateFileCSV.Enabled = false;
                    btnImport.Enabled = false;
                }
            }
        }

        private void cbSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            // item is checked
            if (cbSelectAll.Checked)
            {
                // loop through columns
                for (int i = 0; i < dgvMappings.Rows.Count; i++)
                {
                    // select item
                    DataGridViewCheckBoxCell cbCell = (DataGridViewCheckBoxCell)dgvMappings.Rows[i].Cells[0];

                    cbCell.Value = true;
                }
            }
            else
            {
                // loop through columns
                for (int i = 0; i < dgvMappings.Rows.Count; i++)
                {
                    // select item
                    DataGridViewCheckBoxCell cbCell = (DataGridViewCheckBoxCell)dgvMappings.Rows[i].Cells[0];

                    cbCell.Value = false;
                }
            }
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {

        }

        private void cbListname_SelectedValueChanged(object sender, EventArgs e)
        {
            UpdateMappings();
        }

        private bool listNameExists(string listName)
        {
            foreach (var comboItem in cbListname.Items)
            {
                ComboBoxItem item = (ComboBoxItem)comboItem;

                if (item.Text == listName)
                {
                    return true;
                }
            }

            return false;
        }

        private void cbListname_TextUpdate(object sender, EventArgs e)
        {
            // check if list exists in box, else new list
            bool listExists = listNameExists(cbListname.Text);

            // list not found
            if (!listExists)
            {
                // clear mappings
                foreach (DataGridViewRow row in dgvMappings.Rows)
                {
                    // data type required
                    DataGridViewComboBoxCell cbType = (DataGridViewComboBoxCell)row.Cells[3];
                    cbType.ReadOnly = false;

                    // mapping cell
                    DataGridViewComboBoxCell cbCurrentMapping = (DataGridViewComboBoxCell)row.Cells[2];

                    // unselect any selected mappings
                    cbCurrentMapping.Value = null;
                    cbCurrentMapping.ReadOnly = true;

                    cbCurrentMapping.Items.Clear();
                }
            }
        }

        private void dgvMappings_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvMappings != null && dgvMappings.Rows.Count > 0)
            {
                // mapping select box changed value
                if (e.ColumnIndex == 2)
                {
                    // if row not selected, check box
                    DataGridViewCheckBoxCell cbCell = (DataGridViewCheckBoxCell)dgvMappings.Rows[e.RowIndex].Cells[0];

                    // check box
                    cbCell.Value = true;
                }
            }
        }

        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            // flags
            btnConnect.Enabled = true;
            btnDisconnect.Enabled = false;

            // reset validate group box
            gbValidate.Enabled = false;
            txtImportFilename.Text = string.Empty;
            lblNumRecords.Text = "0 records";

            // reset import group box
            dgvMappings.Rows.Clear(); // clear rows in datagridview
            cbListname.Items.Clear(); // clear sharepoint list drop down
            cbListname.Text = string.Empty;
            gbImport.Enabled = false;
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            // reset mapping grid
            foreach (DataGridViewRow row in dgvMappings.Rows)
            {
                // checkbox
                DataGridViewCheckBoxCell cbSelected = (DataGridViewCheckBoxCell)row.Cells[0];

                if (cbSelected != null)
                {
                    bool isSelected = Convert.ToBoolean(cbSelected.Value);

                    if (isSelected)
                        cbSelected.Value = false;
                }

                // mapping drop down
                DataGridViewComboBoxCell cbMapping = (DataGridViewComboBoxCell)row.Cells[2];

                if (cbMapping != null)
                    cbMapping.Value = null;
            }

            // refresh mappings
            dgvMappings.Update();
        }

        private void ImportForm_FormClosed(Object sender, FormClosedEventArgs e)
        {
            // abort import thread if one exists
            if (importThread != null)
            {
                if (importThread.IsAlive)
                {
                    importThread.Abort();
                }
            }

            // close status window
            if (statusForm != null)
            {
                // close window
                statusForm.CloseWindow();
                statusForm = null;
            }
        }

        // thread safe - enable browse
        private void enableBrowse()
        {
            if (btnBrowse.InvokeRequired)
            {
                EnableBrowseCallback enableBrowseCallback = new EnableBrowseCallback(enableBrowse);
                Invoke(enableBrowseCallback);
            }
            else
            {
                // enable button
                btnBrowse.Enabled = true;
            }
        }

        // thead safe - reset list  name
        private void resetListName()
        {
            if (cbListname.InvokeRequired)
            {
                ResetListNameCallback resetListNameCallback = new ResetListNameCallback(resetListName);
                Invoke(resetListNameCallback);
            }
            else
            {
                // reset selected items
                cbListname.Text = string.Empty;
                cbListname.SelectedIndex = -1;
            }
        }

        private void enableImport()
        {
            if (btnImport.InvokeRequired)
            {
                EnableImportCallback enableImportCallback = new EnableImportCallback(enableImport);
                Invoke(enableImportCallback);
            }
            else
            {
                btnImport.Enabled = true;
            }
        }

        /// <summary>
        /// Test
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Test_Click(object sender, EventArgs e)
        {
            ComboBoxItem selectedItem = (ComboBoxItem)cbListname.Items[cbListname.SelectedIndex];

            var listName = selectedItem.Text;
            var listId = (Guid)selectedItem.Value;

            spManager.Read(listId);

        }

        /// <summary>
        /// ShowHelpForm
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ShowHelpForm_Click(object sender, EventArgs e)
        {
            var help = new Forms.Help();
            help.ShowDialog();

        }

    }
}
