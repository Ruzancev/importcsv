using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Trilogen
{
    public delegate void ImportHandler(ImportEvent e);

    public class ImportEvent : EventArgs
    {
        public bool Success = false;

        public string ErrorMessage = string.Empty;

        public int RecordsImported = 0;

        public int TotalRecords = 0;

        public float ImportProgressValue = 0f;
    }

    public class SharepointManager
    {
        // members
        private ClientContext _context;

        private Web _web;

        private int maxListItemsPerQuery = 40;

        // events
        public event ImportHandler ImportCompleted;

        public event ImportHandler ImportFailed;

        public event ImportHandler ImportUpdate;

        public SharepointManager(string siteUrl)
        {
            // new context
            _context = new ClientContext(siteUrl);

            // web ref
            _web = _context.Web;

            // load context
            _context.Load(_web);
            _context.ExecuteQuery();
        }

        private void importCompleted(ImportEvent e)
        {
            ImportCompleted(e);
        }

        private void importFailed(ImportEvent e)
        {
            ImportFailed(e);
        }

        private void importUpdate(ImportEvent e)
        {
            ImportUpdate(e);
        }

        public List<SharepointList> GetLists()
        {
            List<SharepointList> spLists = new List<SharepointList>();

            _context.Load(_web.Lists,
             lists => lists.Include(list => list.Title, // For each list, retrieve Title and Id. 
                                    list => list.Id,
                                    list => list.Fields));

            // Execute query. 
            _context.ExecuteQuery();

            // Enumerate the web.Lists. 
            foreach (List list in _web.Lists)
            {
                spLists.Add(
                    new SharepointList
                    {
                        Id = list.Id,
                        Title = list.Title,
                        Fields = list.Fields
                    }
                );
            }

            return spLists;
        }

        private string GetSharepointMappingFromCSVColumn(string header, DataGridView dgvMapping)
        {
            for (int i = 0; i < dgvMapping.Rows.Count; i++)
            {
                DataGridViewRow row = dgvMapping.Rows[i];
                string headerCell = Convert.ToString(row.Cells[1].Value);

                if (string.IsNullOrEmpty(headerCell))
                    continue;

                if (headerCell == header)
                {
                    DataGridViewComboBoxCell cbCell = (DataGridViewComboBoxCell)row.Cells[2];

                    return Convert.ToString(cbCell.Value);
                }
            }

            return string.Empty;
        }

        // import to existing list
        public bool Import(Guid listId, string[] headers, List<string[]> records, DataGridView mappings)
        {
            // validation
            if (listId == Guid.Empty)
                return false;

            if (records == null)
                return false;

            if (mappings == null)
                return false;

            try
            {
                // get sp list
                List selectedList = _web.Lists.GetById(listId);

                // try load list
                try
                {
                    // get list including fields
                    _context.Load(selectedList, l => l.Fields);
                    _context.ExecuteQuery();

                }
                catch (Exception e)
                {
                    // send failed event
                    importFailed(new ImportEvent
                    {
                        ErrorMessage = e.Message,
                        Success = false
                    });

                    // list not found
                    return false;
                }

                // get list of valid headers
                Dictionary<string, Field> fieldsFromList = new Dictionary<string, Field>();

                // load headers
                foreach (var field in selectedList.Fields)
                {
                    // not read only
                    if (!field.ReadOnlyField)
                    {
                        // add valid header or selected list
                        fieldsFromList.Add(field.InternalName, field);
                    }
                }

                // record index
                int recordIndex = 0;

                // total items processed
                float totalRecordsProcessed = 0;

                // progress value
                float progressValue = 0;
                float totalRecords = (float)records.Count;

                // loop through records
                foreach (string[] record in records)
                {
                    // commit changes every 20 records
                    if (recordIndex > maxListItemsPerQuery)
                    {
                        // commit changes
                        _context.ExecuteQuery();

                        // reset index
                        recordIndex = 0;

                        // update progress bar
                        progressValue = (totalRecordsProcessed / totalRecords) * 100;

                        // send update event
                        importUpdate(new ImportEvent
                        {
                            ImportProgressValue = progressValue,
                            RecordsImported = (int)totalRecordsProcessed,
                            TotalRecords = (int)totalRecords,
                            Success = true,
                            ErrorMessage = string.Empty
                        });
                    }

                    // new list item
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem newItem = selectedList.AddItem(itemCreateInfo);

                    // loop through 
                    for (int i = 0; i < headers.Length; i++)
                    {
                        string currentHeader = headers[i];

                        if (!isItemChecked(currentHeader, mappings))
                        {
                            continue;
                        }

                        string currentFieldName = GetSharepointMappingFromCSVColumn(currentHeader, mappings);

                        // header is not valid, move on
                        if (fieldsFromList.ContainsKey(currentFieldName))
                        {
                            var currentField = fieldsFromList[currentFieldName];
                            var value = convertValue(currentField, record[i]);

                            // populate new item
                            newItem[currentFieldName] = value;

                            // update item
                            newItem.Update();
                        }
                    }

                    // update total count of processed items
                    totalRecordsProcessed++;

                    // incrememnt
                    recordIndex++;
                }

                // commit remaining changes
                _context.ExecuteQuery();

                // send update event
                importUpdate(new ImportEvent
                {
                    ImportProgressValue = progressValue,
                    RecordsImported = (int)totalRecordsProcessed,
                    TotalRecords = (int)totalRecords,
                    Success = true,
                    ErrorMessage = string.Empty
                });

                // send complete event
                importCompleted(new ImportEvent
                {
                    Success = true,
                    RecordsImported = (int)totalRecordsProcessed,
                    TotalRecords = records.Count,
                    ErrorMessage = string.Empty,
                    ImportProgressValue = 100f
                });
            }
            catch (Exception ex)
            {
                // send failed event
                importFailed(new ImportEvent
                {
                    ErrorMessage = ex.Message,
                    Success = false
                });

                return false;
            }

            // success
            return true;
        }

        public void Read(Guid listId)
        {
            List selectedList = _web.Lists.GetById(listId);

            //Loads the site list
            _context.Load(selectedList);
            //Creates a ListItemCollection object from list
            ListItemCollection listItems = selectedList.GetItems(CamlQuery.CreateAllItemsQuery());
            //Loads the listItems
            _context.Load(listItems);
            //Executes the previous queries on the server
            _context.ExecuteQuery();

            //For each listItem...
            foreach (var listItem in listItems)
            {
                var id = listItem.Id;
            }
        }

        private bool isItemChecked(string columnHeader, DataGridView dgvMapping)
        {
            if (dgvMapping == null)
            {
                return false;
            }

            for (int i = 0; i < dgvMapping.Rows.Count; i++)
            {
                DataGridViewRow row = dgvMapping.Rows[i];
                DataGridViewCheckBoxCell cbCellChecked = (DataGridViewCheckBoxCell)row.Cells[0];
                DataGridViewTextBoxCell cbCellName = (DataGridViewTextBoxCell)row.Cells[1];

                bool isChecked = Convert.ToBoolean(cbCellChecked.Value);
                string name = Convert.ToString(cbCellName.Value);

                if (isChecked && name == columnHeader)
                {
                    return true;
                }
            }

            return false;
        }

        public class SharepointList
        {
            public Guid Id { get; set; }
            public string Title { get; set; }
            public FieldCollection Fields { get; set; }
        }

        /// <summary>
        /// convertValue
        /// </summary>
        /// <param name="field"></param>
        /// <param name="value"></param>
        internal object convertValue(Field field, object value)
        {
            var type = field.TypeAsString;
            object valueConverted = null;

            if (value == null || string.IsNullOrEmpty(value.ToString()))
            {
                return valueConverted;
            }

            try
            {
                switch (type)
                {
                    case "Note":
                    case "Text":
                        {
                            valueConverted = value.ToString();
                        }
                        break;
                    case "Boolean":
                        {
                            valueConverted = Convert.ToBoolean(value.ToString());
                        }
                        break;
                    case "Date":
                        {
                            valueConverted = value?.ToString();
                        }
                        break;
                    case "DateTime":
                        {
                            valueConverted = Convert.ToDateTime(value.ToString());
                        }
                        break;
                    case "User":
                        {
                            User user = _web.EnsureUser(value.ToString());
                            _context.Load(user);
                            _context.ExecuteQuery();
                            if (user != null)
                            {
                                valueConverted = new FieldUserValue() { };
                                ((FieldUserValue)valueConverted).LookupId = user.Id;
                            }
                        }
                        break;
                    case "Lookup":
                        {
                            valueConverted = new FieldLookupValue() { LookupId = int.Parse(value.ToString()) };
                        }
                        break;
                    case "LookupMulti":
                        {
                            List<FieldLookupValue> values = new List<FieldLookupValue>();
                            var valuesList = value.ToString().Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var valueList in valuesList)
                            {
                                FieldLookupValue lookupValue = null;
                                try
                                {
                                    lookupValue = new FieldLookupValue() { LookupId = int.Parse(valueList.ToString()) };
                                }
                                catch (Exception ex)
                                {
                                    lookupValue = null;
                                }
                                if (lookupValue != null)
                                {
                                    values.Add(lookupValue);
                                }
                            }
                            valueConverted = values;
                        }
                        break;
                    case "Currency":
                        {
                            valueConverted = Double.Parse(value.ToString());
                        }
                        break;
                    case "UserMulti":
                        {
                            List<FieldUserValue> users = new List<FieldUserValue>();

                            var usersList = value.ToString().Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var userAsString in usersList)
                            {
                                FieldUserValue userValue = null;

                                try
                                {
                                    User user = _web.EnsureUser(userAsString);
                                    _context.Load(user);
                                    _context.ExecuteQuery();
                                    if (user != null)
                                    {
                                        userValue = new FieldUserValue() { };
                                        userValue.LookupId = user.Id;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    userValue = null;
                                }

                                if (userValue != null)
                                {
                                    users.Add(userValue);
                                }
                            }
                            valueConverted = users;
                        }
                        break;
                    case "Number":
                        {
                            valueConverted = Double.Parse(value.ToString());
                        }
                        break;
                    case "Link":
                        {
                            valueConverted = value.ToString();
                        }
                        break;
                    case "Choice":
                        {
                            valueConverted = value.ToString();
                        }
                        break;
                    case "URL":
                        {
                            valueConverted = value.ToString();
                        }
                        break;
                    case "MultiChoice":
                        {
                            var values = value.ToString().Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                            valueConverted = values;
                        }
                        break;
                    default:
                        {
                            valueConverted = value;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                valueConverted = null;
            }

            return valueConverted;
        }

    }
}
