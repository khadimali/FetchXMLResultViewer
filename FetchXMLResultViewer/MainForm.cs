using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.IO;
using McTools.Xrm.Connection.WinForms;
using McTools.Xrm.Connection;

namespace FetchXMLResultViewer
{
    public partial class MainForm : Form
    {
        #region Form properties
        /// <summary>
        /// Connection control
        /// </summary>
        CrmConnectionStatusBar ccsb;

        /// <summary>
        /// Connection manager
        /// </summary>
        ConnectionManager cManager;

        private FormHelper formHelper;

        /// <summary>
        /// Dynamics CRM 2011 organization service
        /// </summary>
        private IOrganizationService service;
        #endregion

        // Constructor
        public MainForm()
        {
            InitializeComponent();

            // Create the connection manager with its events
            this.cManager = ConnectionManager.Instance;
            this.cManager.ConnectionSucceed += new ConnectionManager.ConnectionSucceedEventHandler(cManager_ConnectionSucceed);
            this.cManager.ConnectionFailed += new ConnectionManager.ConnectionFailedEventHandler(cManager_ConnectionFailed);
            this.cManager.StepChanged += new ConnectionManager.StepChangedEventHandler(cManager_StepChanged);
            this.cManager.RequestPassword += new ConnectionManager.RequestPasswordEventHandler(cManager_RequestPassword);
            formHelper = new FormHelper(this);

            // Instantiate and add the connection control to the form
            ccsb = new CrmConnectionStatusBar(formHelper);
            this.Controls.Add(ccsb);

            //this.ccsb.SetMessage("A message to display...");

            this.Text = "FetchXML Result Viewer";

            dgvResult.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        #region Connection tool event handlers

        private bool cManager_RequestPassword(object sender, RequestPasswordEventArgs e)
        {
            return formHelper.RequestPassword(e.ConnectionDetail);
        }

        /// <summary>
        /// Occurs when the connection to a server failed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cManager_ConnectionFailed(object sender, ConnectionFailedEventArgs e)
        {
            // Set error message
            this.ccsb.SetMessage("Error: " + e.FailureReason);

            // Clear the current organization service
            this.service = null;
        }

        /// <summary>
        /// Occurs when the connection to a server succeed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cManager_ConnectionSucceed(object sender, ConnectionSucceedEventArgs e)
        {
            ccsb.RebuildConnectionList();

            // Store connection Organization Service
            this.service = e.OrganizationService;

            // Displays connection status
            this.ccsb.SetConnectionStatus(true, e.ConnectionDetail);

            // Clear the current action message
            this.ccsb.SetMessage(string.Empty);

            // Do action if needed
            if (e.Parameter != null)
            {
                if (e.Parameter.ToString() == "WhoAmI")
                {
                    //WhoAmI();
                }
            }
        }

        /// <summary>
        /// Occurs when the connection manager sends a step change
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cManager_StepChanged(object sender, StepChangedEventArgs e)
        {
            this.ccsb.SetMessage(e.CurrentStep);
        }

        #endregion Connection event handlers

        #region Exception string builders
        private static void BuildStringForGenericException(Exception ex, StringBuilder sb)
        {
            sb.AppendLine(String.Format("The application terminated with an error."));
            sb.AppendLine(String.Format(ex.Message));

            // Display the details of the inner exception.
            if (ex.InnerException != null)
            {
                sb.AppendLine(String.Format(ex.InnerException.Message));

                FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> fe = ex.InnerException
                    as FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault>;
                if (fe != null)
                {
                    sb.AppendLine(String.Format("Timestamp: {0}", fe.Detail.Timestamp));
                    sb.AppendLine(String.Format("Code: {0}", fe.Detail.ErrorCode));
                    sb.AppendLine(String.Format("Message: {0}", fe.Detail.Message));
                    sb.AppendLine(String.Format("Trace: {0}", fe.Detail.TraceText));
                    sb.AppendLine(String.Format("Inner Fault: {0}",
                        null == fe.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault"));
                }
            }
        }

        private static void BuildStringForTimeoutException(TimeoutException ex, StringBuilder sb)
        {
            sb.AppendLine(String.Format("The application terminated with an error."));
            sb.AppendLine(String.Format("Message: {0}", ex.Message));
            sb.AppendLine(String.Format("Stack Trace: {0}", ex.StackTrace));
            sb.AppendLine(String.Format("Inner Fault: {0}",
                null == ex.InnerException.Message ? "No Inner Fault" : ex.InnerException.Message));
        }

        private static void BuildStringForFaultException(FaultException<OrganizationServiceFault> ex, StringBuilder sb)
        {
            sb.AppendLine(String.Format("The application terminated with an error."));
            sb.AppendLine(String.Format("Timestamp: {0}", ex.Detail.Timestamp));
            sb.AppendLine(String.Format("Code: {0}", ex.Detail.ErrorCode));
            sb.AppendLine(String.Format("Message: {0}", ex.Detail.Message));
            sb.AppendLine(String.Format("Trace: {0}", ex.Detail.TraceText));
            sb.AppendLine(String.Format("Inner Fault: {0}",
                null == ex.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault"));
        }
        #endregion

        #region Toolstrip events
        private void manageConnection_Click(object sender, EventArgs e)
        {
            formHelper.DisplayConnectionsList(this);
        }

        private void executeFetch_Click(object sender, EventArgs e)
        {
            ExecuteFetch();
        }

        private void exportToCsv_Click(object sender, EventArgs e)
        {
            string csvString = ExportToCSV();

            SaveFileDialog savefile = new SaveFileDialog();
            // set filters - this can be done in properties as well
            savefile.Filter = "CSV files (*.csv)|*.csv";

            if (savefile.ShowDialog() == DialogResult.OK)
            {
                if (savefile.FileName != "")
                {
                    // Get file name.
                    string name = savefile.FileName;
                    // Write to the file name selected.
                    File.WriteAllText(name, csvString);
                }
            }

        }

        private void mscrmTools_Click(object sender, EventArgs e)
        {
            ToolStripLabel toolStripLabel1 = (ToolStripLabel)sender;

            // Start default browser and navigate to the URL in the
            // tag property.
            System.Diagnostics.Process.Start(toolStripLabel1.Tag.ToString());

            // Set the LinkVisited property to true to change the color.
            toolStripLabel1.LinkVisited = true;
        }
        #endregion

        #region Form controls' event handlers
        private void txtFetchXml_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && (e.KeyCode == Keys.A))
            {
                if (sender != null)
                    ((TextBox)sender).SelectAll();
                e.Handled = true;
            }
        }
        #endregion

        private void ExecuteFetch()
        {
            // Reset DataGridView
            ResetDataGridView();

            // Ask for connection, if service is not already setup.
            if (this.service == null)
            {
                MessageBox.Show("Connect to any Dynamics CRM organization first.");
                formHelper.AskForConnection("whoami");
                return;
            }

            DataSet ds = new DataSet();
            string fetchXml = txtFetchXml.Text;
            EntityCollection ec = null;

            try
            {
                // Fetch CRM data
                ec = this.service.RetrieveMultiple(new FetchExpression(fetchXml));

                // Convert FetchXml string to DataSet
                StringReader sr = new StringReader(fetchXml);
                ds.ReadXml(sr);

                if (ec != null)
                {
                    // PrimaryKey is necessary to be set for finding rows
                    if (ds.Tables.Contains("link-entity"))
                        ds.Tables["link-entity"].PrimaryKey = new DataColumn[] { ds.Tables["link-entity"].Columns[0] };

                    // If "attribute" table has link-entity_Id column, that means fetch query contains one or more link entities.
                    bool hasLinkedEntities = false;
                    if (ds.Tables["attribute"].Columns.Contains("link-entity_Id"))
                        hasLinkedEntities = true;

                    // Loop through each row in "attribute" table and add corresponding column to DataGridView
                    foreach (DataRow row in ds.Tables["attribute"].Rows)
                    {
                        // Make attribute name (for main entity's attribute or linked entity's attribute appropriately) to be added to the DataGridView.
                        string attributeName = MakeAttributeName(ds, hasLinkedEntities, row);

                        dgvResult.Columns.Add(attributeName, attributeName);
                    }

                    // Loop through each entity (fetched CRM data)
                    foreach (var entity in ec.Entities)
                    {
                        // Make attributes collection
                        var lstAttributes = entity.Attributes.ToList();

                        // Add new DataGridViewRow and keep reference of it for later use in foreach block
                        int dgvRowId = dgvResult.Rows.Add();
                        DataGridViewRow dgvRow = dgvResult.Rows[dgvRowId];

                        // Loop through each entity attribute
                        foreach (var attribute in lstAttributes)
                        {
                            var key = attribute.Key;
                            var value = attribute.Value;

                            // If the corresponding column exist in DataGridView, then move the actual attribute value in the newly added DataGridViewRow
                            if (dgvResult.Columns.Contains(key))
                            {
                                MoveAttributeValueToDataGridViewRow(dgvRowId, dgvRow, key, value);
                            }
                        }
                    }
                }
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                StringBuilder sb = new StringBuilder();
                BuildStringForFaultException(ex, sb);
                MessageBox.Show(sb.ToString());

            }
            catch (System.TimeoutException ex)
            {
                StringBuilder sb = new StringBuilder();
                BuildStringForTimeoutException(ex, sb);
                MessageBox.Show(sb.ToString());
            }
            catch (System.Exception ex)
            {
                StringBuilder sb = new StringBuilder();
                BuildStringForGenericException(ex, sb);
                MessageBox.Show(sb.ToString());
            }
            // Additional exceptions to catch: SecurityTokenValidationException, ExpiredSecurityTokenException,
            // SecurityAccessDeniedException, MessageSecurityException, and SecurityNegotiationException.
        }

        private void ResetDataGridView()
        {
            dgvResult.Rows.Clear();
            dgvResult.Columns.Clear();
            dgvResult.Refresh();
        }

        private string MakeAttributeName(DataSet ds, bool hasLinkedEntities, DataRow row)
        {
            string attributeName = row["name"].ToString();

            // If attribute belongs to link entity then qualify the link entity alias before the attribute name.
            string linkEntityId = null;
            if (hasLinkedEntities == true)
                linkEntityId = row["link-entity_Id"].ToString();

            if (!string.IsNullOrEmpty(linkEntityId))
            {
                DataRow linkRow = ds.Tables["link-entity"].Rows.Find(linkEntityId);
                attributeName = linkRow["name"].ToString() + "." + attributeName;
            }
            return attributeName;
        }

        private void MoveAttributeValueToDataGridViewRow(int dgvRowId, DataGridViewRow dgvRow, string key, object value)
        {
            // Move actual attribute value to DataGridViewRow's cell, according to attribute type
            switch (value.GetType().Name)
            {
                case "OptionSetValue":
                    OptionSetValue optionSet = value as OptionSetValue;
                    dgvResult.Rows[dgvRowId].Cells[key].Value = optionSet.Value;
                    break;

                case "Money":
                    Money money = value as Money;
                    dgvResult.Rows[dgvRowId].Cells[key].Value = money.Value;
                    break;

                case "AliasedValue":
                    AliasedValue aliasedVal = value as AliasedValue;
                    dgvResult.Rows[dgvRowId].Cells[key].Value = aliasedVal.Value;
                    break;

                case "EntityReference":
                    // In case of lookup (EntityReference) type of attribute, add another column to DataGridview, in case not added already
                    var nameColumn = key + "_name";
                    if (!dgvResult.Columns.Contains(nameColumn))
                    {
                        var columnIndex = dgvResult.Columns.IndexOf(dgvResult.Columns[key]);

                        DataGridViewColumn dgvNameColumn = new DataGridViewColumn();
                        dgvNameColumn.Name = nameColumn;
                        dgvNameColumn.HeaderText = nameColumn;
                        dgvNameColumn.CellTemplate = new DataGridViewTextBoxCell();

                        dgvResult.Columns.Insert(columnIndex + 1, dgvNameColumn);
                    }

                    // Move lookup id and name to DataGridViewRow's cell
                    EntityReference entityRef = value as EntityReference;
                    dgvRow.Cells[key].Value = entityRef.Id;
                    dgvRow.Cells[nameColumn].Value = entityRef.Name;
                    break;

                default:
                    dgvResult.Rows[dgvRowId].Cells[key].Value = value.ToString();
                    break;
            }
        }

        private string ExportToCSV()
        {
            var sb = new StringBuilder();

            var headers = dgvResult.Columns.Cast<DataGridViewColumn>();
            sb.AppendLine(string.Join(",", headers.Select(column => "\"" + column.HeaderText + "\"").ToArray()));

            foreach (DataGridViewRow row in dgvResult.Rows)
            {
                var cells = row.Cells.Cast<DataGridViewCell>();
                sb.AppendLine(string.Join(",", cells.Select(cell => "\"" + cell.Value + "\"").ToArray()));
            }

            return sb.ToString();
        }

    }
}