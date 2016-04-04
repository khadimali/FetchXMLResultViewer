using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk.Query;
using System.IO;
using McTools.Xrm.Connection.WinForms;
using McTools.Xrm.Connection;

namespace FetchXMLResultViewer
{
    public partial class MainForm : Form
    {
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

        private string formText = "FetchXML Result Viewer";

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
        
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = formText;
        }

        //private void GetOrganizationService()
        //{
        //    try
        //    {
        //        this.Text = "Please wait ... connecting to CRM";

        //        // Obtain connection configuration information for the Microsoft Dynamics
        //        // CRM organization web service.
        //        String connectionString = GetServiceConfiguration();

        //        if (connectionString != null)
        //        {
        //            // Connect to the CRM web service using a connection string.
        //            CrmServiceClient conn = new CrmServiceClient(connectionString);

        //            // Cast the proxy client to the IOrganizationService interface.
        //            _orgService = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
        //        }
        //    }
        //    catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
        //    {
        //        StringBuilder sb = new StringBuilder();
        //        BuildStringForFaultException(ex, sb);
        //        MessageBox.Show(sb.ToString());

        //    }
        //    catch (System.TimeoutException ex)
        //    {
        //        StringBuilder sb = new StringBuilder();
        //        BuildStringForTimeoutException(ex, sb);
        //        MessageBox.Show(sb.ToString());
        //    }
        //    catch (System.Exception ex)
        //    {
        //        StringBuilder sb = new StringBuilder();
        //        BuildStringForGenericException(ex, sb);
        //        MessageBox.Show(sb.ToString());
        //    }
        //    // Additional exceptions to catch: SecurityTokenValidationException, ExpiredSecurityTokenException,
        //    // SecurityAccessDeniedException, MessageSecurityException, and SecurityNegotiationException.
        //    finally
        //    {
        //        this.Text = formText;
        //    }
        //}

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

        /// <summary>
        /// Gets web service connection information from the app.config file.
        /// If there is more than one available, the user is prompted to select
        /// the desired connection configuration by name.
        /// </summary>
        /// <returns>A string containing web service connection configuration information.</returns>
        //private String GetServiceConfiguration()
        //{
        //    // Get available connection strings from app.config.
        //    int count = ConfigurationManager.ConnectionStrings.Count;

        //    // Create a filter list of connection strings so that we have a list of valid
        //    // connection strings for Microsoft Dynamics CRM only.
        //    List<KeyValuePair<String, String>> filteredConnectionStrings =
        //        new List<KeyValuePair<String, String>>();

        //    for (int a = 0; a < count; a++)
        //    {
        //        if (isValidConnectionString(ConfigurationManager.ConnectionStrings[a].ConnectionString))
        //            filteredConnectionStrings.Add
        //                (new KeyValuePair<string, string>
        //                    (ConfigurationManager.ConnectionStrings[a].Name,
        //                    ConfigurationManager.ConnectionStrings[a].ConnectionString));
        //    }

        //    // No valid connections strings found. Write out and error message.
        //    //if (filteredConnectionStrings.Count == 0)
        //    if (filteredConnectionStrings.Count != 1)
        //    {
        //        StringBuilder sb = new StringBuilder();
        //        sb.AppendLine("An app.config file containing only one valid Microsoft Dynamics CRM " +
        //            "connection string configuration must exist in the run-time folder.");
        //        sb.AppendLine("There are several commented out example connection strings in " +
        //            "the provided app.config file. Uncomment one of them and modify the string according " +
        //            "to your Microsoft Dynamics CRM installation. Then re-run the sample.");
        //        return null;
        //    }

        //    // If one valid connection string is found, use that.
        //    if (filteredConnectionStrings.Count == 1)
        //    {
        //        return filteredConnectionStrings[0].Value;
        //    }

        //    //// If more than one valid connection string is found, let the user decide which to use.
        //    //if (filteredConnectionStrings.Count > 1)
        //    //{
        //    //    Console.WriteLine("The following connections are available:");
        //    //    Console.WriteLine("------------------------------------------------");

        //    //    for (int i = 0; i < filteredConnectionStrings.Count; i++)
        //    //    {
        //    //        Console.Write("\n({0}) {1}\t",
        //    //        i + 1, filteredConnectionStrings[i].Key);
        //    //    }

        //    //    Console.WriteLine();

        //    //    Console.Write("\nType the number of the connection to use (1-{0}) [{0}] : ",
        //    //        filteredConnectionStrings.Count);
        //    //    String input = Console.ReadLine();
        //    //    int configNumber;
        //    //    if (input == String.Empty) input = filteredConnectionStrings.Count.ToString();
        //    //    if (!Int32.TryParse(input, out configNumber) || configNumber > count ||
        //    //        configNumber == 0)
        //    //    {
        //    //        Console.WriteLine("Option not valid.");
        //    //        return null;
        //    //    }

        //    //    return filteredConnectionStrings[configNumber - 1].Value;

        //    //}

        //    return null;

        //}

        /// <summary>
        /// Verifies if a connection string is valid for Microsoft Dynamics CRM.
        /// </summary>
        /// <returns>True for a valid string, otherwise False.</returns>
        //private Boolean isValidConnectionString(String connectionString)
        //{
        //    // At a minimum, a connection string must contain one of these arguments.
        //    if (connectionString.Contains("Url=") ||
        //        connectionString.Contains("Server=") ||
        //        connectionString.Contains("ServiceUri="))
        //        return true;

        //    return false;
        //}

        private void btnExecute_Click(object sender, EventArgs e)
        {

        }

        private void txtFetchXml_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && (e.KeyCode == Keys.A))
            {
                if (sender != null)
                    ((TextBox)sender).SelectAll();
                e.Handled = true;
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

        private void btnExportToCsv_Click(object sender, EventArgs e)
        {
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            dgvResult.Rows.Clear();
            dgvResult.Columns.Clear();
            dgvResult.Refresh();

            if (this.service == null)
            {
                //this.ccsb.SetMessage("Connect to CRM first ...");
                MessageBox.Show("Connect to any Dynamics CRM organization first.");
                formHelper.AskForConnection("whoam");
                return;
            }

            DataSet ds = new DataSet();
            string fetchXml = txtFetchXml.Text;
            EntityCollection ec = null;

            try
            {
                ec = this.service.RetrieveMultiple(new FetchExpression(fetchXml));

                dgvResult.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                if (ec != null)
                {
                    StringReader sr = new StringReader(fetchXml);
                    ds.ReadXml(sr);

                    if (ds.Tables.Contains("link-entity"))
                        ds.Tables["link-entity"].PrimaryKey = new DataColumn[] { ds.Tables["link-entity"].Columns[0] };

                    bool hasLinkedEntities = false;
                    if (ds.Tables["attribute"].Columns.Contains("link-entity_Id"))
                        hasLinkedEntities = true;

                    foreach (DataRow row in ds.Tables["attribute"].Rows)
                    {
                        string attributeName = row["name"].ToString();

                        string linkEntityId = null;
                        if (hasLinkedEntities == true)
                            linkEntityId = row["link-entity_Id"].ToString();

                        if (!string.IsNullOrEmpty(linkEntityId))
                        {
                            DataRow linkRow = ds.Tables["link-entity"].Rows.Find(linkEntityId);
                            attributeName = linkRow["name"].ToString() + "." + attributeName;
                        }
                        dgvResult.Columns.Add(attributeName, attributeName);
                    }

                    foreach (var entity in ec.Entities)
                    {
                        var lstAttributes = entity.Attributes.ToList();

                        int dgvRowId = dgvResult.Rows.Add();
                        DataGridViewRow dgvRow = dgvResult.Rows[dgvRowId];

                        foreach (var attribute in lstAttributes)
                        {
                            //var key = attribute.Key.IndexOf('.') > 0 ? 
                            //    attribute.Key.Substring(attribute.Key.IndexOf('.') + 1) : 
                            //    attribute.Key;

                            var key = attribute.Key;
                            var value = attribute.Value;

                            if (dgvResult.Columns.Contains(key))
                            {
                                // AliasedValue
                                // EntityReference

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

                                    case "EntityReference":
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

                                        EntityReference entityRef = value as EntityReference;
                                        dgvRow.Cells[key].Value = entityRef.Id;
                                        dgvRow.Cells[nameColumn].Value = entityRef.Name;
                                        break;

                                    case "AliasedValue":
                                        AliasedValue aliasedVal = value as AliasedValue;
                                        dgvResult.Rows[dgvRowId].Cells[key].Value = aliasedVal.Value;
                                        break;

                                    default:
                                        dgvResult.Rows[dgvRowId].Cells[key].Value = value.ToString();
                                        break;
                                }
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

        private void toolStripButton3_Click(object sender, EventArgs e)
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

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            formHelper.DisplayConnectionsList(this);
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            ToolStripLabel toolStripLabel1 = (ToolStripLabel)sender;

            // Start default browser and navigate to the URL in the
            // tag property.
            System.Diagnostics.Process.Start(toolStripLabel1.Tag.ToString());

            // Set the LinkVisited property to true to change the color.
            toolStripLabel1.LinkVisited = true;
        }
    }
}
