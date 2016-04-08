namespace FetchXMLResultViewer
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.lblFetchXml = new System.Windows.Forms.Label();
            this.txtFetchXml = new System.Windows.Forms.TextBox();
            this.dgvResult = new System.Windows.Forms.DataGridView();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButton_ManageConnection = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton_ExecuteFetch = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton_ExportToCsv = new System.Windows.Forms.ToolStripButton();
            this.toolStripLabel_MscrmTools = new System.Windows.Forms.ToolStripLabel();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResult)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblFetchXml
            // 
            this.lblFetchXml.AutoSize = true;
            this.lblFetchXml.Location = new System.Drawing.Point(12, 60);
            this.lblFetchXml.Name = "lblFetchXml";
            this.lblFetchXml.Size = new System.Drawing.Size(272, 17);
            this.lblFetchXml.TabIndex = 0;
            this.lblFetchXml.Text = "Enter your FetchXML string below:";
            // 
            // txtFetchXml
            // 
            this.txtFetchXml.Location = new System.Drawing.Point(15, 80);
            this.txtFetchXml.Multiline = true;
            this.txtFetchXml.Name = "txtFetchXml";
            this.txtFetchXml.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtFetchXml.Size = new System.Drawing.Size(1157, 231);
            this.txtFetchXml.TabIndex = 1;
            this.txtFetchXml.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFetchXml_KeyDown);
            // 
            // dgvResult
            // 
            this.dgvResult.AllowUserToAddRows = false;
            this.dgvResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvResult.Location = new System.Drawing.Point(15, 317);
            this.dgvResult.Name = "dgvResult";
            this.dgvResult.ReadOnly = true;
            this.dgvResult.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvResult.Size = new System.Drawing.Size(1157, 413);
            this.dgvResult.TabIndex = 2;
            // 
            // toolStrip1
            // 
            this.toolStrip1.AutoSize = false;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton_ManageConnection,
            this.toolStripButton_ExecuteFetch,
            this.toolStripButton_ExportToCsv,
            this.toolStripLabel_MscrmTools});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1184, 50);
            this.toolStrip1.TabIndex = 5;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton_ManageConnection
            // 
            this.toolStripButton_ManageConnection.AutoSize = false;
            this.toolStripButton_ManageConnection.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton_ManageConnection.Image")));
            this.toolStripButton_ManageConnection.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.toolStripButton_ManageConnection.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton_ManageConnection.Name = "toolStripButton_ManageConnection";
            this.toolStripButton_ManageConnection.Size = new System.Drawing.Size(175, 47);
            this.toolStripButton_ManageConnection.Text = "Manage Connection";
            this.toolStripButton_ManageConnection.Click += new System.EventHandler(this.manageConnection_Click);
            // 
            // toolStripButton_ExecuteFetch
            // 
            this.toolStripButton_ExecuteFetch.AutoSize = false;
            this.toolStripButton_ExecuteFetch.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton_ExecuteFetch.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton_ExecuteFetch.Image")));
            this.toolStripButton_ExecuteFetch.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.toolStripButton_ExecuteFetch.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton_ExecuteFetch.Name = "toolStripButton_ExecuteFetch";
            this.toolStripButton_ExecuteFetch.Size = new System.Drawing.Size(47, 47);
            this.toolStripButton_ExecuteFetch.Text = "Execute FetchXML";
            this.toolStripButton_ExecuteFetch.Click += new System.EventHandler(this.executeFetch_Click);
            // 
            // toolStripButton_ExportToCsv
            // 
            this.toolStripButton_ExportToCsv.AutoSize = false;
            this.toolStripButton_ExportToCsv.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton_ExportToCsv.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton_ExportToCsv.Image")));
            this.toolStripButton_ExportToCsv.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.toolStripButton_ExportToCsv.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton_ExportToCsv.Name = "toolStripButton_ExportToCsv";
            this.toolStripButton_ExportToCsv.Size = new System.Drawing.Size(47, 47);
            this.toolStripButton_ExportToCsv.Text = "Export as CSV";
            this.toolStripButton_ExportToCsv.Click += new System.EventHandler(this.exportToCsv_Click);
            // 
            // toolStripLabel_MscrmTools
            // 
            this.toolStripLabel_MscrmTools.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripLabel_MscrmTools.AutoSize = false;
            this.toolStripLabel_MscrmTools.IsLink = true;
            this.toolStripLabel_MscrmTools.Name = "toolStripLabel_MscrmTools";
            this.toolStripLabel_MscrmTools.Size = new System.Drawing.Size(306, 20);
            this.toolStripLabel_MscrmTools.Tag = "https://github.com/MscrmTools/MscrmTools.Xrm.Connection";
            this.toolStripLabel_MscrmTools.Text = "This application is using connection tool by MscrmTools";
            this.toolStripLabel_MscrmTools.Click += new System.EventHandler(this.mscrmTools_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1184, 762);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.dgvResult);
            this.Controls.Add(this.lblFetchXml);
            this.Controls.Add(this.txtFetchXml);
            this.Font = new System.Drawing.Font("Consolas", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dgvResult)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblFetchXml;
        private System.Windows.Forms.TextBox txtFetchXml;
        private System.Windows.Forms.DataGridView dgvResult;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButton_ExecuteFetch;
        private System.Windows.Forms.ToolStripButton toolStripButton_ExportToCsv;
        private System.Windows.Forms.ToolStripButton toolStripButton_ManageConnection;
        private System.Windows.Forms.ToolStripLabel toolStripLabel_MscrmTools;
    }
}

