namespace AAExcelAddIn
{
    partial class PvtLstObjNavigator
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
            this.tcrNavigator = new System.Windows.Forms.TabControl();
            this.pgePivotTables = new System.Windows.Forms.TabPage();
            this.dgrPivotTables = new System.Windows.Forms.DataGridView();
            this.pgeListObjects = new System.Windows.Forms.TabPage();
            this.dgrListObjects = new System.Windows.Forms.DataGridView();
            this.PivotTable = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtWorksheet = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtDataSourceName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtDataSourceType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtDataSourceDesc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtPageFields = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtColumnFields = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtRowFields = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtDataFields = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Table = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LstObjWorksheet = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LstObjDataSource = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LstObjDataSourceType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LstObjDataSourceDesc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LstObjColumns = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tcrNavigator.SuspendLayout();
            this.pgePivotTables.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgrPivotTables)).BeginInit();
            this.pgeListObjects.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgrListObjects)).BeginInit();
            this.SuspendLayout();
            // 
            // tcrNavigator
            // 
            this.tcrNavigator.Controls.Add(this.pgePivotTables);
            this.tcrNavigator.Controls.Add(this.pgeListObjects);
            this.tcrNavigator.Location = new System.Drawing.Point(12, 12);
            this.tcrNavigator.Name = "tcrNavigator";
            this.tcrNavigator.SelectedIndex = 0;
            this.tcrNavigator.Size = new System.Drawing.Size(638, 370);
            this.tcrNavigator.TabIndex = 0;
            // 
            // pgePivotTables
            // 
            this.pgePivotTables.Controls.Add(this.dgrPivotTables);
            this.pgePivotTables.Location = new System.Drawing.Point(4, 22);
            this.pgePivotTables.Name = "pgePivotTables";
            this.pgePivotTables.Padding = new System.Windows.Forms.Padding(3);
            this.pgePivotTables.Size = new System.Drawing.Size(630, 344);
            this.pgePivotTables.TabIndex = 0;
            this.pgePivotTables.Text = "PiovtTables";
            this.pgePivotTables.UseVisualStyleBackColor = true;
            // 
            // dgrPivotTables
            // 
            this.dgrPivotTables.AllowUserToAddRows = false;
            this.dgrPivotTables.AllowUserToDeleteRows = false;
            this.dgrPivotTables.AllowUserToResizeRows = false;
            this.dgrPivotTables.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dgrPivotTables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgrPivotTables.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.PivotTable,
            this.PvtWorksheet,
            this.PvtDataSourceName,
            this.PvtDataSourceType,
            this.PvtDataSourceDesc,
            this.PvtPageFields,
            this.PvtColumnFields,
            this.PvtRowFields,
            this.PvtDataFields});
            this.dgrPivotTables.Location = new System.Drawing.Point(6, 6);
            this.dgrPivotTables.Name = "dgrPivotTables";
            this.dgrPivotTables.ReadOnly = true;
            this.dgrPivotTables.RowHeadersVisible = false;
            this.dgrPivotTables.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgrPivotTables.Size = new System.Drawing.Size(618, 335);
            this.dgrPivotTables.TabIndex = 1;
            this.dgrPivotTables.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgrPivotTables_CellMouseDoubleClick);
            // 
            // pgeListObjects
            // 
            this.pgeListObjects.Controls.Add(this.dgrListObjects);
            this.pgeListObjects.Location = new System.Drawing.Point(4, 22);
            this.pgeListObjects.Name = "pgeListObjects";
            this.pgeListObjects.Padding = new System.Windows.Forms.Padding(3);
            this.pgeListObjects.Size = new System.Drawing.Size(630, 344);
            this.pgeListObjects.TabIndex = 1;
            this.pgeListObjects.Text = "Tables";
            this.pgeListObjects.UseVisualStyleBackColor = true;
            // 
            // dgrListObjects
            // 
            this.dgrListObjects.AllowUserToAddRows = false;
            this.dgrListObjects.AllowUserToDeleteRows = false;
            this.dgrListObjects.AllowUserToResizeRows = false;
            this.dgrListObjects.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dgrListObjects.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgrListObjects.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Table,
            this.LstObjWorksheet,
            this.LstObjDataSource,
            this.LstObjDataSourceType,
            this.LstObjDataSourceDesc,
            this.LstObjColumns});
            this.dgrListObjects.Location = new System.Drawing.Point(7, 7);
            this.dgrListObjects.Name = "dgrListObjects";
            this.dgrListObjects.ReadOnly = true;
            this.dgrListObjects.RowHeadersVisible = false;
            this.dgrListObjects.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgrListObjects.Size = new System.Drawing.Size(617, 331);
            this.dgrListObjects.TabIndex = 0;
            this.dgrListObjects.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgrListObjects_CellDoubleClick);
            // 
            // PivotTable
            // 
            this.PivotTable.HeaderText = "PivotTable";
            this.PivotTable.Name = "PivotTable";
            this.PivotTable.ReadOnly = true;
            // 
            // PvtWorksheet
            // 
            this.PvtWorksheet.HeaderText = "Worksheet";
            this.PvtWorksheet.Name = "PvtWorksheet";
            this.PvtWorksheet.ReadOnly = true;
            // 
            // PvtDataSourceName
            // 
            this.PvtDataSourceName.HeaderText = "Data Source Name";
            this.PvtDataSourceName.Name = "PvtDataSourceName";
            this.PvtDataSourceName.ReadOnly = true;
            this.PvtDataSourceName.Width = 150;
            // 
            // PvtDataSourceType
            // 
            this.PvtDataSourceType.HeaderText = "Data Source Type";
            this.PvtDataSourceType.Name = "PvtDataSourceType";
            this.PvtDataSourceType.ReadOnly = true;
            this.PvtDataSourceType.Width = 150;
            // 
            // PvtDataSourceDesc
            // 
            this.PvtDataSourceDesc.HeaderText = "Data Source Description";
            this.PvtDataSourceDesc.Name = "PvtDataSourceDesc";
            this.PvtDataSourceDesc.ReadOnly = true;
            this.PvtDataSourceDesc.Width = 300;
            // 
            // PvtPageFields
            // 
            this.PvtPageFields.HeaderText = "Filter Fields";
            this.PvtPageFields.Name = "PvtPageFields";
            this.PvtPageFields.ReadOnly = true;
            this.PvtPageFields.Width = 150;
            // 
            // PvtColumnFields
            // 
            this.PvtColumnFields.HeaderText = "Column Fields";
            this.PvtColumnFields.Name = "PvtColumnFields";
            this.PvtColumnFields.ReadOnly = true;
            this.PvtColumnFields.Width = 150;
            // 
            // PvtRowFields
            // 
            this.PvtRowFields.HeaderText = "Row Fields";
            this.PvtRowFields.Name = "PvtRowFields";
            this.PvtRowFields.ReadOnly = true;
            this.PvtRowFields.Width = 150;
            // 
            // PvtDataFields
            // 
            this.PvtDataFields.HeaderText = "Value Fields";
            this.PvtDataFields.Name = "PvtDataFields";
            this.PvtDataFields.ReadOnly = true;
            this.PvtDataFields.Width = 150;
            // 
            // Table
            // 
            this.Table.HeaderText = "Table";
            this.Table.Name = "Table";
            this.Table.ReadOnly = true;
            // 
            // LstObjWorksheet
            // 
            this.LstObjWorksheet.HeaderText = "Worksheet";
            this.LstObjWorksheet.Name = "LstObjWorksheet";
            this.LstObjWorksheet.ReadOnly = true;
            // 
            // LstObjDataSource
            // 
            this.LstObjDataSource.HeaderText = "Data Source Name";
            this.LstObjDataSource.Name = "LstObjDataSource";
            this.LstObjDataSource.ReadOnly = true;
            this.LstObjDataSource.Width = 150;
            // 
            // LstObjDataSourceType
            // 
            this.LstObjDataSourceType.HeaderText = "Data Source Type";
            this.LstObjDataSourceType.Name = "LstObjDataSourceType";
            this.LstObjDataSourceType.ReadOnly = true;
            this.LstObjDataSourceType.Width = 150;
            // 
            // LstObjDataSourceDesc
            // 
            this.LstObjDataSourceDesc.HeaderText = "Data Source Description";
            this.LstObjDataSourceDesc.Name = "LstObjDataSourceDesc";
            this.LstObjDataSourceDesc.ReadOnly = true;
            this.LstObjDataSourceDesc.Width = 300;
            // 
            // LstObjColumns
            // 
            this.LstObjColumns.HeaderText = "Columns";
            this.LstObjColumns.Name = "LstObjColumns";
            this.LstObjColumns.ReadOnly = true;
            this.LstObjColumns.Width = 150;
            // 
            // PvtLstObjNavigator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(662, 394);
            this.Controls.Add(this.tcrNavigator);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PvtLstObjNavigator";
            this.Text = "PivotTable/List Object Navigator";
            this.Load += new System.EventHandler(this.PvtLstObjNavigator_Load);
            this.tcrNavigator.ResumeLayout(false);
            this.pgePivotTables.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgrPivotTables)).EndInit();
            this.pgeListObjects.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgrListObjects)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tcrNavigator;
        private System.Windows.Forms.TabPage pgePivotTables;
        private System.Windows.Forms.DataGridView dgrPivotTables;
        private System.Windows.Forms.TabPage pgeListObjects;
        private System.Windows.Forms.DataGridView dgrListObjects;
        private System.Windows.Forms.DataGridViewTextBoxColumn PivotTable;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtWorksheet;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtDataSourceName;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtDataSourceType;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtDataSourceDesc;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtPageFields;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtColumnFields;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtRowFields;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtDataFields;
        private System.Windows.Forms.DataGridViewTextBoxColumn Table;
        private System.Windows.Forms.DataGridViewTextBoxColumn LstObjWorksheet;
        private System.Windows.Forms.DataGridViewTextBoxColumn LstObjDataSource;
        private System.Windows.Forms.DataGridViewTextBoxColumn LstObjDataSourceType;
        private System.Windows.Forms.DataGridViewTextBoxColumn LstObjDataSourceDesc;
        private System.Windows.Forms.DataGridViewTextBoxColumn LstObjColumns;
    }
}