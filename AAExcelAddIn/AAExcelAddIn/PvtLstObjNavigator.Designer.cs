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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tcrNavigator = new System.Windows.Forms.TabControl();
            this.pgePivotTables = new System.Windows.Forms.TabPage();
            this.dgrPivotTables = new System.Windows.Forms.DataGridView();
            this.PivotTable = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtWorksheet = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtGoTo = new System.Windows.Forms.DataGridViewButtonColumn();
            this.PvtGrouping = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.PvtDataSourceName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtDataSourceType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtDataSourceDesc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtLastRefreshed = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtPageFields = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtColumnFields = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtRowFields = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PvtDataFields = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pgeListObjects = new System.Windows.Forms.TabPage();
            this.dgrListObjects = new System.Windows.Forms.DataGridView();
            this.Table = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LstObjWorksheet = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LstObjDataSource = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LstObjDataSourceType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LstObjDataSourceDesc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LstObjColumns = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pgeDataSources = new System.Windows.Forms.TabPage();
            this.dgrPvtChcFields = new System.Windows.Forms.DataGridView();
            this.pvtChcFieldSrcName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pvtChcFieldDataType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.dgrDataSources = new System.Windows.Forms.DataGridView();
            this.DataSource = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dtaSrcDescription = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dtaSrcType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dtaSrcPvtCache = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.dtaSrcPvtChcMemory = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dtaSrcLastUpdated = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dtaSrcReadOnly = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.dtaSrcCommandText = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dtaSrcConnectionFile = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dtaSrcCommandType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pgeGroupings = new System.Windows.Forms.TabPage();
            this.dgrGroupings = new System.Windows.Forms.DataGridView();
            this.tcrNavigator.SuspendLayout();
            this.pgePivotTables.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgrPivotTables)).BeginInit();
            this.pgeListObjects.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgrListObjects)).BeginInit();
            this.pgeDataSources.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgrPvtChcFields)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgrDataSources)).BeginInit();
            this.pgeGroupings.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgrGroupings)).BeginInit();
            this.SuspendLayout();
            // 
            // tcrNavigator
            // 
            this.tcrNavigator.Controls.Add(this.pgePivotTables);
            this.tcrNavigator.Controls.Add(this.pgeListObjects);
            this.tcrNavigator.Controls.Add(this.pgeDataSources);
            this.tcrNavigator.Controls.Add(this.pgeGroupings);
            this.tcrNavigator.Location = new System.Drawing.Point(12, 12);
            this.tcrNavigator.Name = "tcrNavigator";
            this.tcrNavigator.SelectedIndex = 0;
            this.tcrNavigator.Size = new System.Drawing.Size(841, 454);
            this.tcrNavigator.TabIndex = 0;
            // 
            // pgePivotTables
            // 
            this.pgePivotTables.Controls.Add(this.dgrPivotTables);
            this.pgePivotTables.Location = new System.Drawing.Point(4, 22);
            this.pgePivotTables.Name = "pgePivotTables";
            this.pgePivotTables.Padding = new System.Windows.Forms.Padding(3);
            this.pgePivotTables.Size = new System.Drawing.Size(833, 428);
            this.pgePivotTables.TabIndex = 0;
            this.pgePivotTables.Text = "PiovtTables";
            this.pgePivotTables.UseVisualStyleBackColor = true;
            // 
            // dgrPivotTables
            // 
            this.dgrPivotTables.AllowUserToAddRows = false;
            this.dgrPivotTables.AllowUserToDeleteRows = false;
            this.dgrPivotTables.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dgrPivotTables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgrPivotTables.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.PivotTable,
            this.PvtWorksheet,
            this.PvtGoTo,
            this.PvtGrouping,
            this.PvtDataSourceName,
            this.PvtDataSourceType,
            this.PvtDataSourceDesc,
            this.PvtLastRefreshed,
            this.PvtPageFields,
            this.PvtColumnFields,
            this.PvtRowFields,
            this.PvtDataFields});
            this.dgrPivotTables.Location = new System.Drawing.Point(6, 6);
            this.dgrPivotTables.Name = "dgrPivotTables";
            this.dgrPivotTables.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgrPivotTables.Size = new System.Drawing.Size(821, 416);
            this.dgrPivotTables.TabIndex = 1;
            this.dgrPivotTables.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgrPivotTables_CellClick);
            this.dgrPivotTables.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgrPivotTables_CellEndEdit);
            // 
            // PivotTable
            // 
            this.PivotTable.HeaderText = "PivotTable";
            this.PivotTable.Name = "PivotTable";
            // 
            // PvtWorksheet
            // 
            this.PvtWorksheet.HeaderText = "Worksheet";
            this.PvtWorksheet.Name = "PvtWorksheet";
            this.PvtWorksheet.ReadOnly = true;
            // 
            // PvtGoTo
            // 
            this.PvtGoTo.HeaderText = "Go To";
            this.PvtGoTo.MinimumWidth = 50;
            this.PvtGoTo.Name = "PvtGoTo";
            this.PvtGoTo.ReadOnly = true;
            this.PvtGoTo.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.PvtGoTo.Text = "Go To";
            this.PvtGoTo.UseColumnTextForButtonValue = true;
            this.PvtGoTo.Width = 50;
            // 
            // PvtGrouping
            // 
            this.PvtGrouping.HeaderText = "Grouping";
            this.PvtGrouping.Name = "PvtGrouping";
            this.PvtGrouping.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.PvtGrouping.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
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
            // PvtLastRefreshed
            // 
            this.PvtLastRefreshed.HeaderText = "Last Refreshed";
            this.PvtLastRefreshed.Name = "PvtLastRefreshed";
            this.PvtLastRefreshed.ReadOnly = true;
            this.PvtLastRefreshed.Width = 150;
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
            // pgeListObjects
            // 
            this.pgeListObjects.Controls.Add(this.dgrListObjects);
            this.pgeListObjects.Location = new System.Drawing.Point(4, 22);
            this.pgeListObjects.Name = "pgeListObjects";
            this.pgeListObjects.Padding = new System.Windows.Forms.Padding(3);
            this.pgeListObjects.Size = new System.Drawing.Size(833, 428);
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
            this.dgrListObjects.Size = new System.Drawing.Size(820, 415);
            this.dgrListObjects.TabIndex = 0;
            this.dgrListObjects.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgrListObjects_CellDoubleClick);
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
            // pgeDataSources
            // 
            this.pgeDataSources.Controls.Add(this.dgrPvtChcFields);
            this.pgeDataSources.Controls.Add(this.label1);
            this.pgeDataSources.Controls.Add(this.dgrDataSources);
            this.pgeDataSources.Location = new System.Drawing.Point(4, 22);
            this.pgeDataSources.Name = "pgeDataSources";
            this.pgeDataSources.Padding = new System.Windows.Forms.Padding(3);
            this.pgeDataSources.Size = new System.Drawing.Size(833, 428);
            this.pgeDataSources.TabIndex = 2;
            this.pgeDataSources.Text = "Data Sources";
            this.pgeDataSources.UseVisualStyleBackColor = true;
            // 
            // dgrPvtChcFields
            // 
            this.dgrPvtChcFields.AllowUserToAddRows = false;
            this.dgrPvtChcFields.AllowUserToDeleteRows = false;
            this.dgrPvtChcFields.AllowUserToResizeRows = false;
            this.dgrPvtChcFields.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dgrPvtChcFields.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgrPvtChcFields.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.pvtChcFieldSrcName,
            this.pvtChcFieldDataType});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgrPvtChcFields.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgrPvtChcFields.Location = new System.Drawing.Point(7, 238);
            this.dgrPvtChcFields.MultiSelect = false;
            this.dgrPvtChcFields.Name = "dgrPvtChcFields";
            this.dgrPvtChcFields.ReadOnly = true;
            this.dgrPvtChcFields.RowHeadersVisible = false;
            this.dgrPvtChcFields.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgrPvtChcFields.Size = new System.Drawing.Size(336, 184);
            this.dgrPvtChcFields.TabIndex = 2;
            // 
            // pvtChcFieldSrcName
            // 
            this.pvtChcFieldSrcName.HeaderText = "Field Name";
            this.pvtChcFieldSrcName.Name = "pvtChcFieldSrcName";
            this.pvtChcFieldSrcName.ReadOnly = true;
            // 
            // pvtChcFieldDataType
            // 
            this.pvtChcFieldDataType.HeaderText = "Data Type";
            this.pvtChcFieldDataType.Name = "pvtChcFieldDataType";
            this.pvtChcFieldDataType.ReadOnly = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 219);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "PivotCache Fields";
            // 
            // dgrDataSources
            // 
            this.dgrDataSources.AllowUserToAddRows = false;
            this.dgrDataSources.AllowUserToDeleteRows = false;
            this.dgrDataSources.AllowUserToResizeRows = false;
            this.dgrDataSources.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dgrDataSources.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgrDataSources.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DataSource,
            this.dtaSrcDescription,
            this.dtaSrcType,
            this.dtaSrcPvtCache,
            this.dtaSrcPvtChcMemory,
            this.dtaSrcLastUpdated,
            this.dtaSrcReadOnly,
            this.dtaSrcCommandText,
            this.dtaSrcConnectionFile,
            this.dtaSrcCommandType});
            this.dgrDataSources.Location = new System.Drawing.Point(3, 6);
            this.dgrDataSources.Name = "dgrDataSources";
            this.dgrDataSources.ReadOnly = true;
            this.dgrDataSources.RowHeadersVisible = false;
            this.dgrDataSources.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgrDataSources.Size = new System.Drawing.Size(824, 198);
            this.dgrDataSources.TabIndex = 0;
            this.dgrDataSources.SelectionChanged += new System.EventHandler(this.dgrDataSources_SelectionChanged);
            // 
            // DataSource
            // 
            this.DataSource.HeaderText = "Data Source";
            this.DataSource.Name = "DataSource";
            this.DataSource.ReadOnly = true;
            // 
            // dtaSrcDescription
            // 
            this.dtaSrcDescription.HeaderText = "Description";
            this.dtaSrcDescription.Name = "dtaSrcDescription";
            this.dtaSrcDescription.ReadOnly = true;
            this.dtaSrcDescription.Width = 200;
            // 
            // dtaSrcType
            // 
            this.dtaSrcType.HeaderText = "Type";
            this.dtaSrcType.Name = "dtaSrcType";
            this.dtaSrcType.ReadOnly = true;
            // 
            // dtaSrcPvtCache
            // 
            this.dtaSrcPvtCache.HeaderText = "PivotCache";
            this.dtaSrcPvtCache.Name = "dtaSrcPvtCache";
            this.dtaSrcPvtCache.ReadOnly = true;
            // 
            // dtaSrcPvtChcMemory
            // 
            this.dtaSrcPvtChcMemory.HeaderText = "Cache Size (MB)";
            this.dtaSrcPvtChcMemory.Name = "dtaSrcPvtChcMemory";
            this.dtaSrcPvtChcMemory.ReadOnly = true;
            this.dtaSrcPvtChcMemory.Width = 125;
            // 
            // dtaSrcLastUpdated
            // 
            this.dtaSrcLastUpdated.HeaderText = "Last Updated";
            this.dtaSrcLastUpdated.Name = "dtaSrcLastUpdated";
            this.dtaSrcLastUpdated.ReadOnly = true;
            // 
            // dtaSrcReadOnly
            // 
            this.dtaSrcReadOnly.HeaderText = "Read Only";
            this.dtaSrcReadOnly.Name = "dtaSrcReadOnly";
            this.dtaSrcReadOnly.ReadOnly = true;
            this.dtaSrcReadOnly.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dtaSrcReadOnly.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // dtaSrcCommandText
            // 
            this.dtaSrcCommandText.HeaderText = "Command Text";
            this.dtaSrcCommandText.Name = "dtaSrcCommandText";
            this.dtaSrcCommandText.ReadOnly = true;
            this.dtaSrcCommandText.Width = 150;
            // 
            // dtaSrcConnectionFile
            // 
            this.dtaSrcConnectionFile.HeaderText = "Connection File";
            this.dtaSrcConnectionFile.Name = "dtaSrcConnectionFile";
            this.dtaSrcConnectionFile.ReadOnly = true;
            this.dtaSrcConnectionFile.Width = 200;
            // 
            // dtaSrcCommandType
            // 
            this.dtaSrcCommandType.HeaderText = "Command Type";
            this.dtaSrcCommandType.Name = "dtaSrcCommandType";
            this.dtaSrcCommandType.ReadOnly = true;
            this.dtaSrcCommandType.Width = 150;
            // 
            // pgeGroupings
            // 
            this.pgeGroupings.Controls.Add(this.dgrGroupings);
            this.pgeGroupings.Location = new System.Drawing.Point(4, 22);
            this.pgeGroupings.Name = "pgeGroupings";
            this.pgeGroupings.Padding = new System.Windows.Forms.Padding(3);
            this.pgeGroupings.Size = new System.Drawing.Size(833, 428);
            this.pgeGroupings.TabIndex = 3;
            this.pgeGroupings.Text = "Groupings";
            this.pgeGroupings.UseVisualStyleBackColor = true;
            // 
            // dgrGroupings
            // 
            this.dgrGroupings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgrGroupings.Location = new System.Drawing.Point(6, 6);
            this.dgrGroupings.Name = "dgrGroupings";
            this.dgrGroupings.Size = new System.Drawing.Size(821, 416);
            this.dgrGroupings.TabIndex = 0;
            this.dgrGroupings.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dgrGroupings_CellBeginEdit);
            this.dgrGroupings.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgrGroupings_CellEndEdit);
            this.dgrGroupings.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dgrGroupings_CellValidating);
            this.dgrGroupings.UserDeletingRow += new System.Windows.Forms.DataGridViewRowCancelEventHandler(this.dgrGroupings_UserDeletingRow);
            // 
            // PvtLstObjNavigator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(865, 478);
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
            this.pgeDataSources.ResumeLayout(false);
            this.pgeDataSources.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgrPvtChcFields)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgrDataSources)).EndInit();
            this.pgeGroupings.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgrGroupings)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tcrNavigator;
        private System.Windows.Forms.TabPage pgePivotTables;
        private System.Windows.Forms.DataGridView dgrPivotTables;
        private System.Windows.Forms.TabPage pgeListObjects;
        private System.Windows.Forms.DataGridView dgrListObjects;
        private System.Windows.Forms.DataGridViewTextBoxColumn Table;
        private System.Windows.Forms.DataGridViewTextBoxColumn LstObjWorksheet;
        private System.Windows.Forms.DataGridViewTextBoxColumn LstObjDataSource;
        private System.Windows.Forms.DataGridViewTextBoxColumn LstObjDataSourceType;
        private System.Windows.Forms.DataGridViewTextBoxColumn LstObjDataSourceDesc;
        private System.Windows.Forms.DataGridViewTextBoxColumn LstObjColumns;
        private System.Windows.Forms.TabPage pgeDataSources;
        private System.Windows.Forms.DataGridView dgrDataSources;
        private System.Windows.Forms.DataGridView dgrPvtChcFields;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridViewTextBoxColumn pvtChcFieldSrcName;
        private System.Windows.Forms.DataGridViewTextBoxColumn pvtChcFieldDataType;
        private System.Windows.Forms.DataGridViewTextBoxColumn DataSource;
        private System.Windows.Forms.DataGridViewTextBoxColumn dtaSrcDescription;
        private System.Windows.Forms.DataGridViewTextBoxColumn dtaSrcType;
        private System.Windows.Forms.DataGridViewCheckBoxColumn dtaSrcPvtCache;
        private System.Windows.Forms.DataGridViewTextBoxColumn dtaSrcPvtChcMemory;
        private System.Windows.Forms.DataGridViewTextBoxColumn dtaSrcLastUpdated;
        private System.Windows.Forms.DataGridViewCheckBoxColumn dtaSrcReadOnly;
        private System.Windows.Forms.DataGridViewTextBoxColumn dtaSrcCommandText;
        private System.Windows.Forms.DataGridViewTextBoxColumn dtaSrcConnectionFile;
        private System.Windows.Forms.DataGridViewTextBoxColumn dtaSrcCommandType;
        private System.Windows.Forms.TabPage pgeGroupings;
        private System.Windows.Forms.DataGridView dgrGroupings;
        private System.Windows.Forms.DataGridViewTextBoxColumn PivotTable;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtWorksheet;
        private System.Windows.Forms.DataGridViewButtonColumn PvtGoTo;
        private System.Windows.Forms.DataGridViewComboBoxColumn PvtGrouping;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtDataSourceName;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtDataSourceType;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtDataSourceDesc;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtLastRefreshed;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtPageFields;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtColumnFields;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtRowFields;
        private System.Windows.Forms.DataGridViewTextBoxColumn PvtDataFields;
    }
}