using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

using System.Diagnostics;
using System.IO;

namespace AAExcelAddIn
{
    public partial class PvtLstObjNavigator : Form
    {
        public PvtLstObjNavigator()
        {
            InitializeComponent();
        }

        //Global variables
        //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        ///<summary>This is the custom xml part that drives the navigator.</summary>
        public Office.CustomXMLPart addInXmlPart;
        ///<summary>Indicates whether a new grouping being created in the Grouping tab or not.</summary>
        public bool creatingNewGrouping;
        ///<summary>Indicates if the current record in the groupings data grid in the new record row.</summary>
        public bool newRecordRow;
        ///<summary>Stores what the grouping was named before it was updated in the Groupings tab.</summary>
        public string previousGrouping;
        ///<summary>Stores what the previous string value of a data grid cell was before it was updated.</summary>
        public string previousDataGridStringValue;
        ///<summary>The index of the row that was previously selected when the row selection changed in the data source data grid.</summary>
        public int previousDataSrcRowIndex = -1;
        //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

        private void PvtLstObjNavigator_Load(object sender, EventArgs e)
        {
            
            //Variables
            Excel.Application app;
            Excel.Workbook thisWorkbook;
            Office.DocumentProperty customXmlPartDocProp;
            string dataSoruceName = "", dataSrouceType = "", dataSoruceDesc = "", pageFields, rowFields, columnFields, dataFields, lstObjColumns, connType, connCommandText, connFilePath, connCommandType, connLastRefreshed, objectGrouping, modifiedObjectName;
            const string xmlPartTitle = "<title>AA Excel Add-In (Navigator)</title>", docPropertyName = "NavCustomXmlPartID";
            Excel.XlCmdType cmdType = Microsoft.Office.Interop.Excel.XlCmdType.xlCmdDefault;
            bool connPivotCache, connReadOnly, correctXmlPart = false;
            Nullable<decimal> connPvtChcSize = null;
            DataTable dtPivots = new DataTable(), dtLstObj = new DataTable(), dtWbConn = new DataTable();

            //Creating the activeworkbook object
            app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            app.Visible = true;
            thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;
            DataSet ds = new DataSet();

            //Grabbing the custom Xml Part
            //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

            //Making sure the document property that contains the xml part id exists in the current workbook
            //If the property doesn't exist, create/re-create it
            try
            {
                customXmlPartDocProp = thisWorkbook.CustomDocumentProperties(docPropertyName);
                addInXmlPart = thisWorkbook.CustomXMLParts.SelectByID(customXmlPartDocProp.Value);
                if (addInXmlPart != null)
                {
                    correctXmlPart = addInXmlPart.XML.Contains(xmlPartTitle);
                }
            }
            catch
            {
                customXmlPartDocProp = thisWorkbook.CustomDocumentProperties.Add(Name: docPropertyName, LinkToContent: false, Type: Office.MsoDocProperties.msoPropertyTypeString, Value: "0");
            }

            //If the part was not properly obtained, take the necessary steps to fix the issue
            if (addInXmlPart == null || correctXmlPart == false)
            {

                //Grabbing the id of the xml part
                customXmlPartDocProp.Value = "0";
                foreach (Office.CustomXMLPart xmlPart in thisWorkbook.CustomXMLParts)
                {
                    if (xmlPart.XML.Contains(xmlPartTitle))
                    {
                        addInXmlPart = xmlPart;
                        customXmlPartDocProp.Value = addInXmlPart.Id;
                        break;
                    }
                }

                //If the id was not found in the loop, create the xml part
                if (customXmlPartDocProp.Value == "0")
                {
                    addInXmlPart = thisWorkbook.CustomXMLParts.Add("<data>" + xmlPartTitle + "<Groupings></Groupings><PivotGroupings></PivotGroupings><ListObjectGroupings></ListObjectGroupings><ConnectionGroupings></ConnectionGroupings></data>");
                    customXmlPartDocProp.Value = addInXmlPart.Id;
                }
            }
            //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

            //Groupings
            //>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

            //Loading the groupings into the data grid in its tab
            StringReader sr = new StringReader(addInXmlPart.SelectSingleNode("data/Groupings").XML);
            ds.ReadXml(sr);
            if (ds.Tables.Count > 0)
            {
                dgrGroupings.DataSource = ds.Tables[0];
                dgrGroupings.Columns[0].Name = "Grouping";
                dgrGroupings.Columns[0].HeaderText = "Grouping";
            }
            else
            {
                dgrGroupings.Columns.Add("Grouping", "Grouping");
            }

            //Loading all the current groupings into the grouping dropdowns
            cboGroupingFilter.Items.Add("");
            ((DataGridViewComboBoxColumn)dgrPivotTables.Columns["pvtGrouping"]).Items.Add("");
            ((DataGridViewComboBoxColumn)dgrListObjects.Columns["lstObjGrouping"]).Items.Add("");
            ((DataGridViewComboBoxColumn)dgrWbConnections.Columns["dtaSrcGrouping"]).Items.Add("");
            foreach (Office.CustomXMLNode node in addInXmlPart.SelectSingleNode("data/Groupings").ChildNodes)
            {
                cboGroupingFilter.Items.Add(node.Text);
                ((DataGridViewComboBoxColumn)dgrPivotTables.Columns["pvtGrouping"]).Items.Add(node.Text);
                ((DataGridViewComboBoxColumn)dgrListObjects.Columns["lstObjGrouping"]).Items.Add(node.Text);
                ((DataGridViewComboBoxColumn)dgrWbConnections.Columns["dtaSrcGrouping"]).Items.Add(node.Text);
            }

            //>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

            //Creating DataTables for grid view's data sources
            //************************************************************************************************

            //PivotTables gird view
            foreach (DataGridViewColumn col in dgrPivotTables.Columns)
            {
                dtPivots.Columns.Add(col.Name);
                col.DataPropertyName = col.Name;
            }

            //List Objects gird view
            foreach (DataGridViewColumn col in dgrListObjects.Columns)
            {
                dtLstObj.Columns.Add(col.Name);
                col.DataPropertyName = col.Name;
            }

            //Workbook Connections gird view
            foreach (DataGridViewColumn col in dgrWbConnections.Columns)
            {
                dtWbConn.Columns.Add(col.Name);
                col.DataPropertyName = col.Name;
            }
            //************************************************************************************************


            //Loading the data grids in the form
            foreach (Excel.Worksheet ws in thisWorkbook.Worksheets)
            {

                //Don't load objects if they reside on a very hidden worksheet
                if (ws.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
                {

                    //Loading the PivotTables tab data grid
                    foreach (Excel.PivotTable pvt in ws.PivotTables())
                    {

                        //Determining where the data for the pivot table is being pulled from
                        Excel.PivotCache pvtCache = pvt.PivotCache();
                        switch (pvt.PivotCache().SourceType)
                        {

                            //Excel Table/Range
                            case Excel.XlPivotTableSourceType.xlDatabase:
                                dataSoruceName = pvtCache.SourceData;
                                dataSrouceType = "Excel Table/Range";
                                dataSoruceDesc = "";
                                break;

                            //Workbook Connection
                            case Excel.XlPivotTableSourceType.xlExternal:
                                dataSoruceName = pvtCache.WorkbookConnection.Name;
                                dataSrouceType = "Workbook Connection";
                                dataSoruceDesc = pvtCache.WorkbookConnection.Description;
                                break;

                            //Unhandled data source
                            default:
                                dataSoruceName = "";
                                dataSrouceType = "Unknown";
                                dataSoruceDesc = "";
                                break;
                        }

                        //Creating concatenations of all the fields in the pivot
                        //-----------------------------------------------------------------------------------

                        //Filter fields
                        pageFields = "";
                        foreach (Excel.PivotField pvtField in pvt.PageFields)
                        {
                            pageFields += (pageFields == "") ? pvtField.Name : "; " + pvtField.Name;
                        }

                        //Row fields
                        rowFields = "";
                        foreach (Excel.PivotField pvtField in pvt.RowFields)
                        {
                            rowFields += (rowFields == "") ? pvtField.Name : "; " + pvtField.Name;
                        }

                        //Column fields
                        columnFields = "";
                        foreach (Excel.PivotField pvtField in pvt.ColumnFields)
                        {
                            columnFields += (columnFields == "") ? pvtField.Name : "; " + pvtField.Name;
                        }

                        //Value fields
                        dataFields = "";
                        foreach (Excel.PivotField pvtField in pvt.DataFields)
                        {
                            dataFields += (dataFields == "") ? pvtField.Name : "; " + pvtField.Name;
                        }
                        //-----------------------------------------------------------------------------------

                        //Checking if a grouping was assigned to the pivot or not
                        //...................................................................................

                        modifiedObjectName = pvt.Name.Replace("'", "&apos;");
                        if (addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[@pivotName='" + modifiedObjectName + "'][@worksheetName='" + ws.Name + "'][@pivotType='Table']") != null) {
                            objectGrouping = addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[@pivotName='" + modifiedObjectName + "'][@worksheetName='" + ws.Name + "'][@pivotType='Table']").Attributes[4].NodeValue.ToString();
                        }
                        else
                        {
                            objectGrouping = "";
                        }

                        //...................................................................................

                        //Creating a new row in the data grid for each pivot
                        dtPivots.Rows.Add(new object[] { pvt.Name, ws.Name, "Go To", objectGrouping, false, dataSoruceName, dataSrouceType, dataSoruceDesc, pvt.RefreshDate, pageFields, rowFields, columnFields, dataFields });
                    }

                    //Loading the List Objects tab data grid
                    foreach (Excel.ListObject lst in ws.ListObjects)
                    {

                        //Determining where the data for the pivot table is being pulled from
                        switch (lst.SourceType)
                        {

                            //Excel Table/Range
                            case Excel.XlListObjectSourceType.xlSrcRange:
                                dataSrouceType = "No External Source";
                                break;

                            //Workbook Connection
                            case Excel.XlListObjectSourceType.xlSrcQuery:
                                dataSoruceName = lst.QueryTable.WorkbookConnection.Name;
                                dataSrouceType = "Workbook Connection";
                                dataSoruceDesc = lst.QueryTable.WorkbookConnection.Description;
                                break;

                            //Unhandled data soruce
                            default:
                                dataSoruceName = "";
                                dataSrouceType = "Unknown";
                                dataSoruceDesc = "";
                                break;
                        }

                        //Create a concatenation for each column in the list object
                        lstObjColumns = "";
                        foreach (Excel.ListColumn lstColumn in lst.ListColumns)
                        {
                            lstObjColumns += (lstObjColumns == "") ? lstColumn.Name : "; " + lstColumn.Name;
                        }


                        //Checking if a grouping was assigned to the pivot or not
                        //...................................................................................

                        modifiedObjectName = lst.Name.Replace("'", "&apos;");
                        if (addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[@lstObjName='" + modifiedObjectName + "'][@worksheetName='" + ws.Name + "']") != null)
                        {
                            objectGrouping = addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[@lstObjName='" + modifiedObjectName + "'][@worksheetName='" + ws.Name + "']").Attributes[3].NodeValue.ToString();
                        }
                        else
                        {
                            objectGrouping = "";
                        }

                        //...................................................................................

                        //Creating a new row in the data grid for each list object
                        dtLstObj.Rows.Add(new object[] { lst.Name, ws.Name, "Go To", objectGrouping, dataSoruceName, dataSrouceType, dataSoruceDesc, lstObjColumns });
                    }
                }
            }

            //Loading each workbook connection into the data sources tab data grid
            foreach (Excel.WorkbookConnection conn in thisWorkbook.Connections)
            {

                //Grabbing the variables
                switch (conn.Type)
                {
                    
                    //Data Feed
                    case Excel.XlConnectionType.xlConnectionTypeDATAFEED:
                        connType = "Data Feed";

                        try
                        {
                            connLastRefreshed = conn.DataFeedConnection.RefreshDate.ToString();
                        }
                        catch
                        {
                            connLastRefreshed = "";
                        }

                        connReadOnly = (conn.DataFeedConnection.Connection.IndexOf("Mode=Read") != -1 || conn.OLEDBConnection.Connection.IndexOf("ReadOnly=True") != -1);
                        connCommandText = conn.DataFeedConnection.CommandText;
                        connFilePath = conn.DataFeedConnection.SourceConnectionFile;
                        cmdType = conn.DataFeedConnection.CommandType;
                        break;

                    //Power Pivot Model
                    case Excel.XlConnectionType.xlConnectionTypeMODEL:
                        connType = "Power Pivot Model";
                        connReadOnly = false;
                        connLastRefreshed = "";
                        connCommandText = conn.ModelConnection.CommandText;
                        connFilePath = "";
                        cmdType = conn.ModelConnection.CommandType;
                        break;

                    //No Source
                    case Excel.XlConnectionType.xlConnectionTypeNOSOURCE:
                        connType = "No Source";
                        connReadOnly = false;
                        connLastRefreshed = "";
                        connCommandText = "";
                        connFilePath = "";
                        break;

                    //ODBC
                    case Excel.XlConnectionType.xlConnectionTypeODBC:
                        connType = "ODBC";

                        try
                        {
                            connLastRefreshed = conn.ODBCConnection.RefreshDate.ToString();
                        }
                        catch
                        {
                            connLastRefreshed = "";
                        }

                        connReadOnly = (conn.ODBCConnection.Connection.IndexOf("Mode=Read") != -1 || conn.OLEDBConnection.Connection.IndexOf("ReadOnly=True") != -1);
                        connCommandText = conn.ODBCConnection.CommandText;
                        connFilePath = conn.ODBCConnection.SourceDataFile;
                        cmdType = conn.ODBCConnection.CommandType;
                        break;

                    //OLEDB
                    case Excel.XlConnectionType.xlConnectionTypeOLEDB:
                        connType = "OLEDB";

                        try
                        {
                            connLastRefreshed = conn.OLEDBConnection.RefreshDate.ToString();
                        }
                        catch
                        {
                            connLastRefreshed = "";
                        }

                        connReadOnly = (conn.OLEDBConnection.Connection.IndexOf("Mode=Read") != -1 || conn.OLEDBConnection.Connection.IndexOf("ReadOnly=True") != -1);
                        connCommandText = conn.OLEDBConnection.CommandText;
                        connFilePath = conn.OLEDBConnection.SourceDataFile;
                        cmdType = conn.OLEDBConnection.CommandType;
                        break;

                    //Text
                    case Excel.XlConnectionType.xlConnectionTypeTEXT:
                        connType = "Text";
                        connLastRefreshed = "";
                        connReadOnly = false;
                        connCommandText = "";
                        connFilePath = conn.TextConnection.Connection.Substring(conn.TextConnection.Connection.IndexOf(';') + 1);
                        break;

                    //Web
                    case Excel.XlConnectionType.xlConnectionTypeWEB:
                        connType = "Web";
                        connLastRefreshed = "";
                        connReadOnly = false;
                        connCommandText = "";
                        connFilePath = "";
                        break;

                    //Worksheet
                    case Excel.XlConnectionType.xlConnectionTypeWORKSHEET:
                        connType = "Worksheet";
                        connLastRefreshed = "";
                        connReadOnly = false;
                        connCommandText = conn.WorksheetDataConnection.CommandText;
                        connFilePath = "";
                        cmdType = conn.OLEDBConnection.CommandType;
                        break;

                    //XML Map
                    case Excel.XlConnectionType.xlConnectionTypeXMLMAP:
                        connType = "XML Map";
                        connLastRefreshed = "";
                        connReadOnly = false;
                        connCommandText = "";
                        connFilePath = "";
                        break;

                    //Unknown
                    default:
                        connType = "Unknown";
                        connLastRefreshed = "";
                        connReadOnly = false;
                        connCommandText = "";
                        connFilePath = "";
                        break;
                }

                //Determining the command type of the connection is the connection type is known
                if (connType != "Unknown" && connType != "No Source" && connType != "Text" && connType != "Web" && connType != "XML Map")
                {

                    switch (cmdType)
                    {
                        case Excel.XlCmdType.xlCmdCube:
                            connCommandType = "Cube Name for OLAP Data Source";
                            break;
                        case Excel.XlCmdType.xlCmdDAX:
                            connCommandType = "Data Analysis Expressions Formula";
                            break;
                        case Excel.XlCmdType.xlCmdDefault:
                            connCommandType = "Default";
                            break;
                        case Excel.XlCmdType.xlCmdExcel:
                            connCommandType = "Excel Formula";
                            break;
                        case Excel.XlCmdType.xlCmdList:
                            connCommandType = "List";
                            break;
                        case Excel.XlCmdType.xlCmdSql:
                            connCommandType = "SQL Statement";
                            break;
                        case Excel.XlCmdType.xlCmdTable:
                            connCommandType = "OLE DB Data Source Table";
                            break;
                        case Excel.XlCmdType.xlCmdTableCollection:
                            connCommandType = "Table Collection";
                            break;
                        default:
                            connCommandType = "Unknown";
                            break;
                    }
                }
                else
                {
                    connCommandType = "Unknown";
                }

                //Determining if the workbook connection is linked to a pivot cache or not
                connPivotCache = false;
                connPvtChcSize = null;
                foreach (Excel.PivotCache pvtCache in thisWorkbook.PivotCaches())
                {

                    try
                    {
                        if (pvtCache.WorkbookConnection.Name == conn.Name)
                        {
                            connPivotCache = true;
                            connPvtChcSize = (Convert.ToDecimal(pvtCache.MemoryUsed) / 1048576);
                            connPvtChcSize = Decimal.Round(Convert.ToDecimal(connPvtChcSize), 2);
                            break;
                        }
                    }
                    catch
                    { }
                }

                //Checking if a grouping was assigned to the connection or not
                //...................................................................................

                modifiedObjectName = conn.Name.Replace("'", "&apos;");
                if (addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[@connectionName='" + modifiedObjectName + "']") != null)
                {
                    objectGrouping = addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[@connectionName='" + modifiedObjectName + "']").Attributes[2].NodeValue.ToString();
                }
                else
                {
                    objectGrouping = "";
                }
                //...................................................................................

                //Creating a new row in the data grid for each list object
                dtWbConn.Rows.Add(new object[] { conn.Name, conn.Description, connType, objectGrouping, connPivotCache, connPvtChcSize, connLastRefreshed, connReadOnly, connCommandText, connFilePath, connCommandType });
            }

            //Loading the data tables in the gird views
            dgrPivotTables.DataSource = dtPivots;
            dgrListObjects.DataSource = dtLstObj;
            dgrWbConnections.DataSource = dtWbConn;

            //Right aligning certain columns in the data source grid view
            dgrWbConnections.Columns["dtaSrcPvtChcMemory"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            //Unselecting the first row in the data grids
            dgrPivotTables.ClearSelection();
            dgrListObjects.ClearSelection();
            dgrWbConnections.ClearSelection();

        }

        //User selects a data source record that is a pivot cache to see its fields
        private void dgrWbConnections_SelectionChanged(object sender, EventArgs e)
        {

            //Variables
            Excel.Application app;
            Excel.Workbook thisWorkbook;
            bool cacheFieldsFound = false;
            string pvtFieldDataType;
            DataGridViewColumn sortedColumn;

            //If the currently selected row did not change, do not reload the pivot cache fields grid
            if (previousDataSrcRowIndex != dgrWbConnections.CurrentCell.RowIndex)
            {

                //Creating the activeworkbook object
                app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                app.Visible = true;
                thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;

                //Restting the sorting in the sub grid
                if (dgrPvtChcFields.SortedColumn != null)
                {
                    sortedColumn = dgrPvtChcFields.Columns[dgrPvtChcFields.SortedColumn.Index];
                    sortedColumn.SortMode = DataGridViewColumnSortMode.NotSortable;
                    sortedColumn.SortMode = DataGridViewColumnSortMode.Automatic;
                }

                //Resetting the fields sub grid if needed
                if (dgrPvtChcFields.RowCount > 0)
                {
                    dgrPvtChcFields.Rows.Clear();
                    dgrPvtChcFields.Refresh();
                }

                //Only check for piovt cache fields if the data source is linked to a pivot cache
                if (Convert.ToBoolean(dgrWbConnections.Rows[dgrWbConnections.CurrentCell.RowIndex].Cells[4].Value) == true)
                {

                    //Looping through each pivot cache to see if the selected connection is linked to any of them
                    foreach (Excel.PivotCache pvtCache in thisWorkbook.PivotCaches())
                    {

                        //Handle scenario if a pivot cache is not linked to a workbook connection
                        try
                        {

                            //If the workbook connection's name matches the name of the connection linked to the current pivot cache, grab its fields
                            if (pvtCache.WorkbookConnection.Name == dgrWbConnections.Rows[dgrWbConnections.CurrentCell.RowIndex].Cells[0].Value.ToString())
                            {

                                //Loop through each pivot in the workbook until one is found that was created from the current pivot cache
                                foreach (Excel.Worksheet ws in thisWorkbook.Worksheets)
                                {
                                    foreach (Excel.PivotTable pvt in ws.PivotTables())
                                    {

                                        //If the current pivot table pulls its data from the pivot cache, grab its fields and load them into the data grid view
                                        if (pvt.PivotCache().Index == pvtCache.Index)
                                        {
                                            foreach (Excel.PivotField pvtField in pvt.PivotFields())
                                            {

                                                //Determining the current field's data type
                                                switch (pvtField.DataType)
                                                {

                                                    case Excel.XlPivotFieldDataType.xlDate:
                                                        pvtFieldDataType = "Date/Time";
                                                        break;

                                                    case Excel.XlPivotFieldDataType.xlNumber:
                                                        pvtFieldDataType = "Number/Boolean";
                                                        break;

                                                    case Excel.XlPivotFieldDataType.xlText:
                                                        pvtFieldDataType = "Text";
                                                        break;

                                                    default:
                                                        pvtFieldDataType = "Unknown";
                                                        break;
                                                }

                                                //Creating a new row in the data grid for each list object
                                                DataGridViewRow row = new DataGridViewRow();
                                                row.CreateCells(dgrPvtChcFields, pvtField.SourceName, pvtFieldDataType);
                                                dgrPvtChcFields.Rows.Add(row);
                                            }
                                            cacheFieldsFound = true;
                                            break;
                                        }
                                    }

                                    //If the related pivot cache was found, exit the loop
                                    if (cacheFieldsFound == true)
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }

                //Storing that this was last row selected
                previousDataSrcRowIndex = dgrWbConnections.CurrentCell.RowIndex;
            }
        }

        //Updating the xml part based on whether a new grouping is being created or an existing grouping is being updated
        private void dgrGroupings_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            //Making sure the current record isnt the new record
            if (!newRecordRow)
            {

                //Determing if the edit was made for a new row or existing
                if (creatingNewGrouping)
                {
                    addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/Groupings"), "Grouping", NodeValue: dgrGroupings[e.ColumnIndex, e.RowIndex].Value.ToString());
                    cboGroupingFilter.Items.Add(dgrGroupings[e.ColumnIndex, e.RowIndex].Value.ToString());
                    ((DataGridViewComboBoxColumn)dgrPivotTables.Columns["pvtGrouping"]).Items.Add(dgrGroupings[e.ColumnIndex, e.RowIndex].Value.ToString());
                    ((DataGridViewComboBoxColumn)dgrListObjects.Columns["lstObjGrouping"]).Items.Add(dgrGroupings[e.ColumnIndex, e.RowIndex].Value.ToString());
                    ((DataGridViewComboBoxColumn)dgrWbConnections.Columns["dtaSrcGrouping"]).Items.Add(dgrGroupings[e.ColumnIndex, e.RowIndex].Value.ToString());
                }
                else
                {

                    //Updating the grouping in the xml part
                    addInXmlPart.SelectSingleNode("data/Groupings/Grouping[text()=\"" + previousGrouping + "\"]").Text = dgrGroupings[e.ColumnIndex, e.RowIndex].Value.ToString();

                    //Updating the grouping in the combobox filter
                    for (int i = 0; i < cboGroupingFilter.Items.Count; i++)
                    {
                        if (cboGroupingFilter.GetItemText(cboGroupingFilter.Items[i]) == previousGrouping)
                        {
                            cboGroupingFilter.Items.Remove(cboGroupingFilter.Items[i]);
                            break;
                        }
                    }
                    cboGroupingFilter.Items.Add(dgrGroupings[0, e.RowIndex].Value.ToString());

                    //Clearing the filters from all the grids
                    if (!String.IsNullOrEmpty(cboGroupingFilter.Text))
                    {
                        cboGroupingFilter.Text = "";
                    }
                    (dgrPivotTables.DataSource as DataTable).DefaultView.RowFilter = "";
                    (dgrListObjects.DataSource as DataTable).DefaultView.RowFilter = "";
                    (dgrWbConnections.DataSource as DataTable).DefaultView.RowFilter = "";

                    //Updating the grouping assigned to the records in the data grid views
                    //--------------------------------------------------------------------------------------------------

                    //PivotTables
                    foreach (DataGridViewRow row in dgrPivotTables.Rows)
                    {
                        if (row.Cells[3].Value != null)
                        {
                            if (row.Cells[3].Value.ToString() == previousGrouping)
                            {
                                row.Cells[3].Value = dgrGroupings[0, e.RowIndex].Value.ToString();
                            }
                        }
                    }
                    ((DataGridViewComboBoxColumn)dgrPivotTables.Columns["pvtGrouping"]).Items.Remove(previousGrouping);
                    ((DataGridViewComboBoxColumn)dgrPivotTables.Columns["pvtGrouping"]).Items.Add(dgrGroupings[e.ColumnIndex, e.RowIndex].Value.ToString());

                    //List Objects
                    foreach (DataGridViewRow row in dgrListObjects.Rows)
                    {
                        if (row.Cells[3].Value != null)
                        {
                            if (row.Cells[3].Value.ToString() == previousGrouping)
                            {
                                row.Cells[3].Value = dgrGroupings[0, e.RowIndex].Value.ToString();
                            }
                        }   
                    }
                    ((DataGridViewComboBoxColumn)dgrListObjects.Columns["lstObjGrouping"]).Items.Remove(previousGrouping);
                    ((DataGridViewComboBoxColumn)dgrListObjects.Columns["lstObjGrouping"]).Items.Add(dgrGroupings[e.ColumnIndex, e.RowIndex].Value.ToString());

                    //Data Sources
                    foreach (DataGridViewRow row in dgrWbConnections.Rows)
                    {
                        if (row.Cells[3].Value != null)
                        {
                            if (row.Cells[3].Value.ToString() == previousGrouping)
                            {
                                row.Cells[3].Value = dgrGroupings[0, e.RowIndex].Value.ToString();
                            }
                        }
                    }
                    ((DataGridViewComboBoxColumn)dgrWbConnections.Columns["dtaSrcGrouping"]).Items.Remove(previousGrouping);
                    ((DataGridViewComboBoxColumn)dgrWbConnections.Columns["dtaSrcGrouping"]).Items.Add(dgrGroupings[e.ColumnIndex, e.RowIndex].Value.ToString());

                    //--------------------------------------------------------------------------------------------------

                    //Updating the groupings assigned to objects in the xml part
                    foreach (Office.CustomXMLNode xmlNode in addInXmlPart.SelectNodes("data/PivotGroupings/PivotGrouping[@grouping=\"" + previousGrouping + "\"]"))
                    {
                        xmlNode.Attributes[4].NodeValue = dgrGroupings[e.ColumnIndex, e.RowIndex].Value.ToString();
                    }
                }
            }
        }

        //If the user deletes the grouping from the grid, delete it in the xml part
        private void dgrGroupings_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {

            DialogResult msgBoxResult;

            newRecordRow = (dgrGroupings.NewRowIndex == e.Row.Index);
            if (!newRecordRow)
            {

                //Confirming with the user that they wish to delete the grouping
                msgBoxResult = MessageBox.Show("Are you sure you want to delete this grouping? If it is assigned to a PivotTable or Table, the object(s) will no longer be assigned to a grouping. This action cannot be undone.", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                e.Cancel = (msgBoxResult == DialogResult.No);
                if (!e.Cancel)
                {

                    //Removing grouping from custom xml part
                    addInXmlPart.SelectSingleNode("data/Groupings/Grouping[text()=\"" + e.Row.Cells[0].Value.ToString() + "\"]").Delete();

                    //Removing the grouping in the combobox filter
                    for (int i = 0; i < cboGroupingFilter.Items.Count; i++)
                    {
                        if (cboGroupingFilter.GetItemText(cboGroupingFilter.Items[i]) == e.Row.Cells[0].Value.ToString())
                        {
                            cboGroupingFilter.Items.Remove(cboGroupingFilter.Items[i]);
                            break;
                        }
                    }

                    //Clearing the filters from all the grids
                    if (!String.IsNullOrEmpty(cboGroupingFilter.Text))
                    {
                        cboGroupingFilter.Text = "";
                    }
                    (dgrPivotTables.DataSource as DataTable).DefaultView.RowFilter = "";
                    (dgrListObjects.DataSource as DataTable).DefaultView.RowFilter = "";
                    (dgrWbConnections.DataSource as DataTable).DefaultView.RowFilter = "";

                    //Removing grouping from comboboxes in data grid views
                    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                    //PivotTables
                    foreach (DataGridViewRow row in dgrPivotTables.Rows)
                    {
                        if (row.Cells[3].Value != null)
                        {
                            if (row.Cells[3].Value.ToString() == dgrGroupings[0, e.Row.Index].Value.ToString())
                            {
                                row.Cells[3].Value = "";
                            }
                        }
                    }
                    ((DataGridViewComboBoxColumn)dgrPivotTables.Columns["pvtGrouping"]).Items.Remove(dgrGroupings[0, e.Row.Index].Value.ToString());

                    //List Objects
                    foreach (DataGridViewRow row in dgrListObjects.Rows)
                    {
                        if (row.Cells[3].Value != null)
                        {
                            if (row.Cells[3].Value.ToString() == dgrGroupings[0, e.Row.Index].Value.ToString())
                            {
                                row.Cells[3].Value = "";
                            }
                        }
                    }
                    ((DataGridViewComboBoxColumn)dgrListObjects.Columns["lstObjGrouping"]).Items.Remove(dgrGroupings[0, e.Row.Index].Value.ToString());

                    //Data Sources
                    foreach (DataGridViewRow row in dgrWbConnections.Rows)
                    {
                        if (row.Cells[3].Value != null)
                        {
                            if (row.Cells[3].Value.ToString() == dgrGroupings[0, e.Row.Index].Value.ToString())
                            {
                                row.Cells[3].Value = "";
                            }
                        }
                    }
                    ((DataGridViewComboBoxColumn)dgrWbConnections.Columns["dtaSrcGrouping"]).Items.Remove(dgrGroupings[0, e.Row.Index].Value.ToString());
                    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                    //Removing the grouping assigned to objects in the xml part
                    //>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

                    //Pivot groupings
                    foreach (Office.CustomXMLNode xmlNode in addInXmlPart.SelectNodes("data/PivotGroupings/PivotGrouping[@grouping=\"" + e.Row.Cells[0].Value.ToString() + "\"]"))
                    {
                        xmlNode.Delete();
                    }

                    //List Object groupings
                    foreach (Office.CustomXMLNode xmlNode in addInXmlPart.SelectNodes("data/ListObjectGroupings/ListObjectGrouping[@grouping=\"" + e.Row.Cells[0].Value.ToString() + "\"]"))
                    {
                        xmlNode.Delete();
                    }

                    //List Object groupings
                    foreach (Office.CustomXMLNode xmlNode in addInXmlPart.SelectNodes("data/ConnectionGroupings/ConnectionGrouping[@grouping=\"" + e.Row.Cells[0].Value.ToString() + "\"]"))
                    {
                        xmlNode.Delete();
                    }
                    //>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                }
            }
        }

        //Indicates whether if the user is creating a new grouping or not
        private void dgrGroupings_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            creatingNewGrouping = dgrGroupings.Rows[e.RowIndex].IsNewRow;
            if (!creatingNewGrouping)
            {
                previousGrouping = dgrGroupings[0, e.RowIndex].Value.ToString();
            }
        }

        //User clicks the Go To button in the pivot data grid
        private void dgrPivotTables_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            //Variables
            Excel.Application app;
            Excel.Workbook thisWorkbook;

            //Making sure the clicked column is the one with the button and the clicked row isn't the header
            if (e.ColumnIndex == 2 && e.RowIndex != -1)
            {

                //Creating the activeworkbook object
                app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                app.Visible = true;
                thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;

                //Determining where the pivottable is
                Excel.Worksheet ws = thisWorkbook.Sheets[dgrPivotTables.Rows[e.RowIndex].Cells[1].Value];
                Excel.PivotTable pvt = ws.PivotTables(dgrPivotTables.Rows[e.RowIndex].Cells[0].Value);
                Excel.Range rng = ws.Range[pvt.TableRange2.Address.Substring(0, pvt.TableRange2.Address.IndexOf(':'))];

                //Checking if the worksheet the pivot resides in is hidden
                if (ws.Visible == Excel.XlSheetVisibility.xlSheetHidden)
                {

                    //Confirming with the user that they wish to unhide the sheet the pivot resides on
                    DialogResult msgboxResult = MessageBox.Show("The worksheet this PivotTable resides in is currently hidden. Do you wish the worksheet?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    //Unhiding the worksheet if the user confirmed it
                    if (msgboxResult == DialogResult.Yes)
                    {
                        ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    }
                }

                //Moving the cell selector to the pivot
                if (ws.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                {
                    ws.Select();
                    rng.Select();
                }
            }
        }

        //User changes data in the pivots grid view
        private void dgrPivotTables_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            //Variables
            string modifiedPrevObjName, modifiedObjectName;
            Excel.Application app;
            Excel.Workbook thisWorkbook;
            Excel.Worksheet ws;
            Excel.PivotTable pvt;

            //Run a different procedure based on which cell was changed
            switch (e.ColumnIndex)
            {

                //Pivot Name
                case 0:

                    //Creating the activeworkbook object
                    app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    app.Visible = true;
                    thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;

                    //Renaming the pivot table
                    ws = thisWorkbook.Worksheets[dgrPivotTables[1, e.RowIndex].Value.ToString()];
                    pvt = ws.PivotTables(previousDataGridStringValue);
                    pvt.Name = dgrPivotTables[0, e.RowIndex].Value.ToString();

                    //Updating the custom xml part of the pivot table grouping if there is one
                    modifiedPrevObjName = previousDataGridStringValue.ToString().Replace("'", "&apos;");
                    modifiedObjectName = pvt.Name.ToString().Replace("'", "&apos;");
                    if (addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[@pivotName = '" + modifiedPrevObjName + "'][@worksheetName='" + dgrPivotTables[1, e.RowIndex].Value.ToString() + "'][@pivotType='Table']") != null)
                    {

                        //Updating the grouping record that is tied to the update pivot
                        addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[@pivotName = '" + modifiedPrevObjName + "'][@worksheetName='" + dgrPivotTables[1, e.RowIndex].Value.ToString() + "'][@pivotType='Table']").Attributes[1].NodeValue = modifiedObjectName;

                        //Moving the attributes back to their original locations
                        //I mean, really? Just updating the value of attribute moves it to the last attribute index of the element?! This hack will have to do I guess
                        for (int i = 1; i <= 3; i++)
                        {
                            addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[@pivotName = '" + modifiedObjectName + "'][@worksheetName='" + dgrPivotTables[1, e.RowIndex].Value.ToString() + "'][@pivotType='Table']").Attributes[1].NodeValue = addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[@pivotName = '" + modifiedObjectName + "'][@worksheetName='" + dgrPivotTables[1, e.RowIndex].Value.ToString() + "'][@pivotType='Table']").Attributes[1].NodeValue;
                        }
                    }
                    break;

                //Grouping dropdown
                case 3:


                    //If the user selects the blank option in the dropdown and there already isnt a gropuing record for the selected pivot, do nothing
                    modifiedObjectName = dgrPivotTables[0, e.RowIndex].Value.ToString().Replace("'", "&apos;");
                    if (dgrPivotTables[e.ColumnIndex, e.RowIndex].Value != null || addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[@pivotName='" + modifiedObjectName + "'][@worksheetName='" + dgrPivotTables[1, e.RowIndex].Value.ToString() + "'][@pivotType='Table']") != null)
                    {

                        //If the pivot, worksheet, and type key exists, update the grouping of the key
                        //Otherwise, create a new record for the pivot grouping
                        if (addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[@pivotName='" + modifiedObjectName + "'][@worksheetName='" + dgrPivotTables[1, e.RowIndex].Value.ToString() + "'][@pivotType='Table']") != null)
                        {

                            //If the user chooses the blank option, then just delete the record in the xml part
                            //Otherwise, update the record to the newly selected grouping
                            if (dgrPivotTables[e.ColumnIndex, e.RowIndex].Value == null)
                            {
                                addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[@pivotName='" + modifiedObjectName + "'][@worksheetName='" + dgrPivotTables[1, e.RowIndex].Value.ToString() + "'][@pivotType='Table']").Delete();
                            }
                            else
                            {
                                addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[@pivotName='" + modifiedObjectName + "'][@worksheetName='" + dgrPivotTables[1, e.RowIndex].Value.ToString() + "'][@pivotType='Table']").Attributes[4].NodeValue = dgrPivotTables[e.ColumnIndex, e.RowIndex].Value.ToString();
                            }
                        }
                        else
                        {

                            //Creating the new pivot grouping
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/PivotGroupings"), "PivotGrouping");

                            //Creating attributes
                            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++

                            //Pivot Name
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[last()]"), "pivotName", "", NodeType: Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue: modifiedObjectName);

                            //Worksheet Name
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[last()]"), "worksheetName", "", NodeType: Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue: dgrPivotTables[1, e.RowIndex].Value.ToString());

                            //Pivot Type
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[last()]"), "pivotType", "", NodeType: Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue: "Table");

                            //Grouping
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/PivotGroupings/PivotGrouping[last()]"), "grouping", "", NodeType: Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue: dgrPivotTables[e.ColumnIndex, e.RowIndex].Value.ToString());
                            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        }
                    }
                    break;
            }
        }

        //User clicks the Go To button in the list object data grid
        private void dgrListObjects_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
            //Variables
            Excel.Application app;
            Excel.Workbook thisWorkbook;

            //Making sure the double clicked row isn't the header
            if (e.ColumnIndex == 2 && e.RowIndex != -1)
            {

                //Creating the activeworkbook object
                app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                app.Visible = true;
                thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;

                //Determining where the list object is
                Excel.Worksheet ws = thisWorkbook.Sheets[dgrListObjects.Rows[e.RowIndex].Cells[1].Value];
                Excel.ListObject lst = ws.ListObjects[dgrListObjects.Rows[e.RowIndex].Cells[0].Value];
                Excel.Range rng = ws.Range[lst.Range.Address.Substring(0, lst.Range.Address.IndexOf(':'))];

                //Checking if the worksheet the list object resides in is hidden
                if (ws.Visible == Excel.XlSheetVisibility.xlSheetHidden)
                {

                    //Confirming with the user that they wish to unhide the sheet the list object resides on
                    DialogResult msgboxResult = MessageBox.Show("The worksheet this Table resides in is currently hidden. Do you wish the worksheet?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    //Unhiding the worksheet if the user confirmed it
                    if (msgboxResult == DialogResult.Yes)
                    {
                        ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    }
                }

                //Moving the cell selector to the list object
                if (ws.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                {
                    ws.Select();
                    rng.Select();
                }
            }
        }

        //User changes data in the list objects grid view
        private void dgrListObjects_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            //Variables
            string modifiedPrevObjName, modifiedObjectName;
            Excel.Application app;
            Excel.Workbook thisWorkbook;
            Excel.Worksheet ws;
            Excel.ListObject lst;

            //Run a different procedure based on which cell was changed
            switch (e.ColumnIndex)
            {

                //List Object Name
                case 0:

                    //Creating the activeworkbook object
                    app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    app.Visible = true;
                    thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;

                    //Updating the custom xml part of the list object grouping if there is one
                    ws = thisWorkbook.Worksheets[dgrListObjects[1, e.RowIndex].Value.ToString()];
                    lst = ws.ListObjects[dgrListObjects[e.ColumnIndex, e.RowIndex].Value.ToString()];
                    modifiedPrevObjName = previousDataGridStringValue.ToString().Replace("'", "&apos;");
                    modifiedObjectName = lst.Name.ToString().Replace("'", "&apos;");
                    if (addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[@lstObjName = '" + modifiedPrevObjName + "'][@worksheetName='" + dgrListObjects[1, e.RowIndex].Value.ToString() + "']") != null)
                    {

                        //Updating the grouping record that is tied to the update list object
                        addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[@lstObjName = '" + modifiedPrevObjName + "'][@worksheetName='" + dgrListObjects[1, e.RowIndex].Value.ToString() + "']").Attributes[1].NodeValue = modifiedObjectName;

                        //Moving the attributes back to their original locations
                        //I mean, really? Just updating the value of attribute moves it to the last attribute index of the element?! This hack will have to do I guess
                        for (int i = 1; i <= 2; i++)
                        {
                            addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[@lstObjName = '" + modifiedObjectName + "'][@worksheetName='" + dgrListObjects[1, e.RowIndex].Value.ToString() + "']").Attributes[1].NodeValue = addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[@lstObjName = '" + modifiedObjectName + "'][@worksheetName='" + dgrListObjects[1, e.RowIndex].Value.ToString() + "']").Attributes[1].NodeValue;
                        }
                    }
                    break;

                //Grouping dropdown
                case 3:


                    //If the user selects the blank option in the dropdown and there already isnt a gropuing record for the selected list object, do nothing
                    modifiedObjectName = dgrListObjects[0, e.RowIndex].Value.ToString().Replace("'", "&apos;");
                    if (dgrListObjects[e.ColumnIndex, e.RowIndex].Value != null || addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[@lstObjName='" + modifiedObjectName + "'][@worksheetName='" + dgrListObjects[1, e.RowIndex].Value.ToString() + "']") != null)
                    {

                        //If the list object and worksheet key exists, update the grouping of the key
                        //Otherwise, create a new record for the pivot grouping
                        if (addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[@lstObjName='" + modifiedObjectName + "'][@worksheetName='" + dgrListObjects[1, e.RowIndex].Value.ToString() + "']") != null)
                        {

                            //If the user chooses the blank option, then just delete the record in the xml part
                            //Otherwise, update the record to the newly selected grouping
                            if (dgrListObjects[e.ColumnIndex, e.RowIndex].Value == null)
                            {
                                addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[@lstObjName='" + modifiedObjectName + "'][@worksheetName='" + dgrListObjects[1, e.RowIndex].Value.ToString() + "']").Delete();
                            }
                            else
                            {
                                addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[@lstObjName='" + modifiedObjectName + "'][@worksheetName='" + dgrListObjects[1, e.RowIndex].Value.ToString() + "']").Attributes[3].NodeValue = dgrListObjects[e.ColumnIndex, e.RowIndex].Value.ToString();
                            }
                        }
                        else
                        {

                            //Creating the new list object grouping
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/ListObjectGroupings"), "ListObjectGrouping");

                            //Creating attributes
                            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++

                            //List Object Name
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[last()]"), "lstObjName", "", NodeType: Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue: modifiedObjectName);

                            //Worksheet Name
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[last()]"), "worksheetName", "", NodeType: Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue: dgrListObjects[1, e.RowIndex].Value.ToString());

                            //Grouping
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/ListObjectGroupings/ListObjectGrouping[last()]"), "grouping", "", NodeType: Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue: dgrListObjects[e.ColumnIndex, e.RowIndex].Value.ToString());
                            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        }
                    }
                    break;
            }
        }

        //Making sure the user made a valid entry for the editted column in the pivot data grid
        private void dgrPivotTables_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {

            //Variables
            Excel.Application app;
            Excel.Workbook thisWorkbook;
            Excel.Worksheet ws;

            //Determine which column was editted
            switch (e.ColumnIndex)
            {

                //Pivot Name
                case 0:

                    //Making sure the entered pivot name isnt blank
                    if (e.FormattedValue.ToString() == "")
                    {
                        MessageBox.Show("PivotTable names cannot be blank.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        e.Cancel = true;
                    }
                    //Making sure the name of the pivot does not exceed 255 characters
                    else if (e.FormattedValue.ToString().Length > 255)
                    {
                        MessageBox.Show("The max length of a PivotTable name is 255 characters.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        e.Cancel = true;
                    }
                    
                    //Making sure the new name for the pivot isnt already assigned to a different pivot in the worksheet
                    if (!e.Cancel)
                    {

                        //Creating the activeworkbook object
                        app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                        app.Visible = true;
                        thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;
                        ws = thisWorkbook.Worksheets[dgrPivotTables[1, e.RowIndex].Value.ToString()];

                        //Looping through each pivot in the worksheet and making sure the new name wasn't already assigned to one of the pivots
                        foreach (Excel.PivotTable pvt in ws.PivotTables())
                        {

                            if (String.Equals(pvt.Name, e.FormattedValue.ToString(), StringComparison.OrdinalIgnoreCase) && !String.Equals(pvt.Name, dgrPivotTables[e.ColumnIndex, e.RowIndex].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            {
                                MessageBox.Show("The name you entered is already assinged to a different PivotTable in the " + ws.Name + " worksheet.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                e.Cancel = true;
                                break;
                            }
                        }
                    }

                    break;
            }
        }

        //If the user is changing a certain value in the list object grid view, run a certain procedure based on which column is being editted
        private void dgrListObjects_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

            //Determine which column is being editted
            switch (e.ColumnIndex)
            {

                //Pivot Name
                case 0:
                    previousDataGridStringValue = dgrListObjects[e.ColumnIndex, e.RowIndex].Value.ToString();
                    break;
            }
        }

        //Making sure the user made a valid entry for the editted column in the list object data grid
        private void dgrListObjects_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {

            //Variables
            Excel.Application app;
            Excel.Workbook thisWorkbook;
            Excel.Worksheet ws;
            Excel.ListObject lst;
            

            //Determine which column was editted
            switch (e.ColumnIndex)
            {

                //List Object Name
                case 0:

                    //Creating the activeworkbook object
                    app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    app.Visible = true;
                    thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;
                    ws = thisWorkbook.Worksheets[dgrListObjects[1, e.RowIndex].Value.ToString()];
                    lst = ws.ListObjects[dgrListObjects[e.ColumnIndex, e.RowIndex].Value.ToString()];

                    //Setting the name of the list object and it fails, notify the user
                    try
                    {
                        lst.DisplayName = e.FormattedValue.ToString();
                    }
                    catch
                    {

                        MessageBox.Show("The name you entered was not except by Excel. It is possible that is may contain an invalid character or the name you wish to give to the table may already belong to another range currently in the workbook.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        e.Cancel = true;
                    }
                    break;
            }
        }

        //If the user is changing a certain value in the pivot grid view, run a certain procedure based on which column is being editted
        private void dgrPivotTables_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            
            //Determine which column is being editted
            switch (e.ColumnIndex)
            {

                //Pivot Name
                case 0:
                    previousDataGridStringValue = dgrPivotTables[e.ColumnIndex, e.RowIndex].Value.ToString();
                    break;
            }
        }

        //Making sure a valid grouping is entered
        private void dgrGroupings_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {

            //No need to validate if the new row is selected
            newRecordRow = (dgrGroupings.NewRowIndex == e.RowIndex);
            if (newRecordRow) { return; }

            //Making sure a value was entered
            if (e.FormattedValue.ToString().TrimStart().TrimEnd() == "")
            {
                MessageBox.Show("A grouping cannot be blank or only contain spaces.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                e.Cancel = true;
            }
            else if (e.FormattedValue.ToString().Contains("\"") || e.FormattedValue.ToString().Contains("'"))
            {
                MessageBox.Show("Quotes are illegal characters.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                e.Cancel = true;
            }
            else
            {
                foreach (DataGridViewRow rw in dgrGroupings.Rows)
                {
                    if (e.RowIndex != dgrGroupings.NewRowIndex && rw.Cells[0].Value != null)
                    {
                        if (e.RowIndex != rw.Index && e.FormattedValue.ToString() == rw.Cells[0].Value.ToString())
                        {
                            MessageBox.Show("This value already exists as a grouping. Please keep your grouping unique.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            e.Cancel = true;
                            break;
                        }
                    }
                }
            }
        }

        //If the user is changing a certain value in the connection grid view, run a certain procedure based on which column is being editted
        private void dgrWbConnections_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

            //Determine which column is being editted
            switch (e.ColumnIndex)
            {

                //Connection Name
                case 0:
                    previousDataGridStringValue = dgrWbConnections[0, e.RowIndex].Value.ToString();
                    break;
            }
        }

        //Making sure the user made a valid entry for the editted column in the connection data grid
        private void dgrWbConnections_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {

            //Variables
            Excel.Application app;
            Excel.Workbook thisWorkbook;

            //Determine which column was editted
            switch (e.ColumnIndex)
            {

                //Connection Name
                case 0:

                    //Making sure the entered connection name isnt blank
                    if (e.FormattedValue.ToString() == "")
                    {
                        MessageBox.Show("Data source names cannot be blank.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        e.Cancel = true;
                    }
                    //Making sure the name of the connection does not exceed 255 characters
                    else if (e.FormattedValue.ToString().Length > 255)
                    {
                        MessageBox.Show("The max length of a Data source name is 255 characters.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        e.Cancel = true;
                    }

                    //Making sure the new name for the data source isnt already assigned to a different connection
                    if (!e.Cancel)
                    {

                        //Creating the activeworkbook object
                        app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                        app.Visible = true;
                        thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;

                        //Looping through each connection and making sure the new name wasn't already assigned to a different connection
                        foreach (Excel.WorkbookConnection conn in thisWorkbook.Connections)
                        {

                            if (String.Equals(conn.Name, e.FormattedValue.ToString(), StringComparison.OrdinalIgnoreCase) && !String.Equals(conn.Name, dgrWbConnections[e.ColumnIndex, e.RowIndex].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            {
                                MessageBox.Show("The name you entered is already assinged to a different data source.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                e.Cancel = true;
                                break;
                            }
                        }
                    }
                    break;

                //Connection Description
                case 1:

                    //Making sure the description of the connection does not exceed 255 characters
                    if (e.FormattedValue.ToString().Length > 255)
                    {
                        MessageBox.Show("The max length of a Data source description is 255 characters.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        e.Cancel = true;
                    }
                    break;
            }
        }

        //User changes data in the connections grid view
        private void dgrWbConnections_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            //Variables
            string modifiedPrevObjName, modifiedObjectName;
            Excel.Application app;
            Excel.Workbook thisWorkbook;
            Excel.WorkbookConnection conn;

            //Creating the activeworkbook object
            app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            app.Visible = true;
            thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;

            //Run a different procedure based on which cell was changed
            switch (e.ColumnIndex)
            {

                //Connection Name
                case 0:

                    //Renaming the connection
                    conn = thisWorkbook.Connections[previousDataGridStringValue];
                    conn.Name = dgrWbConnections[e.ColumnIndex, e.RowIndex].Value.ToString();

                    //Updating the custom xml part of the connection grouping if there is one
                    modifiedPrevObjName = previousDataGridStringValue.ToString().Replace("'", "&apos;");
                    modifiedObjectName = conn.Name.ToString().Replace("'", "&apos;");
                    if (addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[@connectionName = '" + modifiedPrevObjName + "']") != null)
                    {

                        //Updating the grouping record that is tied to the update connection
                        addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[@connectionName = '" + modifiedPrevObjName + "']").Attributes[1].NodeValue = modifiedObjectName;

                        //Moving the attributes back to their original locations
                        //I mean, really? Just updating the value of attribute moves it to the last attribute index of the element?! This hack will have to do I guess
                        addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[@connectionName = '" + modifiedObjectName + "']").Attributes[1].NodeValue = addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[@connectionName = '" + modifiedObjectName + "']").Attributes[1].NodeValue;
                        
                    }
                    break;

                //Connection Description
                case 1:

                    //Updating the connection description
                    conn = thisWorkbook.Connections[dgrWbConnections[0, e.RowIndex].Value.ToString()];
                    conn.Description = dgrWbConnections[e.ColumnIndex, e.RowIndex].Value.ToString();
                    break;

                //Grouping dropdown
                case 3:


                    //If the user selects the blank option in the dropdown and there already isnt a gropuing record for the selected connection, do nothing
                    modifiedObjectName = dgrWbConnections[0, e.RowIndex].Value.ToString().Replace("'", "&apos;");
                    if (dgrWbConnections[e.ColumnIndex, e.RowIndex].Value != null || addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[@connectionName='" + modifiedObjectName + "']") != null)
                    {

                        //If the connection key exists, update the grouping of the key
                        //Otherwise, create a new record for the connection grouping
                        if (addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[@connectionName='" + modifiedObjectName + "']") != null)
                        {

                            //If the user chooses the blank option, then just delete the record in the xml part
                            //Otherwise, update the record to the newly selected grouping
                            if (dgrWbConnections[e.ColumnIndex, e.RowIndex].Value == null)
                            {
                                addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[@connectionName='" + modifiedObjectName + "']").Delete();
                            }
                            else
                            {
                                addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[@connectionName='" + modifiedObjectName + "']").Attributes[2].NodeValue = dgrWbConnections[e.ColumnIndex, e.RowIndex].Value.ToString();
                            }
                        }
                        else
                        {

                            //Creating the new connection grouping
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/ConnectionGroupings"), "ConnectionGrouping");

                            //Creating attributes
                            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++

                            //Connection Name
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[last()]"), "connectionName", "", NodeType: Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue: modifiedObjectName);

                            //Grouping
                            addInXmlPart.AddNode(addInXmlPart.SelectSingleNode("data/ConnectionGroupings/ConnectionGrouping[last()]"), "grouping", "", NodeType: Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue: dgrWbConnections[e.ColumnIndex, e.RowIndex].Value.ToString());
                            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        }
                    }
                    break;
            }
        }

        //User changes the grouping the want to filter on
        private void cboGroupingFilter_SelectedIndexChanged(object sender, EventArgs e)
        {

            //If the suer chooses the blank option, clear the filters
            if (String.IsNullOrEmpty(cboGroupingFilter.Text))
            {
                (dgrPivotTables.DataSource as DataTable).DefaultView.RowFilter = "";
                (dgrListObjects.DataSource as DataTable).DefaultView.RowFilter = "";
                (dgrWbConnections.DataSource as DataTable).DefaultView.RowFilter = "";
            }
            else
            {
                (dgrPivotTables.DataSource as DataTable).DefaultView.RowFilter = "pvtGrouping = '" + cboGroupingFilter.Text + "'";
                (dgrListObjects.DataSource as DataTable).DefaultView.RowFilter = "lstObjGrouping = '" + cboGroupingFilter.Text + "'";
                (dgrWbConnections.DataSource as DataTable).DefaultView.RowFilter = "dtaSrcGrouping = '" + cboGroupingFilter.Text + "'";
            }

            //Reset variable so the data source fields data grid will update
            previousDataSrcRowIndex = -1;

        }

        //User clicks to print the worksheet of any pivot thay has been checked to print in the pivots gird view
        private void btnPivotsQuickPrint_Click(object sender, EventArgs e)
        {

            //Variables
            DialogResult msgBoxResult;
            Excel.Application app;
            Excel.Workbook thisWorkbook;
            Excel.Worksheet ws;
            List<string> printedWs = new List<string>();
            bool printWs;

            //Confirming with the user they wish to print all selected pivots
            msgBoxResult = MessageBox.Show("Are you sure you want to print all the selected Pivots?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (msgBoxResult == DialogResult.Yes)
            {

                //Creating the activeworkbook object
                app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                app.Visible = true;
                thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;

                //Looping through each pivot record in the pivots grid view and print it if it was selected to
                foreach (DataGridViewRow row in dgrPivotTables.Rows)
                {
                    ws = thisWorkbook.Worksheets[row.Cells["PvtWorksheet"].Value.ToString()];
                    printWs = (row.Cells["PvtPrint"].Value.ToString() == "") ? false : Convert.ToBoolean(row.Cells["PvtPrint"].Value);
                    if (printWs && !printedWs.Contains(ws.Name))
                    {
                        try
                        {
                            ws.PrintOutEx();
                        }
                        catch
                        {
                            MessageBox.Show("The " + ws.Name + " worksheet was unable to be printed.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        printedWs.Add(ws.Name);
                    }
                }

                //Notifying the user once everything has been printed
                if (printedWs.Count > 0)
                {
                    MessageBox.Show("All successful PivotTable prints have been sent to the printer.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Nothing was printed.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
    }
}
