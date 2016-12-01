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
        public Office.CustomXMLPart addInXmlPart;
        public bool creatingNewGrouping, newRecordRow;
        public string previousGrouping;

        private void PvtLstObjNavigator_Load(object sender, EventArgs e)
        {

            //Variables
            Excel.Application app;
            Excel.Workbook thisWorkbook;
            Office.DocumentProperty customXmlPartDocProp;
            string dataSoruceName = "", dataSrouceType = "", dataSoruceDesc = "", pageFields, rowFields, columnFields, dataFields, lstObjColumns, connType, connCommandText, connFilePath, connCommandType, connLastRefreshed;
            const string xmlPartTitle = "<title>AA Excel Add-In (Navigator)</title>", docPropertyName = "NavCustomXmlPartID";
            Excel.XlCmdType cmdType = Microsoft.Office.Interop.Excel.XlCmdType.xlCmdDefault;
            bool connPivotCache, connReadOnly, correctXmlPart = false;
            Nullable<decimal> connPvtChcSize = null;

            //Creating the activeworkbook object
            app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            app.Visible = true;
            thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;

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
            if (customXmlPartDocProp.Value == "0" || addInXmlPart == null || correctXmlPart == false)
            {

                //Grabbing the id of the xml part
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
                    addInXmlPart = thisWorkbook.CustomXMLParts.Add("<?xml version=\"1.0\" encoding=\"UTF - 8\"?><data>" + xmlPartTitle + "<Groupings></Groupings></data>");
                    customXmlPartDocProp.Value = addInXmlPart.Id;
                }
            }
            //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

            StringReader sr = new StringReader(addInXmlPart.SelectSingleNode("data/Groupings").XML);
            DataSet ds = new DataSet();
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

                        //Creating a new row in the data grid for each pivot
                        DataGridViewRow row = new DataGridViewRow();
                        row.CreateCells(dgrPivotTables, pvt.Name, ws.Name, "Go To", "", dataSoruceName, dataSrouceType, dataSoruceDesc, pvt.RefreshDate, pageFields, rowFields, columnFields, dataFields);
                        dgrPivotTables.Rows.Add(row);
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

                        //Creating a new row in the data grid for each list object
                        DataGridViewRow row = new DataGridViewRow();
                        row.CreateCells(dgrListObjects, lst.Name, ws.Name, dataSoruceName, dataSrouceType, dataSoruceDesc, lstObjColumns);
                        dgrListObjects.Rows.Add(row);
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

                //Creating a new row in the data grid for each list object
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dgrDataSources, conn.Name, conn.Description, connType, connPivotCache, connPvtChcSize, connLastRefreshed, connReadOnly, connCommandText, connFilePath, connCommandType);
                dgrDataSources.Rows.Add(row);
            }

            //Right aligning certain columns in the data source grid view
            dgrDataSources.Columns["dtaSrcPvtChcMemory"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            //Unselecting the first row in the data grids
            dgrPivotTables.ClearSelection();
            dgrListObjects.ClearSelection();
            dgrDataSources.ClearSelection();
        }

        //User double clicks to go to selected pivot in the workbook
        private void dgrPivotTables_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
        }

        //User double clicks to go to selected list object in the workbook
        private void dgrListObjects_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            //Variables
            Excel.Application app;
            Excel.Workbook thisWorkbook;

            //Making sure the double clicked row isn't the header
            if (e.RowIndex != -1)
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

        private void dgrDataSources_SelectionChanged(object sender, EventArgs e)
        {

            //Variables
            Excel.Application app;
            Excel.Workbook thisWorkbook;
            bool cacheFieldsFound = false;
            string pvtFieldDataType;

            //Creating the activeworkbook object
            app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            app.Visible = true;
            thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;

            //Resetting the fields sub grid if needed
            if (dgrPvtChcFields.RowCount > 0)
            {
                dgrPvtChcFields.Rows.Clear();
                dgrPvtChcFields.Refresh();
            }

            //Only check for piovt cache fields if the data source is linked to a pivot cache
            if (Convert.ToBoolean(dgrDataSources.Rows[dgrDataSources.CurrentCell.RowIndex].Cells[3].Value) == true)
            {

                //Looping through each pivot cache to see if the selected connection is linked to any of them
                foreach (Excel.PivotCache pvtCache in thisWorkbook.PivotCaches())
                {

                    //Handle scenario if a pivot cache is not linked to a workbook connection
                    try
                    {

                        //If the workbook connection's name matches the name of the connection linked to the current pivot cache, grab its fields
                        if (pvtCache.WorkbookConnection.Name == dgrDataSources.Rows[dgrDataSources.CurrentCell.RowIndex].Cells[0].Value.ToString())
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
                }
                else
                {
                    addInXmlPart.SelectSingleNode("data/Groupings/Grouping[text()=\"" + previousGrouping + "\"]").Text = dgrGroupings[e.ColumnIndex, e.RowIndex].Value.ToString();
                }
            }
        }

        //If the user deletes the grouping from the grid, delete it in the xml part
        private void dgrGroupings_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            newRecordRow = (dgrGroupings.NewRowIndex == e.Row.Index);
            if (!newRecordRow)
            {
                addInXmlPart.SelectSingleNode("data/Groupings/Grouping[text()=\"" + e.Row.Cells[0].Value.ToString() + "\"]").Delete();
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

            if (e.ColumnIndex == 2)
            {
                
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
            else if (e.FormattedValue.ToString().Contains("\""))
            {
                MessageBox.Show("Double quotes are illegal characters.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
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
    }
}
