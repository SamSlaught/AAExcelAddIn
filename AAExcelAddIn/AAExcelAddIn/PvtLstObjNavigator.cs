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
using Excel = Microsoft.Office.Interop.Excel;

namespace AAExcelAddIn
{
    public partial class PvtLstObjNavigator : Form
    {
        public PvtLstObjNavigator()
        {
            InitializeComponent();
        }

        private void PvtLstObjNavigator_Load(object sender, EventArgs e)
        {

            //Variables
            Excel.Application app;
            Excel.Workbook thisWorkbook;
            string dataSoruceName = "", dataSrouceType = "", dataSoruceDesc = "", pageFields, rowFields, columnFields, dataFields, lstObjColumns;

            //Creating the activeworkbook object
            app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            app.Visible = true;
            thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;
            
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
                        row.CreateCells(dgrPivotTables, pvt.Name, ws.Name, dataSoruceName, dataSrouceType, dataSoruceDesc, pageFields, rowFields, columnFields, dataFields);
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
                                dataSoruceName = "";
                                dataSrouceType = "No External Source";
                                dataSoruceDesc = "";
                                break;

                            //Workbook Connection
                            case Excel.XlListObjectSourceType.xlSrcQuery:
                                dataSoruceName = lst.QueryTable.WorkbookConnection.Name;
                                dataSrouceType = "Workbook Connection";
                                dataSoruceDesc = lst.QueryTable.WorkbookConnection.Description;
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

                //Unselecting the first row in the data grids
                dgrPivotTables.ClearSelection();
                dgrListObjects.ClearSelection();
            }
        }

        //User double clicks to go to selected pivot in the workbook
        private void dgrPivotTables_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
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
    }
}
