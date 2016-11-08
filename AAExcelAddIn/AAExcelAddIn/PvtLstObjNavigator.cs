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
            string dataSoruceName = "", dataSrouceType = "", dataSoruceDesc = "";

            //Creating the activeworkbook object
            app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            app.Visible = true;
            thisWorkbook = (Excel.Workbook)app.ActiveWorkbook;
            
            //Loading the data grids in the form
            foreach (Excel.Worksheet ws in thisWorkbook.Worksheets)
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

                    //Creating a new row in the data grid for each pivot
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(dgrPivotTables, pvt.Name, dataSoruceName, dataSrouceType, dataSoruceDesc);
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

                    //Creating a new row in the data grid for each list object
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(dgrListObjects, lst.Name, dataSoruceName, dataSrouceType, dataSoruceDesc);
                    dgrListObjects.Rows.Add(row);
                }
            }
        }
    }
}
