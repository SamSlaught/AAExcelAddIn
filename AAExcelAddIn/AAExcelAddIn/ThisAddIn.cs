using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace AAExcelAddIn
{
    public partial class ThisAddIn
    {

        //Globals
        public ribMain ribbon;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            //Storing the ribbon object globally
            Type type = typeof(ribMain);
            ribbon = Globals.Ribbons.GetRibbon(type) as ribMain;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        //Excel Events
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        private void Excel_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            if (Wb.Application.Workbooks.Count == 1)
            {
                ToggleRibbonLock(true);
            }
        }

        private void Excel_NewWorkbook(Excel.Workbook Wb)
        {
            if (Wb.Application.Workbooks.Count == 1)
            {
                ToggleRibbonLock(false);
            }
        }

        private void Excel_WorkbookOpen(Excel.Workbook Wb)
        {
            if (Wb.Application.Workbooks.Count == 1)
            {
                ToggleRibbonLock(false);
            }
        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        //Toggles certain controls to be enable/disabled for certain conditions with the application
        void ToggleRibbonLock(bool lockRibbon)
        {
            ribbon.btnNavigator.Enabled = !lockRibbon;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);

            //Excel Events
            ((Excel.AppEvents_Event)this.Application).NewWorkbook += new Microsoft.Office.Interop.Excel.AppEvents_NewWorkbookEventHandler(Excel_NewWorkbook);
            this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Excel_WorkbookOpen);
            this.Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(Excel_WorkbookBeforeClose);
        }
        
        #endregion
    }
}
