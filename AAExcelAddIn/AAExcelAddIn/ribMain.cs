using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace AAExcelAddIn
{
    public partial class ribMain
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        //User clicks to open the navigator form
        private void btnNavigator_Click(object sender, RibbonControlEventArgs e)
        {
            PvtLstObjNavigator navigator = new PvtLstObjNavigator();
            navigator.ShowDialog();
        }
    }
}
