using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace AdrkhaTypograph // <--- غير هذا الاسم إذا كان اسم مشروعك مختلفاً
{
    public partial class ThisAddIn
    {
        private TypoTaskPane myTaskPaneControl;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            myTaskPaneControl = new TypoTaskPane();
            myCustomTaskPane = this.CustomTaskPanes.Add(myTaskPaneControl, "أداة التايبوجرافي العربي");
            myCustomTaskPane.Width = 300;
            myCustomTaskPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}