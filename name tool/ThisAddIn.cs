using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace name_tool
{
    public partial class ThisAddIn
    {
        public ShapeManagerForm ActiveShapeManager { get; set; }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            if (ActiveShapeManager != null && !ActiveShapeManager.IsDisposed)
            {
                ActiveShapeManager.SyncSelectionFromPowerPoint(Sel);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
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
        }
        
        #endregion
    }
}
