using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Drawing;
using System.Windows.Forms;

namespace name_tool
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("name_tool.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnOpenManagerClick(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn.ActiveShapeManager == null || Globals.ThisAddIn.ActiveShapeManager.IsDisposed)
                {
                    Globals.ThisAddIn.ActiveShapeManager = new ShapeManagerForm(Globals.ThisAddIn.Application);
                }
                Globals.ThisAddIn.ActiveShapeManager.Show();
                Globals.ThisAddIn.ActiveShapeManager.BringToFront();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening Shape Manager: " + ex.Message);
            }
        }

        public void OnQuickSelectClick(Office.IRibbonControl control)
        {
            try
            {
                QuickSelectHelper.Execute(Globals.ThisAddIn.Application);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in Quick Select: " + ex.Message, "Quick Select",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public System.Drawing.Image GetImage(Office.IRibbonControl control)
        {
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                string[] resourceNames = asm.GetManifestResourceNames();
                
                // Search for the image resource. 
                // We'll name it 'pic.png' when we embed it.
                foreach (string resourceName in resourceNames)
                {
                    if (resourceName.EndsWith("pic.png", StringComparison.OrdinalIgnoreCase))
                    {
                        using (Stream stream = asm.GetManifestResourceStream(resourceName))
                        {
                            if (stream != null)
                            {
                                return System.Drawing.Image.FromStream(stream);
                            }
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
