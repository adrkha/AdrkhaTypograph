using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace AdrkhaTypograph
{
    [ComVisible(true)]
    public class TypoRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""Ribbon_Load"">
  <ribbon>
    <tabs>
      <tab id=""TabAdrkhaTypo"" label=""أدركها تايبوجراف"">
        <group id=""GroupTypo"" label=""تحكم اللوحة"">
          <button id=""BtnToggleTaskPane"" 
                  label=""إظهار / إخفاء اللوحة"" 
                  size=""large"" 
                  imageMso=""FontProperties"" 
                  onAction=""OnToggleTaskPaneClicked"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnToggleTaskPaneClicked(Office.IRibbonControl control)
        {
            Globals.ThisAddIn?.ToggleTaskPane();
        }
    }
}
