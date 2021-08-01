using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace ITools
{
    public partial class ThisAddIn
    {
        UserControl1 uc = new UserControl1();
        public Microsoft.Office.Tools.CustomTaskPane ctp1;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {


            ctp1 = Globals.ThisAddIn.CustomTaskPanes.Add(uc, "仪表表格辅助工具");
            ctp1.Visible = true;
        }



        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}