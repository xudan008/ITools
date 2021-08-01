using Microsoft.Office.Tools.Ribbon;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ITools
{
	public partial class Ribbon1
	{
		private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void checkBox1_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.ctp1.Visible = checkBox1.Checked;
		}
	}
}
