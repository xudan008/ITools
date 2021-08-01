using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;


namespace ITools
{
	public partial class UserControl1 : UserControl
	{
		public UserControl1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			//Globals.ThisAddIn.Application.ActiveCell.Value = "Hellow World";
			//Excel.Workbook newWorkbook = Globals.ThisAddIn.Application.Workbooks.Add(System.Type.Missing);
			//外接程序中创建新的工作簿

			string item = Globals.ThisAddIn.Application.ActiveCell.Value + "," + Globals.ThisAddIn.Application.ActiveCell.Address + "\r";
			this.richTextBox1.Text += item;
		}

		private void button3_Click(object sender, EventArgs e)
		{
			Excel.Worksheet newWorksheet;
			newWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
			newWorksheet.Name = "Data-1";
			newWorksheet.Tab.Color = Color.BlueViolet;
			//ListSheets();
		}

		private void button2_Click(object sender, EventArgs e)
		{
			string item = Globals.ThisAddIn.Application.ActiveCell.Value + "," + Globals.ThisAddIn.Application.ActiveCell.Address + "\r";
			this.listBox1.Items.Add(item);
		}
	}
}
