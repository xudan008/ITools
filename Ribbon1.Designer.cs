
namespace ITools
{
	partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// 必需的设计器变量。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public Ribbon1()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// 清理所有正在使用的资源。
		/// </summary>
		/// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region 组件设计器生成的代码

		/// <summary>
		/// 设计器支持所需的方法 - 不要修改
		/// 使用代码编辑器修改此方法的内容。
		/// </summary>
		private void InitializeComponent()
		{
			this.INS = this.Factory.CreateRibbonTab();
			this.Instrument = this.Factory.CreateRibbonGroup();
			this.checkBox1 = this.Factory.CreateRibbonCheckBox();
			this.INS.SuspendLayout();
			this.Instrument.SuspendLayout();
			this.SuspendLayout();
			// 
			// INS
			// 
			this.INS.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.INS.Groups.Add(this.Instrument);
			this.INS.Label = "TabAddIns";
			this.INS.Name = "INS";
			// 
			// Instrument
			// 
			this.Instrument.Items.Add(this.checkBox1);
			this.Instrument.Label = "INS";
			this.Instrument.Name = "Instrument";
			// 
			// checkBox1
			// 
			this.checkBox1.Label = "checkBox1";
			this.checkBox1.Name = "checkBox1";
			this.checkBox1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox1_Click);
			// 
			// Ribbon1
			// 
			this.Name = "Ribbon1";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.INS);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
			this.INS.ResumeLayout(false);
			this.INS.PerformLayout();
			this.Instrument.ResumeLayout(false);
			this.Instrument.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab INS;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup Instrument;
		internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
	}

	partial class ThisRibbonCollection
	{
		internal Ribbon1 Ribbon1
		{
			get { return this.GetRibbon<Ribbon1>(); }
		}
	}
}
