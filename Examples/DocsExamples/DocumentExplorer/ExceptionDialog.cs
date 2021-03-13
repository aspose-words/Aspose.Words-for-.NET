using System;
using System.Windows.Forms;

namespace DocumentExplorer
{

	/// <summary>
	/// Provides full information about application exception.
	/// </summary>
	public class ExceptionDialog : Form 
	{
		#region Windows Form Designer generated code

		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button buttonOk;
		private System.Windows.Forms.TextBox text1;

		private void InitializeComponent() {
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(ExceptionDialog));
			this.text1 = new System.Windows.Forms.TextBox();
			this.buttonOk = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// Text1
			// 
			this.text1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.text1.AutoSize = false;
			this.text1.Location = new System.Drawing.Point(8, 8);
			this.text1.Multiline = true;
			this.text1.Name = "text1";
			this.text1.ReadOnly = true;
			this.text1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.text1.Size = new System.Drawing.Size(524, 244);
			this.text1.TabIndex = 0;
			this.text1.Text = "";
			this.text1.WordWrap = false;
			// 
			// ButtonOk
			// 
			this.buttonOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.buttonOk.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.buttonOk.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.buttonOk.Location = new System.Drawing.Point(432, 260);
			this.buttonOk.Name = "buttonOk";
			this.buttonOk.Size = new System.Drawing.Size(100, 24);
			this.buttonOk.TabIndex = 12;
			this.buttonOk.Text = "Continue";
			// 
			// ExceptionDialog
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(540, 288);
			this.Controls.Add(this.buttonOk);
			this.Controls.Add(this.text1);
			this.DockPadding.All = 8;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ExceptionDialog";
			this.ShowInTaskbar = false;
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Unexpected error occured";
			this.ResumeLayout(false);

		}

		protected override void Dispose( bool disposing ) 
        {
			if( disposing ) 
            {
				if (components != null) 
                {
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#endregion

		public ExceptionDialog() 
		{
			InitializeComponent();
		}

		public ExceptionDialog(Exception ex) 
        {
			InitializeComponent();
            Text = "Document Explorer - unexpected error occured";
			text1.Text = "\r\n" + Application.ProductName + ".exe \r\n\r\n" + 
				"Version " + Application.ProductVersion + "\r\n\r\n" + 
				DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString() + "\r\n\r\n" +
				ex.ToString() + "\r\n";
			text1.SelectionStart = text1.Text.Length;
		}
	}
}