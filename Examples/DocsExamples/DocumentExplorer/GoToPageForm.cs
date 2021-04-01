using System;
using System.Windows.Forms;

namespace DocumentExplorer
{
	/// <summary>
	/// Lets the user specify the page number to go to.
	/// </summary>
	public class GoToPageForm : System.Windows.Forms.Form
	{
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Label promptLabel;
		private System.Windows.Forms.TextBox pageNumberTextBox;
		private int mPageNumber;
		private System.Windows.Forms.Button okBtn;
		private System.Windows.Forms.Button cancelBtn;
		private int mMaxPageNumber = 1;

		public GoToPageForm()
		{
			InitializeComponent();
		}

		public int MaxPageNumber
		{
			set { mMaxPageNumber = value; }
		}

		public int PageNumber
		{
			get { return mPageNumber; }
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// The contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.promptLabel = new System.Windows.Forms.Label();
			this.pageNumberTextBox = new System.Windows.Forms.TextBox();
			this.okBtn = new System.Windows.Forms.Button();
			this.cancelBtn = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// PromptLabel
			// 
			this.promptLabel.Location = new System.Drawing.Point(16, 16);
			this.promptLabel.Name = "promptLabel";
			this.promptLabel.Size = new System.Drawing.Size(192, 23);
			this.promptLabel.TabIndex = 0;
			// 
			// PageNumberTextBox
			// 
			this.pageNumberTextBox.Location = new System.Drawing.Point(16, 40);
			this.pageNumberTextBox.MaxLength = 5;
			this.pageNumberTextBox.Name = "pageNumberTextBox";
			this.pageNumberTextBox.Size = new System.Drawing.Size(192, 20);
			this.pageNumberTextBox.TabIndex = 1;
			this.pageNumberTextBox.Text = "";
			this.pageNumberTextBox.TextChanged += new System.EventHandler(this.pageNumberTextBox_TextChanged);
			// 
			// OkBtn
			// 
			this.okBtn.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.okBtn.Enabled = false;
			this.okBtn.Location = new System.Drawing.Point(56, 72);
			this.okBtn.Name = "okBtn";
			this.okBtn.TabIndex = 2;
			this.okBtn.Text = "OK";
			this.okBtn.Click += new System.EventHandler(this.okButton_Click);
			// 
			// CancelBtn
			// 
			this.cancelBtn.CausesValidation = false;
			this.cancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.cancelBtn.Location = new System.Drawing.Point(136, 72);
			this.cancelBtn.Name = "cancelBtn";
			this.cancelBtn.TabIndex = 3;
			this.cancelBtn.Text = "Cancel";
			// 
			// GoToPageForm
			// 
			this.AcceptButton = this.okBtn;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.cancelBtn;
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(218, 104);
			this.Controls.Add(this.okBtn);
			this.Controls.Add(this.pageNumberTextBox);
			this.Controls.Add(this.promptLabel);
			this.Controls.Add(this.cancelBtn);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "GoToPageForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Go to Page";
			this.Load += new System.EventHandler(this.GoToPageForm_Load);
			this.ResumeLayout(false);

		}
		#endregion

		private void GoToPageForm_Load(object sender, System.EventArgs e)
		{
			promptLabel.Text = String.Format("Enter page number (1-{0})", mMaxPageNumber);
		}

		private void pageNumberTextBox_TextChanged(object sender, System.EventArgs e)
		{
			okBtn.Enabled = (pageNumberTextBox.Text.Length > 0);
		}

		private void okButton_Click(object sender, System.EventArgs e)
		{
			try
			{
				if (!TryParse(pageNumberTextBox.Text, out mPageNumber))
					throw new Exception("Please enter a valid page number.");

				if ((mPageNumber < 1) || (mPageNumber > mMaxPageNumber))
					throw new Exception(string.Format("Page number must be between 1 and {0}.", mMaxPageNumber));
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				DialogResult = DialogResult.None;
			}
		}

		private static bool TryParse(string text, out int value)
		{
			value = 0;
			try
			{
				value = int.Parse(text);
				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}
	}
}