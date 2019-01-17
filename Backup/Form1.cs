using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace OutlookBar
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call				 
			//

			OutlookBar outlookBar=new OutlookBar();
			outlookBar.Location=new Point(0, 0);
			outlookBar.Size=new Size(150, this.ClientSize.Height);
			outlookBar.BorderStyle=BorderStyle.FixedSingle;
			Controls.Add(outlookBar);
			outlookBar.Initialize();

			IconPanel iconPanel1=new IconPanel();
			IconPanel iconPanel2=new IconPanel();
			IconPanel iconPanel3=new IconPanel();
			outlookBar.AddBand("Outlook Shortcuts", iconPanel1);
			outlookBar.AddBand("My Shortcuts", iconPanel2);
			outlookBar.AddBand("Other Shortcuts", iconPanel3);

			iconPanel1.AddIcon("Outlook Today", Image.FromFile("img1.ico"), new EventHandler(PanelEvent));
			iconPanel1.AddIcon("Calendar", Image.FromFile("img2.ico"), new EventHandler(PanelEvent));
			iconPanel1.AddIcon("Contacts", Image.FromFile("img3.ico"), new EventHandler(PanelEvent));
			iconPanel1.AddIcon("Tasks", Image.FromFile("img4.ico"), new EventHandler(PanelEvent));
			outlookBar.SelectBand(0);
		}

		public void PanelEvent(object sender, EventArgs e)
		{
			Control ctrl=(Control)sender;
			PanelIcon panelIcon=ctrl.Tag as PanelIcon;
			MessageBox.Show("#"+panelIcon.Index.ToString(), "Panel Event");
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(304, 259);
			this.Name = "Form1";
			this.Text = "Form1";

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}
	}
}
