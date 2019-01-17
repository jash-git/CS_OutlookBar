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
        private IContainer components;
        private ImageList imageList1;
        private OutlookBar outlookBar1;
		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call				 
			//

		}

		public void PanelEvent(object sender, EventArgs e)
		{
			Control ctrl=(Control)sender;
			PanelIcon panelIcon=ctrl.Tag as PanelIcon;
            MessageBox.Show("#" + outlookBar1.SelectedBand+ "," + panelIcon.Index.ToString(), "Panel Event");
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.outlookBar1 = new OutlookBar();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.SuspendLayout();
            // 
            // outlookBar1
            // 
            this.outlookBar1.AutoScroll = true;
            this.outlookBar1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.outlookBar1.ButtonHeight = 25;
            this.outlookBar1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.outlookBar1.Location = new System.Drawing.Point(0, 0);
            this.outlookBar1.Name = "outlookBar1";
            this.outlookBar1.SelectedBand = 0;
            this.outlookBar1.Size = new System.Drawing.Size(304, 259);
            this.outlookBar1.TabIndex = 0;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "home.png");
            this.imageList1.Images.SetKeyName(1, "envelope.png");
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(8, 23);
            this.ClientSize = new System.Drawing.Size(304, 259);
            this.Controls.Add(this.outlookBar1);
            this.Font = new System.Drawing.Font("ÐÂ¼šÃ÷ów", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Resize += new System.EventHandler(this.Form1_Resize);
            this.ResumeLayout(false);

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
        public void CreateMenu()
        {
            IconPanel iconPanel1 = new IconPanel();
            IconPanel iconPanel2 = new IconPanel();
            IconPanel iconPanel3 = new IconPanel();
            outlookBar1.DelAllBand();
            outlookBar1.AddBand(imageList1.Images[0], "Outlook Shortcuts", iconPanel1);
            outlookBar1.AddBand("My Shortcuts", iconPanel2);
            outlookBar1.AddBand("Other Shortcuts", iconPanel3);

            iconPanel1.AddIcon("Outlook Today", imageList1.Images[1], new EventHandler(PanelEvent), outlookBar1.ButtonHeight);
            iconPanel1.AddIcon("Calendar", imageList1.Images[1], new EventHandler(PanelEvent), outlookBar1.ButtonHeight);
            iconPanel1.AddIcon("Contacts", imageList1.Images[1], new EventHandler(PanelEvent), outlookBar1.ButtonHeight);
            iconPanel1.AddIcon("Tasks", imageList1.Images[1], new EventHandler(PanelEvent), outlookBar1.ButtonHeight);

            iconPanel2.AddIcon("Calendar", imageList1.Images[1], new EventHandler(PanelEvent), outlookBar1.ButtonHeight);
            iconPanel2.AddIcon("Outlook Today", imageList1.Images[1], new EventHandler(PanelEvent), outlookBar1.ButtonHeight);
            iconPanel2.AddIcon("Contacts", imageList1.Images[1], new EventHandler(PanelEvent), outlookBar1.ButtonHeight);
            iconPanel2.AddIcon("Tasks", imageList1.Images[1], new EventHandler(PanelEvent),outlookBar1.ButtonHeight);

            iconPanel3.AddIcon("Calendar", imageList1.Images[1], new EventHandler(PanelEvent), outlookBar1.ButtonHeight);
            iconPanel3.AddIcon("Outlook Today", imageList1.Images[1], new EventHandler(PanelEvent), outlookBar1.ButtonHeight);
            iconPanel3.AddIcon("Tasks", imageList1.Images[1], new EventHandler(PanelEvent), outlookBar1.ButtonHeight);
            iconPanel3.AddIcon("Contacts", imageList1.Images[1], new EventHandler(PanelEvent), outlookBar1.ButtonHeight);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            outlookBar1.Initialize();
            CreateMenu();
            outlookBar1.SelectBand(2);
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            CreateMenu();
        }
	}
}
