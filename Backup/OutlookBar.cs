using System;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookBar
{
	internal class BandTagInfo
	{
		public OutlookBar outlookBar;
		public int index;

		public BandTagInfo(OutlookBar ob, int index)
		{
			outlookBar=ob;
			this.index=index;
		}
	}

	public class OutlookBar : Panel
	{
		private int buttonHeight;
		private int selectedBand;
		private int selectedBandHeight;

		public int ButtonHeight
		{
			get
			{
				return buttonHeight;
			}

			set
			{
				buttonHeight=value;
				// do recalc layout for entire bar
			}
		}

		public int SelectedBand
		{
			get
			{
				return selectedBand;
			}
			set
			{
				SelectBand(value);
			}
		}

		public OutlookBar()
		{
			buttonHeight=25;
			selectedBand=0;
			selectedBandHeight=0;
		}

		public void Initialize()
		{
			// parent must exist!
			Parent.SizeChanged+=new EventHandler(SizeChangedEvent);
		}

		public void AddBand(string caption, ContentPanel content)
		{
			content.outlookBar=this;
			int index=Controls.Count;
			BandTagInfo bti=new BandTagInfo(this, index);
			BandPanel bandPanel=new BandPanel(caption, content, bti);
			Controls.Add(bandPanel);
			UpdateBarInfo();
			RecalcLayout(bandPanel, index);
		}

		public void SelectBand(int index)
		{
			selectedBand=index;
			RedrawBands();
		}

		private void RedrawBands()
		{
			for (int i=0; i<Controls.Count; i++)
			{
				BandPanel bp=Controls[i] as BandPanel;
				RecalcLayout(bp, i);
			}
		}

		private void UpdateBarInfo()
		{
			selectedBandHeight=ClientRectangle.Height-(Controls.Count * buttonHeight);
		}

		private void RecalcLayout(BandPanel bandPanel, int index)
		{
			int vPos=(index <= selectedBand) ? buttonHeight*index : buttonHeight*index+selectedBandHeight;
			int height=selectedBand==index ? selectedBandHeight+buttonHeight : buttonHeight;

			// the band dimensions
			bandPanel.Location=new Point(0, vPos);
			bandPanel.Size=new Size(ClientRectangle.Width, height);

			// the contained button dimensions
			bandPanel.Controls[0].Location=new Point(0, 0);
			bandPanel.Controls[0].Size=new Size(ClientRectangle.Width, buttonHeight);

			// the contained content panel dimensions
			bandPanel.Controls[1].Location=new Point(0, buttonHeight);
			bandPanel.Controls[1].Size=new Size(ClientRectangle.Width-2, height-8);
		}

		private void SizeChangedEvent(object sender, EventArgs e)
		{
			Size=new Size(Size.Width, ((Control)sender).ClientRectangle.Size.Height);
			UpdateBarInfo();
			RedrawBands();
		}
	}

	internal class BandPanel : Panel
	{
		public BandPanel(string caption, ContentPanel content, BandTagInfo bti)
		{
			BandButton bandButton=new BandButton(caption, bti);
			Controls.Add(bandButton);
			Controls.Add(content);
		}
	}

	internal class BandButton : Button
	{
		private BandTagInfo bti;

		public BandButton(string caption, BandTagInfo bti)
		{
			Text=caption;
			FlatStyle=FlatStyle.Standard;
			Visible=true;
			this.bti=bti;
			Click+=new EventHandler(SelectBand);
		}

		private void SelectBand(object sender, EventArgs e)
		{
			bti.outlookBar.SelectBand(bti.index);
		}
	}

	public abstract class ContentPanel : Panel
	{
		public OutlookBar outlookBar;

		public ContentPanel()
		{
			// initial state
			Visible=true;
		}
	}

	public class IconPanel : ContentPanel
	{
		protected int iconSpacing;
		protected int margin;

		public int IconSpacing
		{
			get
			{
				return iconSpacing;
			}
		}

		public int Margin
		{
			get
			{
				return margin;
			}
		}

		public IconPanel()
		{
			margin=10;
			iconSpacing=32+15+10;	// icon height + text height + margin
			BackColor=Color.LightBlue;
			AutoScroll=true;
		}

		public void AddIcon(string caption, Image image, EventHandler onClickEvent)
		{
			int index=Controls.Count/2;	// two entries per icon
			PanelIcon panelIcon=new PanelIcon(this, image, index, onClickEvent);
			Controls.Add(panelIcon);

			Label label=new Label();
			label.Text=caption;
			label.Visible=true;
			label.Location=new Point(0, margin+image.Size.Height+index*iconSpacing);
			label.Size=new Size(Size.Width, 15);
			label.TextAlign=ContentAlignment.TopCenter;
			label.Click+=onClickEvent;
			label.Tag=panelIcon;
			Controls.Add(label);
		}
	}

	public class PanelIcon : PictureBox
	{
		public int index;
		public IconPanel iconPanel;

		private Color bckgColor;
		private bool mouseEnter;

		public int Index
		{
			get
			{
				return index;
			}
		}

		public PanelIcon(IconPanel parent, Image image, int index, EventHandler onClickEvent)
		{
			this.index=index;
			this.iconPanel=parent;
			Image=image;
			Visible=true;
			Location=new Point(iconPanel.outlookBar.Size.Width/2-image.Size.Width/2,
							iconPanel.Margin + index*iconPanel.IconSpacing);
			Size=image.Size;
			Click+=onClickEvent;
			Tag=this;

			MouseEnter+=new EventHandler(OnMouseEnter);
			MouseLeave+=new EventHandler(OnMouseLeave);
			MouseMove+=new MouseEventHandler(OnMouseMove);

			bckgColor=iconPanel.BackColor;
			mouseEnter=false;
		}

		private void OnMouseMove(object sender, MouseEventArgs args)
		{
			if ( (args.X < Size.Width-2) &&
				(args.Y < Size.Width-2) &&
				(!mouseEnter) )
			{
				BackColor=Color.LightCyan;
				BorderStyle=BorderStyle.FixedSingle;
				Location=Location-new Size(1, 1);
				mouseEnter=true;
			}
		}

		private void OnMouseEnter(object sender, EventArgs e)
		{
		}

		private void OnMouseLeave(object sender, EventArgs e)
		{
			if (mouseEnter)
			{
				BackColor=bckgColor;
				BorderStyle=BorderStyle.None;
				Location=Location+new Size(1, 1);
				mouseEnter=false;
			}
		}
	}
}
