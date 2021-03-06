/*
 * Created by SharpDevelop.
 * User: TheRedLord
 * Date: 4/4/2017
 * Time: 10:46
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace ExportExiel
{
	/// <summary>
	/// Description of gridsecound.
	/// </summary>
	public partial class gridsecound : Form
	{
		public gridsecound()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
	}
	public class TextBoxAndCheckBoxHeaderCella : DataGridViewColumnHeaderCell
    {
        enum SortType
        {
            ASC,
            DesC,
            None
        }
 		public List<String> list = new List<String>();
        private  TextBox tb = new TextBox();
        private Label cb= new Label();
        public  event EventHandler TextBoxValueChanged;//event fired when TextBoxValue property is changed.
        public event EventHandler CheckBoxValueChanged;//event fired when CheckBoxValue property is changed.
	
        SortType sType;

        
        public string TextBoxValue
        {
            get
            {
                return tb.Text;
            }
            set
            {
                if (value != tb.Text)
                {
                    tb.Text = value;
                    if (TextBoxValueChanged != null)
                    {
                        TextBoxValueChanged(this, EventArgs.Empty);
                    }
                }
            }
        }
        protected override void OnClick(DataGridViewCellEventArgs e)
        {
            base.OnClick(e);
            if(sType == SortType.None)
            {
                sType = SortType.ASC;
            }
            else if (sType == SortType.ASC)
            {
                sType = SortType.DesC;
            }
            else
            {
                sType = SortType.ASC;
            }
            ListaDetali.ColoumnIndex=this.ColumnIndex;
            
        }
        public TextBoxAndCheckBoxHeaderCella()
        {
        	String id=MainForm.ID;
            tb.MouseClick += new MouseEventHandler(tb_MouseClick);
            this.cb.MouseClick += new MouseEventHandler(cb_MouseClick);
            //this.tb.TextChanged += new EventHandler(tb_TextChanged);
            this.tb.Leave+= new EventHandler(tb_Leave);
            this.tb.KeyDown+= new KeyEventHandler(tb_KeyDown);
            sType = SortType.None;
            this.cb.Text = id;
        }

        void cb_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBoxValueChanged != null)
            {
                CheckBoxValueChanged(this, e);
            }
        }
        void tb_KeyDown(object sender,KeyEventArgs e)
        {
        	if (e.KeyCode == Keys.Return)
      		 {
        		 int ColoumIndex=this.ColumnIndex;
            String myString="-"+ColoumIndex.ToString()+"-";
            String Data=this.TextBoxValue;
            
            //hide
            //check if elements exists;
            //new
            if(Data=="")
            {
            	//cand e gol, facem golire la lista,activam functia de sort
            	//functia de unhide all
            	//MainForm.fo.oneMoreUnhide();
            	//Remove item
            	var matcha = ListaDetali.list.FirstOrDefault(stringToCheck => stringToCheck.Contains(myString));
				if(matcha != null)
				{
					ListaDetali.list.Remove(matcha);
				}
				//functia de resortare
				//MainForm.fo.MasterSort();
				//QueryAgain
				ListaDetali.fo.QueryAgain();
            }else
            {
            	//Cand nu e gold
            	//verificam daca avem item in lista
            	var match = ListaDetali.list.FirstOrDefault(stringToCheck => stringToCheck.Contains(myString));
            	if(match != null)
            	{
            		//daca itemul este in lista
            		//functia de unhide
            		//MainForm.fo.oneMoreUnhide();
            		//modifica itemul din lista
            		for (int i = 0; i < ListaDetali.list.Count; i++)
					{
						if (ListaDetali.list[i].Contains("-"+ColoumIndex.ToString()+"-"))
    					{
        					ListaDetali.list[i] = "-"+ColoumIndex.ToString()+"-"+Data;
        					break;
    					}
					}
            		//functia de resortare
            		//MainForm.fo.MasterSort();
            		//QueryAgain
            		ListaDetali.fo.QueryAgain();
            		
            	}else
            	{
            		//daca e itemul nu este in lista
            		//functia de unhide
            		//MainForm.fo.oneMoreUnhide();
            		//adauga in lista
            		String dataToList="-"+ColoumIndex.ToString()+"-"+Data;
					ListaDetali.list.Add(dataToList);
            		//functia de resortare
            		//MainForm.fo.MasterSort();
            		//queryAgain
            		ListaDetali.fo.QueryAgain();
            	}
            }
        	}
        }
        void tb_Leave(object sender,EventArgs e)
        {
        	
        	int ColoumIndex=this.ColumnIndex;
            String myString="-"+ColoumIndex.ToString()+"-";
            String Data=this.TextBoxValue;
            
            //hide
            //check if elements exists;
            //new
            if(Data=="")
            {
            	//cand e gol, facem golire la lista,activam functia de sort
            	//functia de unhide all
            	//MainForm.fo.oneMoreUnhide();
            	//Remove item
            	var matcha = ListaDetali.list.FirstOrDefault(stringToCheck => stringToCheck.Contains(myString));
				if(matcha != null)
				{
					ListaDetali.list.Remove(matcha);
				}
				//functia de resortare
				//MainForm.fo.MasterSort();
				//QueryAgain
				ListaDetali.fo.QueryAgain();
            }else
            {
            	//Cand nu e gold
            	//verificam daca avem item in lista
            	string match;
            	try
            	{
            	match = ListaDetali.list.FirstOrDefault(stringToCheck => stringToCheck.Contains(myString));
            	}catch{
            	match=null;
            	}
            	if(match != null)
            	{
            		//daca itemul este in lista
            		//functia de unhide
            		//MainForm.fo.oneMoreUnhide();
            		//modifica itemul din lista
            		for (int i = 0; i < ListaDetali.list.Count; i++)
					{
						if (ListaDetali.list[i].Contains("-"+ColoumIndex.ToString()+"-"))
    					{
        					ListaDetali.list[i] = "-"+ColoumIndex.ToString()+"-"+Data;
        					break;
    					}
					}
            		//functia de resortare
            		//MainForm.fo.MasterSort();
            		//QueryAgain
            		ListaDetali.fo.QueryAgain();
            		
            	}else
            	{
            		//daca e itemul nu este in lista
            		//functia de unhide
            		//MainForm.fo.oneMoreUnhide();
            		//adauga in lista
            		String dataToList="-"+ColoumIndex.ToString()+"-"+Data;
					ListaDetali.list.Add(dataToList);
            		//functia de resortare
            		//MainForm.fo.MasterSort();
            		//queryAgain
            		ListaDetali.fo.QueryAgain();
            	}
            }
        }
        void tb_TextChanged(object sender, EventArgs e)
        {
        	//hide if text not likewise
            if (TextBoxValueChanged != null)
            {
                TextBoxValueChanged(this, e);
            }
            //get coloum number
            int ColoumIndex=this.ColumnIndex;
            String myString="-"+ColoumIndex.ToString()+"-";
            String Data=this.TextBoxValue;
            
            //hide
            //check if elements exists;
            //new
            if(Data=="")
            {
            	//cand e gol, facem golire la lista,activam functia de sort
            	//functia de unhide all
            	//MainForm.fo.oneMoreUnhide();
            	//Remove item
            	var matcha = MainForm.fo.list.FirstOrDefault(stringToCheck => stringToCheck.Contains(myString));
				if(matcha != null)
				{
					ListaDetali.list.Remove(matcha);
				}
				//functia de resortare
				//MainForm.fo.MasterSort();
				//QueryAgain
				//MainForm.fo.QueryAgain();
            }else
            {
            	//Cand nu e gold
            	//verificam daca avem item in lista
            	var match = ListaDetali.list.FirstOrDefault(stringToCheck => stringToCheck.Contains(myString));
            	if(match != null)
            	{
            		//daca itemul este in lista
            		//functia de unhide
            		//MainForm.fo.oneMoreUnhide();
            		//modifica itemul din lista
            		for (int i = 0; i < MainForm.fo.list.Count; i++)
					{
						if (ListaDetali.list[i].Contains("-"+ColoumIndex.ToString()+"-"))
    					{
        					ListaDetali.list[i] = "-"+ColoumIndex.ToString()+"-"+Data;
        					break;
    					}
					}
            		//functia de resortare
            		//MainForm.fo.MasterSort();
            		//QueryAgain
            		//MainForm.fo.QueryAgain();
            		
            	}else
            	{
            		//daca e itemul nu este in lista
            		//functia de unhide
            		//MainForm.fo.oneMoreUnhide();
            		//adauga in lista
            		String dataToList="-"+ColoumIndex.ToString()+"-"+Data;
					ListaDetali.list.Add(dataToList);
            		//functia de resortare
            		//MainForm.fo.MasterSort();
            		//queryAgain
            		//MainForm.fo.QueryAgain();
            	}
            }
           
        }
        void cb_MouseClick(object sender, MouseEventArgs e)
        {
            this.OnClick(new DataGridViewCellEventArgs(this.ColumnIndex, this.RowIndex));
            ListaDetali.ColoumnIndex=ColumnIndex;
        }

        void tb_MouseClick(object sender, MouseEventArgs e)
        {
            this.OnClick(new DataGridViewCellEventArgs(this.ColumnIndex, this.RowIndex));
        }
        
        protected override void Paint(Graphics graphics, Rectangle clipBounds, Rectangle cellBounds, int rowIndex, DataGridViewElementStates dataGridViewElementState, object value, object formattedValue, string errorText, DataGridViewCellStyle cellStyle, DataGridViewAdvancedBorderStyle advancedBorderStyle, DataGridViewPaintParts paintParts)
        {
            base.Paint(graphics, clipBounds, cellBounds, rowIndex, dataGridViewElementState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts);
			
            //Here Change the location, height, width property of the TextBox and CheckBox..
            this.tb.Location = new Point(cellBounds.Left, cellBounds.Top);
            this.tb.Height = (cellBounds.Height-2) / 2+1;
            this.tb.Width = cellBounds.Width-2;
            this.cb.Location = new Point(cellBounds.Left, cellBounds.Height / 2);
            this.cb.Height = cellBounds.Height / 2;
            this.cb.Width = cellBounds.Width*3/5;
			
            //Painting code for sort icon
			//only paint 4
            if(this.ColumnIndex==ListaDetali.ColoumnIndex){
            GraphicsPath path = new GraphicsPath();

            if (this.sType == SortType.ASC)
            {
            	Point p1 = new Point(cellBounds.Left + cellBounds.Width * 3 / 4, cellBounds.Top + cellBounds.Height / 2);
                Point p2 = new Point(cellBounds.Left + cellBounds.Width * 1 / 2 + 15, cellBounds.Bottom - 15);
                Point p3 = new Point(cellBounds.Right - 15, cellBounds.Bottom - 15);
                path.AddPolygon(new Point[] { p1, p2, p3 });
                graphics.FillPath(Brushes.Gray, path);
            }
            else if (this.sType == SortType.DesC)
            {
                Point p1 = new Point(cellBounds.Left + cellBounds.Width * 2/3, cellBounds.Top + cellBounds.Height / 2);
                Point p3 = new Point(cellBounds.Left + cellBounds.Width * 2/3 + 18, cellBounds.Top + cellBounds.Height / 2);
                Point p2 = new Point(cellBounds.Left+cellBounds.Width*3/5+15, cellBounds.Bottom - 15);
                path.AddPolygon(new Point[] { p1, p2, p3 });
                graphics.FillPath(Brushes.Gray, path);
            }
            else
            {

            }
            }
			
        }
        protected override void OnDataGridViewChanged()
        {
            //when the column is just attached to the DataGridView, we add the TextBox and CheckBox to the dataGridView.
            if (this.DataGridView != null)
            {
            	this.DataGridView.Controls.Add(cb);
                this.
                	DataGridView.Controls.Add(tb);
            }
        }
                     
    }
	public class DataGridViewTextBoxColumnWithSpecialColumnHeaderSecoundForm : DataGridViewTextBoxColumn
    {

        TextBoxAndCheckBoxHeaderCella headerCell = new TextBoxAndCheckBoxHeaderCella();
        public DataGridViewTextBoxColumnWithSpecialColumnHeaderSecoundForm()
        {
            
            this.CellTemplate = new DataGridViewTextBoxCell();

            //assign the custom header cell instance to the HeaderCell Property of the column
            this.HeaderCell = headerCell;
			
            //hide the default sort icon for the dataGridViewColumnHeaderCell
            this.SortMode = DataGridViewColumnSortMode.Automatic;
        }
    
        protected override void OnDataGridViewChanged()
        {
          
      
        }
        public TextBoxAndCheckBoxHeaderCella HeaderCellEx
        {

            get
            {
                return this.headerCell;
            }
        }
    }
}
