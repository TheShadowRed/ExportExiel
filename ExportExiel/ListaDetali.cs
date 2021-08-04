/*
 * Created by SharpDevelop.
 * User: TheRedLord
 * Date: 3/31/2017
 * Time: 16:17
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using DataGridViewAutoFilter;

namespace ExportExiel
{
	/// <summary>
	/// Description of ListaDetali.
	/// </summary>
	public partial class ListaDetali : Form
	{
		public static int ColoumnIndex=0;
		public static ListaDetali fo;
		public static string Flag;
		public static DataGridViewEx datagridViewEx1;
		public static DataGridView dataGridView2;
		public static Form ListaDTable;
		public static Panel LocatieUndePunPanel;
		public static string loadTableQuery;
		public static List<String> listaTabDateType=new List<string>();
		public static List<String> list = new List<String>();
		public static List<String> ListaCapColoane=new List<string>();
		public static ToolStripStatusLabel filterStatusLabel = new ToolStripStatusLabel();
        public static ToolStripStatusLabel showAllLabel = new ToolStripStatusLabel("show all");
        public static ToolStripStatusLabel custom = new ToolStripStatusLabel("customRaport");
        public static StatusStrip statusStrip1 = new StatusStrip();
		public ListaDetali()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			LocatieUndePunPanel=panel3;
			ListaDTable=this;
			fo=this;
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		public void queryLoadTableFiller(DataTable dt)
		{
			
            BindingSource dataSource = new BindingSource(dt, null);
            dataGridView2.DataSource=dataSource;
            panel3.Controls.Add(dataGridView2);
		}
		void Button6Click(object sender, EventArgs e)
		{
			this.Close();
		}
		public void loadTableInt(String myStringCon,String LoadDetaliQuery,String Header,String Footer,String Body,String Query,string id_raport,String denumire,String date1,String date2)
		{
			string limitLoad="limit "+MainForm.Limit;
			using (var connP = new MySqlConnection(myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{
          						//cmdP.Parameters.Add(new MySqlParameter("@an_lucru", id_raport));
          						string theDate = dateTimePicker2.Value.ToString("yyyy-MM-dd");
          						cmdP.Parameters.AddWithValue("@data", theDate);
          						cmdP.Parameters.AddWithValue("@Id_Lista", id_raport);
          						cmdP.CommandText = LoadDetaliQuery;
          						using (MySqlDataReader readerP = cmdP.ExecuteReader())
          						{
               						while (readerP.Read())
               							{
               							//elementele de la header
               							Header=readerP["header"].ToString();
               							Footer=readerP["footer"].ToString();
               							Body=readerP["body"].ToString();
               							Query=readerP["query"].ToString();
               							
               }	
          }
     }
}
	DataSet ds = new DataSet("New_DataSet");
	DataTable dt = new DataTable(denumire);
	using (var connP = new MySqlConnection(myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{
     							int x = Int32.Parse(MainForm.Limit);
     							cmdP.Parameters.AddWithValue("@data1", date1);
     							cmdP.Parameters.AddWithValue("@data2", date2);
     							cmdP.Parameters.AddWithValue("@limit", x);
     							cmdP.CommandText = Query;
     							MainForm.loadTableQuery=Query;
          						 MySqlDataAdapter Da = new MySqlDataAdapter(cmdP);
          						 
                						Da.Fill(dt);
                						
     						}
						}
			dataGridView2 = new DataGridView();
			dataGridView2.Name = "dataGridView2";
            dataGridView2.Dock = DockStyle.Fill;
            dataGridView2.BindingContextChanged += new EventHandler(dataGridView2_BindingContextChanged);
            dataGridView2.KeyDown += new KeyEventHandler(dataGridView2_KeyDown);
            dataGridView2.DataBindingComplete +=
                new DataGridViewBindingCompleteEventHandler(
                dataGridView2_DataBindingComplete);

            showAllLabel.Visible = false;
            showAllLabel.IsLink = true;
            showAllLabel.LinkBehavior = LinkBehavior.HoverUnderline;
            showAllLabel.Click += new EventHandler(showAllLabel_Click);
 			custom.Visible = false;
            custom.IsLink = true;
            custom.LinkBehavior = LinkBehavior.HoverUnderline;
            custom.Click += new EventHandler(showAllLabel_Click);
            statusStrip1.Items.AddRange(new ToolStripItem[] { 
            filterStatusLabel, showAllLabel });


            LocatieUndePunPanel.Controls.AddRange(new Control[] {
                dataGridView2, statusStrip1 });
            BindingSource dataSource = new BindingSource(dt, null);
            dataGridView2.DataSource=dataSource;
	
		}
		private void showAllLabel_Click(object sender, EventArgs e)
        {
        	MessageBox.Show("Dot Net Perls is awesome.");
            //DataGridViewAutoFilterColumnHeaderCell.RemoveFilter(dataGridView1);
        }
		private void dataGridView2_DataBindingComplete(object sender,
            DataGridViewBindingCompleteEventArgs e)
        {
            String filterStatus = DataGridViewAutoFilterColumnHeaderCell
                .GetFilterStatus(dataGridView2);
            if (String.IsNullOrEmpty(filterStatus))
            {
                showAllLabel.Visible = false;
                filterStatusLabel.Visible = false;
            }
            else
            {
                showAllLabel.Visible = true;
                filterStatusLabel.Visible = true;
                filterStatusLabel.Text = filterStatus;
            }
        }
		private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Alt && (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up))
            {
                DataGridViewAutoFilterColumnHeaderCell filterCell =
                    dataGridView2.CurrentCell.OwningColumn.HeaderCell as
                    DataGridViewAutoFilterColumnHeaderCell;
                if (filterCell != null)
                {
                    filterCell.ShowDropDownList();
                    e.Handled = true;
                }
            }
        }
		private void dataGridView2_BindingContextChanged(object sender, EventArgs e)
        {
            // Continue only if the data source has been set.
            if (dataGridView2.DataSource == null)
            {
                return;
            }

            // Add the AutoFilter header cell to each column.
            foreach (DataGridViewColumn col in dataGridView2.Columns)
            {
                col.HeaderCell = new
                    DataGridViewAutoFilterColumnHeaderCell(col.HeaderCell);
            }

            // Format the OrderTotal column as currency. 
            //dataGridView2.Columns["OrderTotal"].DefaultCellStyle.Format = "c";

            // Resize the columns to fit their contents.
            dataGridView2.AutoResizeColumns();
        }
		public void QueryAgain()
		{
			//String testtime="System.TimeSpan";
			//String testdatetime="MySql.Data.Types.MySqlDateTime";
			String FristPArt="";
			String test=listaTabDateType[1];
			if(list.Count==0)
			{
				
			}else{
			FristPArt=" and ";
			String Frist=list[0];
			string[] FristToken=Frist.Split('-');
			int Toke=int.Parse(FristToken[1]);
			String ColoumName=ListaCapColoane[Toke].ToString();
			if("System.String"==listaTabDateType[Toke])
			{
			FristPArt=FristPArt+ColoumName+" like '%"+FristToken[2]+"%'";
			}else
			{
				if("System.TimeSpan"==listaTabDateType[Toke])
				{
					
				}else
				{
					if("MySql.Data.Types.MySqlDateTime"==listaTabDateType[Toke])
					{
						FristPArt=FristPArt+ColoumName+" = '"+FristToken[2]+"'";
					}else
					{
						FristPArt=FristPArt+ColoumName+" = '"+FristToken[2]+"'";
					}
				}
			}
			for (var i = 1; i < list.Count; i++) {
				FristPArt=FristPArt+" and ";
				String Data=list[i];
				string[] tokens = Data.Split('-');
				int Tokea=int.Parse(tokens[1]);
				String ColoumNameLast=ListaCapColoane[Tokea].ToString();
				FristPArt=FristPArt+ColoumNameLast+" like '%"+tokens[2]+"%'";
			}
			}
DataSet ds = new DataSet("New_DataSet");
DataTable dt = new DataTable("New_DataTable");
			using (var connP = new MySqlConnection(MainForm.myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{	
     							
          						cmdP.CommandText = loadTableQuery+FristPArt;
          						cmdP.Parameters.AddWithValue("@flag", Flag);
          						MySqlDataAdapter Da = new MySqlDataAdapter(cmdP);
          						 
                						Da.Fill(dt);
     						}
						}
			
           	  datagridViewEx1.DataSource=dt;
				
			//dataGridView1.Columns[5].Visible = false;
			
		}
		void Button1Click(object sender, EventArgs e)
		{
			datagridViewEx1.Invalidate();
		}
		void TextBox1Leave(object sender, EventArgs e)
		{
	MainForm.Limit=textBox1.Text.ToString();
		}
		void DateTimePicker1Leave(object sender, EventArgs e)
		{
				MainForm.data1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
				MainForm.data2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
		}
		void DateTimePicker2Leave(object sender, EventArgs e)
		{
				MainForm.data1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
				MainForm.data2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
		}
		void ListaDetaliActivated(object sender, EventArgs e)
		{
				MainForm.data1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
				MainForm.data2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
		}
	}
}
