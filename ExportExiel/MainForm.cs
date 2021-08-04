/*
 * Created by SharpDevelop.
 * User: TheRedLord
 * Date: 3/6/2017
 * Time: 10:30
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Timers;
using System.Windows.Forms;
using ClosedXML.Excel;
using DataGridViewAutoFilter;
using MySql.Data.MySqlClient; 
using System.Configuration;
namespace ExportExiel
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		public ExportExiel.DataGridViewEx dataGridView1;
		public static  System.Windows.Forms.DataGridView dataGridView2;
        StatusStrip statusStrip1 = new StatusStrip();
        ToolStripStatusLabel filterStatusLabel = new ToolStripStatusLabel();
        ToolStripStatusLabel showAllLabel = new ToolStripStatusLabel("show all");
        ToolStripStatusLabel custom = new ToolStripStatusLabel("custom");
        public static string Limit="2500";
		public static int ColoumnIndex=0;
        public static string data1;
        public string denumiremail="";
        public System.Timers.Timer aTimer;
        public static string data2;
	    public string Blank="";
	    public string TrimiteMailTrue="0";
	    public static string produs = "";
	    public string ErrorMAil="";
	    public static string schimb="";
	    int today=0;
	    public string destinatie="";
		//public static DataGridViewEx dataGridView1;
		public static MainForm fo;
		public List<String> list = new List<String>();
		public List<String> ListaCapColoane=new List<string>();
		public List<String> listaTabDateType=new List<string>();
		public static String myStringCon;
		public static String loadTableQuery;
		public String LoadDetaliQuery;
		public String SirQueryLoadSort;
		public bool didit=false;
		public static String ID;
		//string baza="";
		//string versiune="";
		//public static string flag="";
		//string host="";
		
		public static string baza="";
		public static string vesiune="";
		public static string host="";
		public static string cod="";
		public static string locatie="";
		public static string flag="";
		
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			label1.Font=new Font(label1.Font.Name, 12, FontStyle.Bold);
			button2.Font=new Font(label1.Font.Name, 12, FontStyle.Bold);
			button6.Font=new Font(label1.Font.Name, 12, FontStyle.Bold);
			button7.Font=new Font(label1.Font.Name, 12, FontStyle.Bold);
			button9.Font=new Font(label1.Font.Name, 12, FontStyle.Bold);
			label2.Font=new Font(label1.Font.Name, 12, FontStyle.Bold);
			//comboBox1.Font=new Font(label1.Font.Name, 12, FontStyle.Bold);
			dateTimePicker1.Font=new Font(label1.Font.Name, 12, FontStyle.Bold);
			dateTimePicker2.Font=new Font(label1.Font.Name, 12, FontStyle.Bold);
			GetString();
			loadinitQuery();
			fo=this;//30*60*1000
			//System.Timers.Timer timer = new System.Timers.Timer(10 * 10 * 1000);
			//timer.Elapsed += OnTick; // Which can also be written as += new ElapsedEventHandler(OnTick);
			//timer.Start();
			aTimer = new System.Timers.Timer();
    		aTimer.Elapsed+=new ElapsedEventHandler(OnTimedEvent);
   		 	aTimer.Interval=20000;
    		aTimer.Enabled=true;
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		public void generateForSendMail()
		{
			foreach (DataGridViewRow row in dataGridView2.Rows)
{
			String date1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
			String date2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
			String Header="";
			String Footer="";
			String Body="";
			string Query ="";
			int rowCount=0;
	String id_raport=row.Cells[0].Value.ToString();
	String Flag=row.Cells[6].Value.ToString();
	String denumire=row.Cells[2].Value.ToString();
	String tipMail=row.Cells[6].Value.ToString();
	String whoMail=row.Cells[8].Value.ToString();
	TrimiteMailTrue=tipMail;
	
	if(tipMail=="1"){
	switch (Flag)
{
   case "P1"://fabrica
			LoadProlyteP(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
   case "L"://laborator
      		LoadsPITALL(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
   default://default
      		LoadDefault(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
}
	
	
	//send mail function
	 string theDatenowa = DateTime.Now.ToString("yyyy-MM-dd");
	if(File.Exists(destinatie+denumiremail+" "+theDatenowa+".xlsx")){
	trimiteMail(tipMail,whoMail);	
	 }else
	 {
	 	
	 }
	}
		
}
			aTimer.Enabled=true;
		}
		private void OnTimedEvent(object source, ElapsedEventArgs e)
 {
			String t1 = DateTime.Parse("12:24:00.000").ToString("THH:mm:");
           String t2=DateTime.Now.ToString("THH:mm:");
			if (t1 == t2)
			{
				aTimer.Enabled=false;
				if(didit==false){
				didit=true;
				generateForSendMail();
				}
			}
 }
		private void OnTick(object source, ElapsedEventArgs e)
		{ 
			TimeSpan start = new TimeSpan(10, 0, 0); //10 o'clock
			TimeSpan end = new TimeSpan(12, 0, 0); //12 o'clock
			TimeSpan now = DateTime.Now.TimeOfDay;

			if ((now > start) && (now < end))
			{
				generateForSendMail();
   				
			}
		}
		public void loadinitQuery()
		{
			String query="";
				using (var connP = new MySqlConnection(myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{
          						//cmdP.Parameters.Add(new MySqlParameter("@an_lucru", id_raport));
          						cmdP.CommandText = "select query from raport_detali where id_lista='"+1+"'";
          						using (MySqlDataReader readerP = cmdP.ExecuteReader())
          						{
               						while (readerP.Read())
               							{
               							//elementele de la header
               							query=readerP["query"].ToString();
               							
               }	
          }
     }
}
				string[] lines = query.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
				loadTableQuery=lines[0];
               		
               	LoadDetaliQuery=lines[1];
				loadTableInt();
		}
		public void oneMoreUnhide()
		{
			int TotalRows=dataGridView1.RowCount-1;
			for(int CurrentInde=0;CurrentInde<TotalRows;CurrentInde++)
            {
            		CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dataGridView1.DataSource];  
					currencyManager1.SuspendBinding();
					dataGridView1.Rows[CurrentInde].Visible=true;
					currencyManager1.ResumeBinding();
            }
		}
		public void MasterSort()
		{
			for (var i = 0; i < list.Count; i++) {
				String Data=list[i];
				string[] tokens = Data.Split('-');
				int Toke=int.Parse(tokens[1]);
				oneMoreSort(Toke,tokens[2]);
			}
		}
		public void oneMoreSort(int ColoumIndex,String CheckData)
        {
			int TotalRows=dataGridView1.RowCount-1;
			for(int CurrentInde=0;CurrentInde<TotalRows;CurrentInde++)
            {
            	String data=dataGridView1.Rows[CurrentInde].Cells[ColoumIndex].Value.ToString();
            	if(data.Contains(CheckData))
            	{
            		//e ok 
            	}else
            	{
            		//hide
            		CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dataGridView1.DataSource];  
					currencyManager1.SuspendBinding();
					dataGridView1.Rows[CurrentInde].Visible=false;
					currencyManager1.ResumeBinding();
            		
            	}
            }
			
        }
		public void GetString()
		{
			
		 baza=System.Configuration.ConfigurationManager.AppSettings["baza"].ToString();
		 vesiune=System.Configuration.ConfigurationManager.AppSettings["vesiune"].ToString();
		 host=System.Configuration.ConfigurationManager.AppSettings["host"].ToString();
		 cod=System.Configuration.ConfigurationManager.AppSettings["cod"].ToString();
		 locatie=System.Configuration.ConfigurationManager.AppSettings["locatie"].ToString();
		 flag=System.Configuration.ConfigurationManager.AppSettings["flag"].ToString();
		
        myStringCon = "SERVER=" + host + ";" +
                 "DATABASE=" + baza + ";" +
                 "UID=florin1; PASSWORD=123321" + ";pooling=true;Allow Zero Datetime=True;Min Pool Size=1; Max Pool Size=100; default command timeout=120";
            
               		 	loadTableQuery=System.Configuration.ConfigurationManager.AppSettings["loadTableQuery"].ToString();
               		
               		 	LoadDetaliQuery=System.Configuration.ConfigurationManager.AppSettings["LoadDetaliQuery"].ToString();
               		 
               
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
			using (var connP = new MySqlConnection(myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{	
     							
          						cmdP.CommandText = loadTableQuery+FristPArt;
          						cmdP.Parameters.AddWithValue("@flag", flag);
          						MySqlDataAdapter Da = new MySqlDataAdapter(cmdP);
          						 
                						Da.Fill(dt);
     						}
						}
			
           	  dataGridView1.DataSource=dt;
				
			//dataGridView1.Columns[5].Visible = false;
			
		}
		public void loadTableInt()
		{
			//schimbs
			SirQueryLoadSort=" where flag='"+flag+"'";
		DataSet ds = new DataSet("New_DataSet");
DataTable dt = new DataTable("New_DataTable");
			using (var connP = new MySqlConnection(myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{
     							
          						cmdP.CommandText = loadTableQuery+SirQueryLoadSort;
          						//cmdP.Parameters.AddWithValue("@schimb", flag);
          						cmdP.Parameters.AddWithValue("@flag", flag);
          						MySqlDataAdapter Da = new MySqlDataAdapter(cmdP);
          						 
                						Da.Fill(dt);
     						}
						}
			dataGridView2 = new DataGridView();
			dataGridView2.Name = "dataGridView2";
            dataGridView2.Dock = DockStyle.Fill;
            dataGridView2.CellClick +=dataGridView1_CellClick;
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
            statusStrip1.Cursor = Cursors.Default;
            statusStrip1.Items.AddRange(new ToolStripItem[] { 
                filterStatusLabel, showAllLabel });


            this.Text = "Raportare";
            this.Width *= 3;
            this.Height *= 2;
            panel3.Controls.AddRange(new Control[] {
                dataGridView2, statusStrip1 });
            BindingSource dataSource = new BindingSource(dt, null);
            dataGridView2.DataSource=dataSource;
            dataGridView2.Columns[0].MinimumWidth=120;
            dataGridView2.Columns[1].MinimumWidth=220;
            dataGridView2.Columns[2].MinimumWidth=320;
            dataGridView2.Columns[3].MinimumWidth=220;
            dataGridView2.Columns[4].Visible = false;
            dataGridView2.Columns[5].Visible = false;
            dataGridView2.Columns[6].Visible = false;
            dataGridView2.Columns[7].Visible = false;
            dataGridView2.Columns[8].Visible = false;
            dataGridView2.Columns[9].Visible = false;
            dataGridView2.Columns[10].Visible = false;
           
		}
		private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)

{
			foreach (DataGridViewRow row in dataGridView2.Rows)
            {

                
                    row.DefaultCellStyle.BackColor = Color.White;
					//row.DefaultCellStyle.ForeColor = Color.Blue;                 

            }
	int b = dataGridView2.CurrentCellAddress.Y;
	String datatip=dataGridView2.Rows[b].Cells[9].Value.ToString();
	dataGridView2.Rows[b].DefaultCellStyle.BackColor = Color.Yellow;
	if(datatip=="1")
	{
		label1.Visible=false;
		dateTimePicker1.Visible = false;
		label2.Text="Data";
	}
	else
	{
		label1.Visible=true;
		dateTimePicker1.Visible = true;
		label2.Text="Data Sfarsit";
	}
	String schimb=dataGridView2.Rows[b].Cells[10].Value.ToString();
	if(schimb!="1")
	{
		groupBox1.Visible=false;
	}
	else
	{
		groupBox1.Visible=true;
	}
}
		private void Form1_Load(object sender, EventArgs e)
        {
            
        }
		    void HeaderCellEx_TextBoxValueChanged(object sender, EventArgs e)
        {
            MessageBox.Show("TextBoxValueChanged");
        }
		
		void Button1Click(object sender, EventArgs e)
		{
			
DataSet ds = new DataSet("New_DataSet");
DataTable dt = new DataTable("New_DataTable");
	using (var connP = new MySqlConnection(myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{
          						cmdP.CommandText = "select * from aTemp";
          						 MySqlDataAdapter Da = new MySqlDataAdapter(cmdP);
          						 
                						Da.Fill(dt);
                						
     }
}
 var wb = new XLWorkbook();

            // Add a DataTable as a worksheet
            wb.Worksheets.Add(dt);
            wb.SaveAs("AddingDataTableAsWorksheet.xlsx");
             String Temporal="";
             String TemporalNr_crt="";
             String sef_schimb="";
             String Calitate="";
             String Sudor="";
             String Cordonator="";
             //add coloum
              var workbookAddCOloun = new XLWorkbook("AddingDataTableAsWorksheet.xlsx");
            var wsAdd = workbookAddCOloun.Worksheets.Worksheet("New_DataTable");
            int vectorToAdd=0;
             for(int a=0;a<dt.Rows.Count;a++)
             {
             	String FristElement="";
             	String SecoundElement="";
             	try
             	{
             		FristElement=dt.Rows[a][1].ToString();
             	}catch
             	{
             		
             	}
             	int b=1+a;
             	try
             	{
             		SecoundElement=dt.Rows[b][1].ToString();
             	}catch
             	{
             		
             	}
             	if(FristElement!=SecoundElement&&FristElement!=""&&SecoundElement!="")
             	{
             		int add=3+a+vectorToAdd;
             		wsAdd.Range("A"+add+":M"+add).InsertRowsAbove(2);
					vectorToAdd=vectorToAdd+2;
					add=3+a+vectorToAdd;
					string ValueToCopy=wsAdd.Cell("B"+add).Value.ToString();
					string SchimbToCopy=wsAdd.Cell("A"+add).Value.ToString();
					int addminus=add-1;
					int adminus=add-2;
					wsAdd.Cells("B"+addminus).Value=ValueToCopy;
					wsAdd.Cells("B"+adminus).Value=ValueToCopy;
					wsAdd.Cells("A"+addminus).Value=SchimbToCopy;
					wsAdd.Cells("A"+adminus).Value=SchimbToCopy;
					//COPY COL A AND B
             	}
             	
             	
             }
             workbookAddCOloun.SaveAs("AddingDataTableAsWorksheet.xlsx");
            
       
             //merge
            var workbook = new XLWorkbook("AddingDataTableAsWorksheet.xlsx");
            var ws = workbook.Worksheets.Worksheet("New_DataTable");
            //int lasta=0;
            int maxLenght=0;
            //int indexadaugari=0;
            //String ADDTWOROWS="";
            String lastPossibleAddress = ws.LastCellUsed().Address.ToString();
            Regex regexObj = new Regex(@"[^\d]");
            String resultString = regexObj.Replace(lastPossibleAddress, "");
            int x = Int32.Parse(resultString);
            x=x+1;
            for(int a=1;a<x;a++)
            {
            	String FristElement="";
             	String SecoundElement="";
             	try
             	{
             		FristElement=ws.Cell("B"+a).Value.ToString();
             	}catch
             	{
             		
             	}
             	int b=a-1;
             	try
             	{
             		SecoundElement=ws.Cell("B"+b).Value.ToString();
             	}catch
             	{
             		
             	}
             	if(FristElement==SecoundElement)
             	{
             		maxLenght++;
             	}else
             	{
             		if(maxLenght!=0)
             		{
             		int frist=a-maxLenght-1;
            		int LastMod=a-1;
            		ws.Cell("B"+frist).Value = Temporal;
            		ws.Cell("B"+frist).Style.Alignment.WrapText = true;
            		ws.Range("B"+frist+":C"+LastMod).Column(1).Merge();
            		ws.Cell("J"+frist).Value = sef_schimb;
            		ws.Cell("J"+frist).Style.Alignment.WrapText = true;
            		ws.Range("J"+frist+":K"+LastMod).Column(1).Merge();
            		ws.Cell("K"+frist).Value = Calitate;
            		ws.Cell("K"+frist).Style.Alignment.WrapText = true;
            		ws.Range("K"+frist+":L"+LastMod).Column(1).Merge();
            		ws.Cell("L"+frist).Value = Sudor;
            		ws.Cell("L"+frist).Style.Alignment.WrapText = true;
            		ws.Range("L"+frist+":M"+LastMod).Column(1).Merge();
            		ws.Cell("M"+frist).Value = Cordonator;
            		ws.Cell("M"+frist).Style.Alignment.WrapText = true;
            		ws.Range("M"+frist+":N"+LastMod).Column(1).Merge();
            		ws.Cell("A"+frist).Value = TemporalNr_crt;
            		ws.Cell("A"+frist).Style.Alignment.WrapText = true;
            		ws.Range("A"+frist+":A"+LastMod).Column(1).Merge();
            		maxLenght=0;
             		}
             		Temporal=ws.Cell("B"+a).Value.ToString();
            		TemporalNr_crt=ws.Cell("A"+a).Value.ToString();
            		sef_schimb=ws.Cell("J"+a).Value.ToString();
            		Calitate=ws.Cell("K"+a).Value.ToString();
            		Sudor=ws.Cell("L"+a).Value.ToString();
            		Cordonator=ws.Cell("M"+a).Value.ToString();
             	}
            }
            int ad=78;
            if(maxLenght!=0)
             		{
             		int frist=ad-maxLenght;
            		int LastMod=ad;
            		ws.Cell("B"+frist).Value = Temporal;
            		ws.Cell("B"+frist).Style.Alignment.WrapText = true;
            		ws.Range("B"+frist+":C"+LastMod).Column(1).Merge();
            		ws.Cell("J"+frist).Value = sef_schimb;
            		ws.Cell("J"+frist).Style.Alignment.WrapText = true;
            		ws.Range("J"+frist+":K"+LastMod).Column(1).Merge();
            		ws.Cell("K"+frist).Value = Calitate;
            		ws.Cell("K"+frist).Style.Alignment.WrapText = true;
            		ws.Range("K"+frist+":L"+LastMod).Column(1).Merge();
            		ws.Cell("L"+frist).Value = Sudor;
            		ws.Cell("L"+frist).Style.Alignment.WrapText = true;
            		ws.Range("L"+frist+":M"+LastMod).Column(1).Merge();
            		ws.Cell("M"+frist).Value = Cordonator;
            		ws.Cell("M"+frist).Style.Alignment.WrapText = true;
            		ws.Range("M"+frist+":N"+LastMod).Column(1).Merge();
            		ws.Cell("A"+frist).Value = TemporalNr_crt;
            		ws.Cell("A"+frist).Style.Alignment.WrapText = true;
            		ws.Range("A"+frist+":A"+LastMod).Column(1).Merge();
            		maxLenght=0;
             		}
            
          
              workbook.SaveAs("AddingDataTableAsWorksheet.xlsx");
            /*
            for(int a=0;a<;a++)
            {
            	//Adauga Doua Randrui
            	
            	//merge cnp
            	if(Temporal==dt.Rows[a][1].ToString())
            	{
            		maxLenght++;	lasta=a;
            		
            	}else
            	{
            		if(maxLenght==0)
            		{
            			
            		}
            		if(maxLenght!=0){
            		int frist=2+indexadaugari+lasta-maxLenght;
            		int LastMod=indexadaugari+lasta+2;
            		ws.Cell("B"+frist).Value = Temporal;
            		ws.Cell("B"+frist).Style.Alignment.WrapText = true;
            		ws.Range("B"+frist+":C"+LastMod).Column(1).Merge();
            		ws.Cell("J"+frist).Value = sef_schimb;
            		ws.Cell("J"+frist).Style.Alignment.WrapText = true;
            		ws.Range("J"+frist+":K"+LastMod).Column(1).Merge();
            		ws.Cell("K"+frist).Value = Calitate;
            		ws.Cell("K"+frist).Style.Alignment.WrapText = true;
            		ws.Range("K"+frist+":L"+LastMod).Column(1).Merge();
            		ws.Cell("L"+frist).Value = Sudor;
            		ws.Cell("L"+frist).Style.Alignment.WrapText = true;
            		ws.Range("L"+frist+":M"+LastMod).Column(1).Merge();
            		ws.Cell("M"+frist).Value = Cordonator;
            		ws.Cell("M"+frist).Style.Alignment.WrapText = true;
            		ws.Range("M"+frist+":N"+LastMod).Column(1).Merge();
            		ws.Cell("A"+frist).Value = TemporalNr_crt;
            		ws.Cell("A"+frist).Style.Alignment.WrapText = true;
            		ws.Range("A"+frist+":A"+LastMod).Column(1).Merge();
            		
            		}
            		maxLenght=0;
            		Temporal=dt.Rows[a][1].ToString();
            		TemporalNr_crt=dt.Rows[a][0].ToString();
            		sef_schimb=dt.Rows[a][9].ToString();
            		Calitate=dt.Rows[a][10].ToString();
            		Sudor=dt.Rows[a][11].ToString();
            		Cordonator=dt.Rows[a][12].ToString();
            		
            	}
            }
            //merge cnp final
            if(maxLenght!=0)
            {
            		int frist=2+lasta-maxLenght;
            		int LastMod=lasta+2;
            		ws.Cell("B"+frist).Value = Temporal;
            		ws.Cell("B"+frist).Style.Alignment.WrapText = true;
            		ws.Range("B"+frist+":C"+LastMod).Column(1).Merge();
            		ws.Cell("A"+frist).Value = TemporalNr_crt;
            		ws.Cell("A"+frist).Style.Alignment.WrapText = true;
            		ws.Range("A"+frist+":A"+LastMod).Column(1).Merge();
            		
            }
            
            // From a list of strings
            
          
            
            
           
            
            
            
            
            
            
            
            
            //merge
            //var workbook = new XLWorkbook("AddingDataTableAsWorksheet.xlsx");
            //var ws = workbook.Worksheets.Worksheet("New_DataTable");
            //ws.Cell("F2").Value = "Merged Column(1) of Range (F2:G8)";
            //ws.Cell("F2").Style.Alignment.WrapText = true;
            //ws.Range("F2:G8").Column(1).Merge();
            workbook.SaveAs("AddingDataTableAsWorksheet.xlsx");
            //SelectSpecificMerge
            var workBook= new XLWorkbook("AddingDataTableAsWorksheet.xlsx");
            var wsMerge = workbook.Worksheets.Worksheet("New_DataTable");
             */
            
            
             
		}
		void MainFormLoad(object sender, EventArgs e)
		{
	
		}
		void DataGridView1MouseClick(object sender, MouseEventArgs e)
		{
			
		}
		void Button2Click(object sender, EventArgs e)
		{
			
			String date1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
			String date2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
			String Header="";
			String Footer="";
			String Body="";
			string Query ="";
			int rowCount=0;
	int b = dataGridView2.CurrentCellAddress.Y;
	String id_raport=dataGridView2.Rows[b].Cells[0].Value.ToString();
	String Flag=dataGridView2.Rows[b].Cells[5].Value.ToString();
	String denumire=dataGridView2.Rows[b].Cells[2].Value.ToString();
	denumiremail=denumire;
	String trimitemail=dataGridView2.Rows[b].Cells[7].Value.ToString();
	String mailTip=dataGridView2.Rows[b].Cells[6].Value.ToString();
	switch (Flag)
{
   case "P1"://fabrica
			LoadProlyteP(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
   case "L"://laborator
      		LoadsPITALL(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
   default://default
      		LoadDefault(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
}
	   string theDatenowa = DateTime.Now.ToString("yyyy-MM-dd");
	   if(File.Exists(destinatie+denumire+" "+theDatenowa+".xlsx")){
             System.Diagnostics.Process.Start(destinatie+denumire+" "+theDatenowa+".xlsx");
	   }
	//if(trimitemail=="1")
	//{
	//	trimiteMail(mailTip);
	//}
	
	
	
	
	
	
	
	
	
            
         
		}
		public void LoadDefault(String date1,String date2,String Header,String Footer,String Body,string Query,int rowCount,String id_raport,String Denumire)
		{
	
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
               							Blank=readerP["blank"].ToString();
               							Header=readerP["header"].ToString();
               							Footer=readerP["footer"].ToString();
               							Body=readerP["body"].ToString();
               							Query=readerP["query"].ToString();
               							destinatie=readerP["destinatie"].ToString();
               }	
          }
     }
}
	DataSet ds = new DataSet("New_DataSet");
	
	DataTable dt = new DataTable("Raport");
	using (var connP = new MySqlConnection(myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{
     							
     							cmdP.Parameters.AddWithValue("@data1", date1);
     							cmdP.Parameters.AddWithValue("@data2", date2);
     							cmdP.CommandText = Query;
          						 MySqlDataAdapter Da = new MySqlDataAdapter(cmdP);
          						 
                						Da.Fill(dt);
                						
{
    //
}	
     						}
						}
	
	if (dt.Rows.Count == 0)
	{//nu are randuri
		if(TrimiteMailTrue=="1"){ErrorMAil="Nu sunt date in intervalul de timp";
			TrimiteMailTrue="2";
		
		}else{
    MessageBox.Show("Nu sunt date in intervalul de timp");
		}
	}else{
		if(Blank!="")
		{
			
			var workbook = new XLWorkbook(Blank);
			
            workbook.Worksheet(1).Cell(5, 1).InsertTable(dt);
            //ws.Cell("F2").Value = "Merged Column(1) of Range (F2:G8)";
            //ws.Cell("F2").Style.Alignment.WrapText = true;
            //ws.Range("F2:G8").Column(1).Merge();
            //workbook.SaveAs("AddingDataTableAsWorksheet.xlsx");
            //SelectSpecificMerge
            //var workBook= new XLWorkbook("AddingDataTableAsWorksheet.xlsx");
            //var wsMerge = workbook.Worksheets.Worksheet(Denumire);
            string theDatenow = DateTime.Now.ToString("yyyy-MM-dd");
            workbook.SaveAs(destinatie+Denumire+" "+theDatenow+".xlsx");
		}else{
			//Add the table to the data set
			var wb = new XLWorkbook();
            // Add a DataTable as a worksheet
            wb.Worksheets.Add(dt);
            string theDatenow = DateTime.Now.ToString("yyyy-MM-dd");
            int extraRowCount=1;
            if(Header!="")
            {
            var ws = wb.Worksheet("Raport");
            ws.Range("A1:AG1").InsertRowsAbove(2);
            ws.Cell("A1").Value =Header;
            extraRowCount++;
            extraRowCount++;
            }
            if(Body!="")
            {
            var ws = wb.Worksheet("Raport");
            ws.Range("A"+extraRowCount+":AG"+extraRowCount).InsertRowsAbove(1);
            ws.Cell("A"+extraRowCount).Value =Body;
            extraRowCount++;
            }
            int rowsa = dt.Rows.Count;
            int aasd=rowsa+extraRowCount;
            var wsa = wb.Worksheet("Raport");
            wsa.Range("A1:H"+aasd).Style.Border.OutsideBorder=XLBorderStyleValues.Medium;
            wsa.Range("A1:H"+aasd).Style.Border.DiagonalBorderColor = XLColor.Red;
            if(Footer!="")
            {
            	int rows = dt.Rows.Count;
            	rows=rows+extraRowCount+1;
            var ws = wb.Worksheet("Raport");
            ws.Range("A"+rows+":AG"+rows).InsertRowsAbove(2);
            rows=rows+1;
            ws.Cell("A"+rows).Value =Footer;
          	
            }
            //black contur
           
            // save xlsx
            denumiremail=Denumire;
            wb.SaveAs(destinatie+Denumire+" "+theDatenow+".xlsx");
		}
	}
		}
		public void LoadProlyteP(String date1,String date2,String Header,String Footer,String Body,string Query,int rowCount,String id_raport,String Denumire)
		{
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
               							destinatie=readerP["destinatie"].ToString();
               							}	
         						}
     						}
						}
		
		for(int a=0;a<5;a++)
		{
			
		}
		
		
		
		
		
		
		
		
		
		
		
		
DataSet ds = new DataSet("New_DataSet");
DataTable dt = new DataTable(Denumire);
	using (var connP = new MySqlConnection(myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{
     							cmdP.Parameters.AddWithValue("@data1", date1);
     							cmdP.Parameters.AddWithValue("@data2", date2);
     							cmdP.CommandText = Query;
          						 MySqlDataAdapter Da = new MySqlDataAdapter(cmdP);
          						 
                						Da.Fill(dt);
                						
     }
}
	/*Adauga nr_crt la date
	int nrcount=1;
     for (int r = 0; r < dt.Rows.Count; r++)
{
     	if("Total"==dt.Rows[r][0].ToString())
     	{
     		nrcount++;
     	}else
     	{
    dt.Rows[r].SetField("nr_crt", nrcount);
     	}
}
		*/
		//Add the table to the data set
 var wb = new XLWorkbook();

            // Add a DataTable as a worksheet
            wb.Worksheets.Add(dt);
              string theDatenow = DateTime.Now.ToString("yyyy-MM-dd");
            // save xlsx
            wb.SaveAs(destinatie+Denumire+" "+theDatenow+".xlsx");
             String Temporal="";
             String TemporalNr_crt="";
            var workbook = new XLWorkbook(destinatie+Denumire+" "+theDatenow+".xlsx");
            var ws = workbook.Worksheets.Worksheet(Denumire);
            int lasta=0;
            int maxLenght=0;
            //int nrCRTindex=1;
            for(int a=0;a<dt.Rows.Count;a++)
            {
            	
            	//merge cnp
            	if(Temporal==dt.Rows[a][1].ToString())
            	{
            		maxLenght++;	lasta=a;
            	}else
            	{
            		if(maxLenght!=0){
            		
            		int frist=2+lasta-maxLenght;
            		int LastMod=lasta+2;
            		ws.Cell("B"+frist).Value = Temporal;
            		ws.Cell("B"+frist).Style.Alignment.WrapText = true;
            		ws.Range("B"+frist+":C"+LastMod).Column(1).Merge();
            		ws.Cell("A"+frist).Value = TemporalNr_crt;
            		ws.Cell("A"+frist).Style.Alignment.WrapText = true;
            		ws.Range("A"+frist+":A"+LastMod).Column(1).Merge();
            		
            		}
            		maxLenght=0;
            		Temporal=dt.Rows[a][1].ToString();
            		TemporalNr_crt=dt.Rows[a][0].ToString();
            	}
            	rowCount++;
            }
            //merge cnp final
            if(maxLenght!=0)
            {
            		int frist=2+lasta-maxLenght;
            		int LastMod=lasta+2;
            		ws.Cell("B"+frist).Value = Temporal;
            		ws.Cell("B"+frist).Style.Alignment.WrapText = true;
            		ws.Range("B"+frist+":C"+LastMod).Column(1).Merge();
            		ws.Cell("A"+frist).Value = TemporalNr_crt;
            		ws.Cell("A"+frist).Style.Alignment.WrapText = true;
            		ws.Range("A"+frist+":A"+LastMod).Column(1).Merge();
            		
            }
            // add Body
            ws.Range("A1:F1").InsertRowsAbove(3);
            ws.Cell("A1").Value = Body;
            ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell("A1").Style.Alignment.WrapText = true;
            ws.Range("A1:P3").Merge();
            
            
           
            
            
            
            
            
            
            
            
            //merge
            //var workbook = new XLWorkbook("AddingDataTableAsWorksheet.xlsx");
            //var ws = workbook.Worksheets.Worksheet("New_DataTable");
            //ws.Cell("F2").Value = "Merged Column(1) of Range (F2:G8)";
            //ws.Cell("F2").Style.Alignment.WrapText = true;
            //ws.Range("F2:G8").Column(1).Merge();
              string theDatenowa = DateTime.Now.ToString("yyyy-MM-dd");
            // save xlsx
            wb.SaveAs(destinatie+Denumire+" "+theDatenowa+".xlsx");
            //SelectSpecificMerge
            var workBook= new XLWorkbook(destinatie+Denumire+" "+theDatenowa+".xlsx");
            var wsMerge = workbook.Worksheets.Worksheet(Denumire);
         
		}
		public void LoadProlytePL(String date1,String date2,String header,String Footer,String Body,String Query,int rowCount,String id_raport,String Denumire)
		{
					
DataSet ds = new DataSet("New_DataSet");
DataTable dt = new DataTable("New_DataTable");
	using (var connP = new MySqlConnection(myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{
          						cmdP.CommandText =Query;
          						 MySqlDataAdapter Da = new MySqlDataAdapter(cmdP);
          						 
                						Da.Fill(dt);
                						
     }
}
 var wb = new XLWorkbook();

            // Add a DataTable as a worksheet
            wb.Worksheets.Add(dt);
            wb.SaveAs("AddingDataTableAsWorksheet.xlsx");
             String Temporal="";
             String TemporalNr_crt="";
             String sef_schimb="";
             String Calitate="";
             String Sudor="";
             String Cordonator="";
             //add coloum
              var workbookAddCOloun = new XLWorkbook("AddingDataTableAsWorksheet.xlsx");
            var wsAdd = workbookAddCOloun.Worksheets.Worksheet("New_DataTable");
            int vectorToAdd=0;
             for(int a=0;a<dt.Rows.Count;a++)
             {
             	String FristElement="";
             	String SecoundElement="";
             	try
             	{
             		FristElement=dt.Rows[a][1].ToString();
             	}catch
             	{
             		
             	}
             	int b=1+a;
             	try
             	{
             		SecoundElement=dt.Rows[b][1].ToString();
             	}catch
             	{
             		
             	}
             	if(FristElement!=SecoundElement&&FristElement!=""&&SecoundElement!="")
             	{
             		int add=3+a+vectorToAdd;
             		wsAdd.Range("A"+add+":M"+add).InsertRowsAbove(2);
					vectorToAdd=vectorToAdd+2;
					add=3+a+vectorToAdd;
					string ValueToCopy=wsAdd.Cell("B"+add).Value.ToString();
					string SchimbToCopy=wsAdd.Cell("A"+add).Value.ToString();
					int addminus=add-1;
					int adminus=add-2;
					wsAdd.Cells("B"+addminus).Value=ValueToCopy;
					wsAdd.Cells("B"+adminus).Value=ValueToCopy;
					wsAdd.Cells("A"+addminus).Value=SchimbToCopy;
					wsAdd.Cells("A"+adminus).Value=SchimbToCopy;
					//COPY COL A AND B
             	}
             	
             	
             }
             workbookAddCOloun.SaveAs("AddingDataTableAsWorksheet.xlsx");
            
       
             //merge
            var workbook = new XLWorkbook("AddingDataTableAsWorksheet.xlsx");
            var ws = workbook.Worksheets.Worksheet("New_DataTable");
            //int lasta=0;
            int maxLenght=0;
            //int indexadaugari=0;
            //String ADDTWOROWS="";
            String lastPossibleAddress = ws.LastCellUsed().Address.ToString();
            Regex regexObj = new Regex(@"[^\d]");
            String resultString = regexObj.Replace(lastPossibleAddress, "");
            int x = Int32.Parse(resultString);
            x=x+1;
            for(int a=1;a<x;a++)
            {
            	String FristElement="";
             	String SecoundElement="";
             	try
             	{
             		FristElement=ws.Cell("B"+a).Value.ToString();
             	}catch
             	{
             		
             	}
             	int b=a-1;
             	try
             	{
             		SecoundElement=ws.Cell("B"+b).Value.ToString();
             	}catch
             	{
             		
             	}
             	if(FristElement==SecoundElement)
             	{
             		maxLenght++;
             	}else
             	{
             		if(maxLenght!=0)
             		{
             		int frist=a-maxLenght-1;
            		int LastMod=a-1;
            		ws.Cell("B"+frist).Value = Temporal;
            		ws.Cell("B"+frist).Style.Alignment.WrapText = true;
            		ws.Range("B"+frist+":C"+LastMod).Column(1).Merge();
            		ws.Cell("J"+frist).Value = sef_schimb;
            		ws.Cell("J"+frist).Style.Alignment.WrapText = true;
            		ws.Range("J"+frist+":K"+LastMod).Column(1).Merge();
            		ws.Cell("K"+frist).Value = Calitate;
            		ws.Cell("K"+frist).Style.Alignment.WrapText = true;
            		ws.Range("K"+frist+":L"+LastMod).Column(1).Merge();
            		ws.Cell("L"+frist).Value = Sudor;
            		ws.Cell("L"+frist).Style.Alignment.WrapText = true;
            		ws.Range("L"+frist+":M"+LastMod).Column(1).Merge();
            		ws.Cell("M"+frist).Value = Cordonator;
            		ws.Cell("M"+frist).Style.Alignment.WrapText = true;
            		ws.Range("M"+frist+":N"+LastMod).Column(1).Merge();
            		ws.Cell("A"+frist).Value = TemporalNr_crt;
            		ws.Cell("A"+frist).Style.Alignment.WrapText = true;
            		ws.Range("A"+frist+":A"+LastMod).Column(1).Merge();
            		maxLenght=0;
             		}
             		Temporal=ws.Cell("B"+a).Value.ToString();
            		TemporalNr_crt=ws.Cell("A"+a).Value.ToString();
            		sef_schimb=ws.Cell("J"+a).Value.ToString();
            		Calitate=ws.Cell("K"+a).Value.ToString();
            		Sudor=ws.Cell("L"+a).Value.ToString();
            		Cordonator=ws.Cell("M"+a).Value.ToString();
             	}
            }
            int ad=78;
            if(maxLenght!=0)
             		{
             		int frist=ad-maxLenght;
            		int LastMod=ad;
            		ws.Cell("B"+frist).Value = Temporal;
            		ws.Cell("B"+frist).Style.Alignment.WrapText = true;
            		ws.Range("B"+frist+":C"+LastMod).Column(1).Merge();
            		ws.Cell("J"+frist).Value = sef_schimb;
            		ws.Cell("J"+frist).Style.Alignment.WrapText = true;
            		ws.Range("J"+frist+":K"+LastMod).Column(1).Merge();
            		ws.Cell("K"+frist).Value = Calitate;
            		ws.Cell("K"+frist).Style.Alignment.WrapText = true;
            		ws.Range("K"+frist+":L"+LastMod).Column(1).Merge();
            		ws.Cell("L"+frist).Value = Sudor;
            		ws.Cell("L"+frist).Style.Alignment.WrapText = true;
            		ws.Range("L"+frist+":M"+LastMod).Column(1).Merge();
            		ws.Cell("M"+frist).Value = Cordonator;
            		ws.Cell("M"+frist).Style.Alignment.WrapText = true;
            		ws.Range("M"+frist+":N"+LastMod).Column(1).Merge();
            		ws.Cell("A"+frist).Value = TemporalNr_crt;
            		ws.Cell("A"+frist).Style.Alignment.WrapText = true;
            		ws.Range("A"+frist+":A"+LastMod).Column(1).Merge();
            		maxLenght=0;
             		}
            
          
              workbook.SaveAs("AddingDataTableAsWorksheet.xlsx");

            
		}
		public void LoadsPITALL(String date1,String date2,String Header,String Footer,String Body,string Query,int rowCount,String id_raport,String Denumire)
		{
			
	
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
		//for(int a=0;a<5;a++)
		//{
			
		//}
		
		
		
		
		
		
		
		
		
		
		
		
DataSet ds = new DataSet("New_DataSet");
DataTable dt = new DataTable(Denumire);
	using (var connP = new MySqlConnection(myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{
     							cmdP.Parameters.AddWithValue("@data1", date1);
     							cmdP.Parameters.AddWithValue("@data2", date2);
     							cmdP.CommandText = Query;
          						 MySqlDataAdapter Da = new MySqlDataAdapter(cmdP);
          						 
                						Da.Fill(dt);
                						
     }
}
	//Adauga nr_crt la date
	int nrcount=1;
     for (int r = 0; r < dt.Rows.Count; r++)
{
     	if("Total"==dt.Rows[r][0].ToString())
     	{
     		nrcount++;
     	}else
     	{
    dt.Rows[r].SetField("nr_crt", nrcount);
     	}
}
		//Add the table to the data set
 var wb = new XLWorkbook();

            // Add a DataTable as a worksheet
            wb.Worksheets.Add(dt);
            wb.SaveAs("AddingDataTableAsWorksheet.xlsx");
             String Temporal="";
             String TemporalNr_crt="";
            var workbook = new XLWorkbook("AddingDataTableAsWorksheet.xlsx");
            var ws = workbook.Worksheets.Worksheet(Denumire);
            int lasta=0;
            int maxLenght=0;
            //int nrCRTindex=1;
            for(int a=0;a<dt.Rows.Count;a++)
            {
            	
            	//merge cnp
            	if(Temporal==dt.Rows[a][1].ToString())
            	{
            		maxLenght++;	lasta=a;
            	}else
            	{
            		if(maxLenght!=0){
            		
            		int frist=2+lasta-maxLenght;
            		int LastMod=lasta+2;
            		ws.Cell("B"+frist).Value = Temporal;
            		ws.Cell("B"+frist).Style.Alignment.WrapText = true;
            		ws.Range("B"+frist+":C"+LastMod).Column(1).Merge();
            		ws.Cell("A"+frist).Value = TemporalNr_crt;
            		ws.Cell("A"+frist).Style.Alignment.WrapText = true;
            		ws.Range("A"+frist+":A"+LastMod).Column(1).Merge();
            		
            		}
            		maxLenght=0;
            		Temporal=dt.Rows[a][1].ToString();
            		TemporalNr_crt=dt.Rows[a][0].ToString();
            	}
            	rowCount++;
            }
            //merge cnp final
            if(maxLenght!=0)
            {
            		int frist=2+lasta-maxLenght;
            		int LastMod=lasta+2;
            		ws.Cell("B"+frist).Value = Temporal;
            		ws.Cell("B"+frist).Style.Alignment.WrapText = true;
            		ws.Range("B"+frist+":C"+LastMod).Column(1).Merge();
            		ws.Cell("A"+frist).Value = TemporalNr_crt;
            		ws.Cell("A"+frist).Style.Alignment.WrapText = true;
            		ws.Range("A"+frist+":A"+LastMod).Column(1).Merge();
            		
            }
            // From a list of strings
            int rowCountLast=rowCount+8;
            int RowCountLastSecound=rowCountLast+4;
            ws.Range("A1:F1").InsertRowsAbove(6);
            ws.Cell("A1").Value = Header;
            ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell("A1").Style.Alignment.WrapText = true;
            ws.Range("A1:F3").Merge();
            ws.Cell("A4").Value = Body;
            ws.Cell("A4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("A4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell("A4").Style.Alignment.WrapText = true;
            ws.Range("A4:F6").Merge();
            ws.Cell("A"+rowCountLast).Value = Footer;
            ws.Cell("A"+rowCountLast).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("A"+rowCountLast).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell("A"+rowCountLast).Style.Alignment.WrapText = true;
            ws.Range("A"+rowCountLast+":F"+RowCountLastSecound).Merge();
            
            
           
            
            
            
            
            
            
            
            
            //merge
            //var workbook = new XLWorkbook("AddingDataTableAsWorksheet.xlsx");
            //var ws = workbook.Worksheets.Worksheet("New_DataTable");
            //ws.Cell("F2").Value = "Merged Column(1) of Range (F2:G8)";
            //ws.Cell("F2").Style.Alignment.WrapText = true;
            //ws.Range("F2:G8").Column(1).Merge();
            workbook.SaveAs("AddingDataTableAsWorksheet.xlsx");
            //SelectSpecificMerge
            var workBook= new XLWorkbook("AddingDataTableAsWorksheet.xlsx");
            var wsMerge = workbook.Worksheets.Worksheet(Denumire);
             
		}
		void Button3Click(object sender, EventArgs e)
		{
			QueryAgain();
		}
		void Button4Click(object sender, EventArgs e)
		{
	 		CodDetaliat f2 = new CodDetaliat();
        	f2.ShowDialog(); // Shows Form2
		}
		public static void DoubleBuffered(DataGridView dgv, bool setting)
{
    Type dgvType = dgv.GetType();
    PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
          BindingFlags.Instance | BindingFlags.NonPublic);
    pi.SetValue(dgv, setting, null);
}
		void Button5Click(object sender, EventArgs e)
		{
			//load secound form
	 		ListaDetali f2 = new ListaDetali();
	 			String date1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
				String date2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
				String Header="";
				String Footer="";
				String Body="";
				string Query ="";
				int b = dataGridView2.CurrentCellAddress.Y;
				String id_raport=dataGridView2.Rows[b].Cells[0].Value.ToString();
				String Flag=dataGridView2.Rows[b].Cells[6].Value.ToString();
				ListaDetali.Flag=Flag;
				String denumire=dataGridView2.Rows[b].Cells[2].Value.ToString();
	 		f2.loadTableInt(myStringCon,LoadDetaliQuery,Header,Footer,Body,Query,id_raport,denumire,date1,date2);
			f2.ShowDialog(); 
        	
		}
		void Button6Click(object sender, EventArgs e)
		{
			this.Close();
		}
		void Button8Click(object sender, EventArgs e)
		{
			dataGridView2.Dispose();
			loadTableInt();
           
		}
		  private void DataGridView2_CellMouseDoubleClick(Object sender, DataGridViewCellMouseEventArgs e)    {

   				ListaDetali f2 = new ListaDetali();
	 			String date1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
				String date2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
				String Header="";
				String Footer="";
				String Body="";
				string Query ="";
				int b = dataGridView2.CurrentCellAddress.Y;
				String id_raport=dataGridView2.Rows[b].Cells[0].Value.ToString();
				String Flag=dataGridView2.Rows[b].Cells[6].Value.ToString();
				ListaDetali.Flag=Flag;
				String denumire=dataGridView2.Rows[b].Cells[2].Value.ToString();
	 		f2.loadTableInt(myStringCon,LoadDetaliQuery,Header,Footer,Body,Query,id_raport,denumire,date1,date2);
			f2.ShowDialog(); 
        	
}
		public static void queryvoid(DataTable dt)
		{
			if(null==ListaDetali.ActiveForm)
			{
				MainForm.fo.queryLoadTableFiller(dt);
			}else
			{
				ListaDetali.fo.queryLoadTableFiller(dt);
			}
		}
		public void queryLoadTableFiller(DataTable dt)
		{
			
            BindingSource dataSource = new BindingSource(dt, null);
            dataGridView2.DataSource=dataSource;
            panel3.Controls.Add(dataGridView2);
		}
		 private void showAllLabel_Click(object sender, EventArgs e)
        {
        	MessageBox.Show("Dot Net Perls is awesome.");
            //DataGridViewAutoFilterColumnHeaderCell.RemoveFilter(dataGridView1);
        }
		 private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
{
    if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null && !string.IsNullOrWhiteSpace(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()))
    {
        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style = new DataGridViewCellStyle { ForeColor = Color.Orange, BackColor = Color.Blue };
    }
    else
    {
        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style = dataGridView1.DefaultCellStyle;
    }
}
		private void dataGridView2_DataBindingComplete(object sender,
            DataGridViewBindingCompleteEventArgs e)
        {
		 	foreach (DataGridViewRow row in dataGridView2.Rows)
            {

                //if (row.Cells["id"].Value.ToString() == "0") //if check ==0
                //{
                    //row.DefaultCellStyle.BackColor = Color.Red;
					//row.DefaultCellStyle.ForeColor = Color.Blue;                    //then change row color to red
                //} 

            }
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
		void MainFormShown(object sender, EventArgs e)
		{
			GetString();
		}
		void MainFormEnter(object sender, EventArgs e)
		{
			GetString();
		}
		void MainFormActivated(object sender, EventArgs e)
		{
			GetString();
			data1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
			data2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
		}
		void TextBox1Leave(object sender, EventArgs e)
		{
			Limit=textBox1.Text.ToString();
		}
		void DateTimePicker1Leave(object sender, EventArgs e)
		{	
				 data1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
				 data2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
		}
		void DateTimePicker2Leave(object sender, EventArgs e)
		{	
			 data1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
			 data2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
	
		}
		public void trimiteMail(string mail,string whoSend)
		{
			int Port=587;
			string convertint;
			using (var connP = new MySqlConnection(myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{
          						//cmdP.Parameters.Add(new MySqlParameter("@an_lucru", id_raport));
          						cmdP.CommandText = " SELECT * FROM trimite_mail where id ='6'";
          						using (MySqlDataReader readerP = cmdP.ExecuteReader())
          						{
               						while (readerP.Read())
               							{
               							//elementele de la header
               							readerP["corp"].ToString();
               							 //string attachmentFilename = "C:\\Users\\bif1\\Desktop\\send.txt";
                        
               	
                        //string attachmentFilename = "C:\\Users\\bif1\\Desktop\\send.txt";
                        SmtpClient smtpClient = new SmtpClient();
                        NetworkCredential basicCredential = new NetworkCredential(readerP["smtp_user"].ToString(), readerP["smtp_pass"].ToString());
                        MailMessage message = new MailMessage();
                        MailAddress fromAddress = new MailAddress( readerP["mail"].ToString());
                        whoSend=readerP["destinatie"].ToString();
                        // setup up the host, increase the timeout to 5 minutes
                        smtpClient.Host =readerP["smtp_host"].ToString();
                        if (readerP["smtp_host"].ToString()== "smtp.gmail.com")
                        {

                            smtpClient.EnableSsl = true;
                        }
                        else
                        {
                            smtpClient.EnableSsl = true;
                            smtpClient.ServicePoint.MaxIdleTime = 1;
                            convertint = readerP["smtp_port"].ToString();
                            Port = Int32.Parse(convertint);
                            if (Port == 486)
                            {
                                smtpClient.Port = Port;
                            }
                            smtpClient.Port = Port;
                        }
                        smtpClient.UseDefaultCredentials = true;
                        smtpClient.Credentials = basicCredential;
                        smtpClient.Timeout = (60 * 5 * 1000);




                        message.From = fromAddress;
                        message.Subject = readerP["subiect"].ToString();
                        message.IsBodyHtml = false;
                        if(TrimiteMailTrue=="2"){
                        	message.Body = readerP["corp"].ToString()+" - "+ErrorMAil;
                        }else{
                        message.Body = readerP["corp"].ToString();
                        }
                        int i = 0;
                        mail = whoSend;
                        string mailtrimits;
                        while (i == 0)
                        {

                            mailtrimits = mail.Substring(0, mail.IndexOf("|"));
                            mail = mail.Substring(mail.IndexOf("|"), (mail.Length - mailtrimits.Length));
                            mail = mail.Substring(1, (mail.Length - 1));
                            message.To.Add(mailtrimits);
                            if (mail.Length == 0)
                            {
                                i++;
                            }
                        };

                      string theDatenowa = DateTime.Now.ToString("yyyy-MM-dd");
                      if(TrimiteMailTrue=="2"){
                      
                      }else{
                                 message.Attachments.Add(new Attachment(destinatie+denumiremail+" "+theDatenowa+".xlsx"));
                      }
                        bool tip;
                       // try
                       // {
                            smtpClient.Send(message);
                            tip = true;
                       // }
                       // catch
                        //{
                        //    tip = false;
                       // }
                       
               }	
          }
     }
}
			
            
		}
		public void trimiteMailToALL()
		{
			foreach (DataGridViewRow row in dataGridView2.Rows)
			{
				String date1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
				String date2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
				String Header="";
				String Footer="";
				String Body="";
				string Query ="";
				int rowCount=0;
				String id_raport=row.Cells[0].Value.ToString();
				String Flag=row.Cells[6].Value.ToString();
				String denumire=row.Cells[2].Value.ToString();
				String tipMail=row.Cells[7].Value.ToString();
				String Mail=row.Cells[9].Value.ToString();
			switch (Flag)
{
   case "P1"://fabrica
			LoadProlyteP(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
   case "L"://laborator
      		LoadsPITALL(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
   default://default
      		LoadDefault(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
}
			
			
			 string theDatenowa = DateTime.Now.ToString("yyyy-MM-dd");
	if(File.Exists(destinatie+denumiremail+" "+theDatenowa+".xlsx")){
	trimiteMail(tipMail,Mail);	
	 }else
	 {
	 	
	 }
	
   //currQty += row.Cells["qty"].Value;
   //More code here
			}
		}
		void Button7Click(object sender, EventArgs e)
		{
			String date1=dateTimePicker1.Value.ToString("yyyy-MM-dd");;
			String date2=dateTimePicker2.Value.ToString("yyyy-MM-dd");;
			String Header="";
			String Footer="";
			String Body="";
			string Query ="";
			int rowCount=0;
	int b = dataGridView2.CurrentCellAddress.Y;
	String id_raport=dataGridView2.Rows[b].Cells[0].Value.ToString();
	String Flag=dataGridView2.Rows[b].Cells[6].Value.ToString();
	String denumire=dataGridView2.Rows[b].Cells[2].Value.ToString();
	String tipMail=dataGridView2.Rows[b].Cells[7].Value.ToString();
	String Mail=dataGridView2.Rows[b].Cells[9].Value.ToString();
	switch (Flag)
{
   case "P1"://fabrica
			LoadProlyteP(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
   case "L"://laborator
      		LoadsPITALL(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
   default://default
      		LoadDefault(date1,date2,Header,Footer,Body,Query,rowCount,id_raport,denumire);
      break;
}
	
	
	//send mail function
	 string theDatenowa = DateTime.Now.ToString("yyyy-MM-dd");
	if(File.Exists(destinatie+denumiremail+" "+theDatenowa+".xlsx")){
	trimiteMail(tipMail,Mail);	
	 }else
	 {
	 	
	 }
	
	
	
		}
		void ComboBox1GiveFeedback(object sender, GiveFeedbackEventArgs e)
		{
	
		}
		
		void Button9Click(object sender, EventArgs e)
		{
			
        
		
		}
		


		
        
		}
}
