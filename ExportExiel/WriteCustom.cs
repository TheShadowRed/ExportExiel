/*
 * Created by SharpDevelop.
 * User: TheRedLord
 * Date: 4/7/2017
 * Time: 11:46
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace ExportExiel
{
	/// <summary>
	/// Description of WriteCustom.
	/// </summary>
	public partial class WriteCustom : Form
	{
		public static String CautaDupaColoana;
		public static String flag;
		public String FristPArt;
		public static Form closewrite;
		public static String data1;
		public static String data2;
		public WriteCustom()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			FristPArt=" and "+CautaDupaColoana+" like ?cautadupa ";
			flag=MainForm.flag;
			closewrite=this;
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		public static void functiequeryAll()
		{
			DataSet ds = new DataSet("New_DataSet");
DataTable dt = new DataTable("New_DataTable");
			using (var connP = new MySqlConnection(MainForm.myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{	
     	//					cmdP.CommandText = "select * from raport";
          						cmdP.Parameters.AddWithValue("@data1", MainForm.data1);
          						cmdP.Parameters.AddWithValue("@data2", MainForm.data2);
          						cmdP.Parameters.AddWithValue("@limit", MainForm.Limit);
          						cmdP.Parameters.AddWithValue("@flag", flag);
     						cmdP.CommandText = MainForm.loadTableQuery;
     						  
          						MySqlDataAdapter Da = new MySqlDataAdapter(cmdP);
          						 
                						Da.Fill(dt);
     						}
						}
			
			MainForm.queryvoid(dt);
			WriteCustom.closewrite.Close();
		}
		void Button1Click(object sender, EventArgs e)
		{
			
DataSet ds = new DataSet("New_DataSet");
DataTable dt = new DataTable("New_DataTable");
			using (var connP = new MySqlConnection(MainForm.myStringCon))
						{
     					connP.Open();
     					using (MySqlCommand cmdP = connP.CreateCommand())
     						{	
     	//					cmdP.CommandText = "select * from raport";
          						cmdP.Parameters.AddWithValue("@flag", flag);
          						cmdP.Parameters.AddWithValue("@data1", MainForm.data1);
          						cmdP.Parameters.AddWithValue("@data2", MainForm.data2);
          						cmdP.Parameters.AddWithValue("@limit", MainForm.Limit);
          						String test="'%"+textBox1.Text.ToString()+"%'";
          						cmdP.Parameters.AddWithValue("@cautadupa",test);
          						cmdP.CommandText = MainForm.loadTableQuery+FristPArt;
          						MySqlDataAdapter Da = new MySqlDataAdapter(cmdP);
          						 
                						Da.Fill(dt);
     						}
						}
			
			MainForm.queryvoid(dt);
			this.Close();
		}
		void Button2Click(object sender, EventArgs e)
		{
			this.Close();
		}
	}
}
