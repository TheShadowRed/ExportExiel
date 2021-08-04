/*
 * Created by SharpDevelop.
 * User: TheRedLord
 * Date: 3/29/2017
 * Time: 13:46
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace ExportExiel
{
	partial class MainForm
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Panel panel3;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.DateTimePicker dateTimePicker2;
		private System.Windows.Forms.DateTimePicker dateTimePicker1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button5;
		private System.Windows.Forms.Button button4;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.Button button6;
		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.Button button7;
		private System.Windows.Forms.Button button8;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.RadioButton radioButton3;
		private System.Windows.Forms.RadioButton radioButton2;
		private System.Windows.Forms.RadioButton radioButton1;
		private System.Windows.Forms.Button button9;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing) {
				if (components != null) {
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// This method is required for Windows Forms designer support.
		/// Do not change the method contents inside the source code editor. The Forms designer might
		/// not be able to load this method if it was changed manually.
		/// </summary>
		private void InitializeComponent()
		{
			this.panel1 = new System.Windows.Forms.Panel();
			this.panel3 = new System.Windows.Forms.Panel();
			this.button5 = new System.Windows.Forms.Button();
			this.button8 = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.button4 = new System.Windows.Forms.Button();
			this.panel2 = new System.Windows.Forms.Panel();
			this.button9 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.radioButton3 = new System.Windows.Forms.RadioButton();
			this.radioButton2 = new System.Windows.Forms.RadioButton();
			this.radioButton1 = new System.Windows.Forms.RadioButton();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.button7 = new System.Windows.Forms.Button();
			this.button6 = new System.Windows.Forms.Button();
			this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
			this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
			this.panel1.SuspendLayout();
			this.panel3.SuspendLayout();
			this.panel2.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// panel1
			// 
			this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.panel1.Controls.Add(this.panel3);
			this.panel1.Controls.Add(this.panel2);
			this.panel1.Location = new System.Drawing.Point(12, 12);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(1330, 709);
			this.panel1.TabIndex = 0;
			// 
			// panel3
			// 
			this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.panel3.Controls.Add(this.button5);
			this.panel3.Controls.Add(this.button8);
			this.panel3.Controls.Add(this.button1);
			this.panel3.Controls.Add(this.button3);
			this.panel3.Controls.Add(this.button4);
			this.panel3.Location = new System.Drawing.Point(3, 47);
			this.panel3.Name = "panel3";
			this.panel3.Size = new System.Drawing.Size(1320, 655);
			this.panel3.TabIndex = 1;
			// 
			// button5
			// 
			this.button5.Location = new System.Drawing.Point(119, 522);
			this.button5.Name = "button5";
			this.button5.Size = new System.Drawing.Size(96, 23);
			this.button5.TabIndex = 6;
			this.button5.Text = "Raport Secundar";
			this.button5.UseVisualStyleBackColor = true;
			this.button5.Visible = false;
			this.button5.Click += new System.EventHandler(this.Button5Click);
			// 
			// button8
			// 
			this.button8.Location = new System.Drawing.Point(54, 462);
			this.button8.Name = "button8";
			this.button8.Size = new System.Drawing.Size(75, 23);
			this.button8.TabIndex = 10;
			this.button8.Text = "button8";
			this.button8.UseVisualStyleBackColor = true;
			this.button8.Visible = false;
			this.button8.Click += new System.EventHandler(this.Button8Click);
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(152, 427);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(96, 38);
			this.button1.TabIndex = 2;
			this.button1.Text = "Epoxrt simplu pentru prolyte";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Visible = false;
			this.button1.Click += new System.EventHandler(this.Button1Click);
			// 
			// button3
			// 
			this.button3.Location = new System.Drawing.Point(66, 598);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(94, 23);
			this.button3.TabIndex = 4;
			this.button3.Text = "Sort Button";
			this.button3.UseVisualStyleBackColor = true;
			this.button3.Visible = false;
			this.button3.Click += new System.EventHandler(this.Button3Click);
			// 
			// button4
			// 
			this.button4.Location = new System.Drawing.Point(166, 598);
			this.button4.Name = "button4";
			this.button4.Size = new System.Drawing.Size(94, 23);
			this.button4.TabIndex = 5;
			this.button4.Text = "Cod Detaliat";
			this.button4.UseVisualStyleBackColor = true;
			this.button4.Visible = false;
			this.button4.Click += new System.EventHandler(this.Button4Click);
			// 
			// panel2
			// 
			this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.panel2.Controls.Add(this.button9);
			this.panel2.Controls.Add(this.button2);
			this.panel2.Controls.Add(this.label4);
			this.panel2.Controls.Add(this.textBox1);
			this.panel2.Controls.Add(this.groupBox1);
			this.panel2.Controls.Add(this.label2);
			this.panel2.Controls.Add(this.label1);
			this.panel2.Controls.Add(this.button7);
			this.panel2.Controls.Add(this.button6);
			this.panel2.Controls.Add(this.dateTimePicker2);
			this.panel2.Controls.Add(this.dateTimePicker1);
			this.panel2.Location = new System.Drawing.Point(3, 3);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(1320, 38);
			this.panel2.TabIndex = 0;
			// 
			// button9
			// 
			this.button9.Location = new System.Drawing.Point(1117, 0);
			this.button9.Name = "button9";
			this.button9.Size = new System.Drawing.Size(96, 35);
			this.button9.TabIndex = 16;
			this.button9.Text = "Tastatura";
			this.button9.UseVisualStyleBackColor = true;
			this.button9.Click += new System.EventHandler(this.Button9Click);
			// 
			// button2
			// 
			this.button2.Location = new System.Drawing.Point(512, 0);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(96, 35);
			this.button2.TabIndex = 3;
			this.button2.Text = "Raport";
			this.button2.UseVisualStyleBackColor = true;
			this.button2.Click += new System.EventHandler(this.Button2Click);
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(3, 298);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 19);
			this.label4.TabIndex = 15;
			this.label4.Text = "Limita";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// textBox1
			// 
			this.textBox1.Location = new System.Drawing.Point(3, 320);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(96, 20);
			this.textBox1.TabIndex = 8;
			this.textBox1.Text = "2500";
			this.textBox1.Leave += new System.EventHandler(this.TextBox1Leave);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.radioButton3);
			this.groupBox1.Controls.Add(this.radioButton2);
			this.groupBox1.Controls.Add(this.radioButton1);
			this.groupBox1.Location = new System.Drawing.Point(3, 97);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(96, 198);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Schimb";
			// 
			// radioButton3
			// 
			this.radioButton3.BackColor = System.Drawing.SystemColors.ActiveCaption;
			this.radioButton3.Location = new System.Drawing.Point(6, 139);
			this.radioButton3.Name = "radioButton3";
			this.radioButton3.Size = new System.Drawing.Size(84, 55);
			this.radioButton3.TabIndex = 2;
			this.radioButton3.TabStop = true;
			this.radioButton3.Text = "Schimb 3";
			this.radioButton3.UseVisualStyleBackColor = false;
			// 
			// radioButton2
			// 
			this.radioButton2.BackColor = System.Drawing.SystemColors.ActiveCaption;
			this.radioButton2.Location = new System.Drawing.Point(6, 84);
			this.radioButton2.Name = "radioButton2";
			this.radioButton2.Size = new System.Drawing.Size(84, 49);
			this.radioButton2.TabIndex = 1;
			this.radioButton2.TabStop = true;
			this.radioButton2.Text = "Schimb 2";
			this.radioButton2.UseVisualStyleBackColor = false;
			// 
			// radioButton1
			// 
			this.radioButton1.BackColor = System.Drawing.SystemColors.ActiveCaption;
			this.radioButton1.Location = new System.Drawing.Point(6, 19);
			this.radioButton1.Name = "radioButton1";
			this.radioButton1.Size = new System.Drawing.Size(84, 59);
			this.radioButton1.TabIndex = 0;
			this.radioButton1.TabStop = true;
			this.radioButton1.Text = "Schimb 1";
			this.radioButton1.UseVisualStyleBackColor = false;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(247, 0);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(130, 36);
			this.label2.TabIndex = 13;
			this.label2.Text = "Data Sfarsit";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(3, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(96, 38);
			this.label1.TabIndex = 12;
			this.label1.Text = "Data Start";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button7
			// 
			this.button7.Location = new System.Drawing.Point(614, 0);
			this.button7.Name = "button7";
			this.button7.Size = new System.Drawing.Size(110, 35);
			this.button7.TabIndex = 9;
			this.button7.Text = "Trimite Mail";
			this.button7.UseVisualStyleBackColor = true;
			this.button7.Click += new System.EventHandler(this.Button7Click);
			// 
			// button6
			// 
			this.button6.Location = new System.Drawing.Point(1224, 0);
			this.button6.Name = "button6";
			this.button6.Size = new System.Drawing.Size(94, 35);
			this.button6.TabIndex = 7;
			this.button6.Text = "Inchide";
			this.button6.UseVisualStyleBackColor = true;
			this.button6.Click += new System.EventHandler(this.Button6Click);
			// 
			// dateTimePicker2
			// 
			this.dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.dateTimePicker2.Location = new System.Drawing.Point(383, 3);
			this.dateTimePicker2.Name = "dateTimePicker2";
			this.dateTimePicker2.Size = new System.Drawing.Size(123, 20);
			this.dateTimePicker2.TabIndex = 1;
			this.dateTimePicker2.Leave += new System.EventHandler(this.DateTimePicker2Leave);
			// 
			// dateTimePicker1
			// 
			this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.dateTimePicker1.Location = new System.Drawing.Point(105, 3);
			this.dateTimePicker1.Name = "dateTimePicker1";
			this.dateTimePicker1.Size = new System.Drawing.Size(136, 20);
			this.dateTimePicker1.TabIndex = 0;
			this.dateTimePicker1.Leave += new System.EventHandler(this.DateTimePicker1Leave);
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.AutoSize = true;
			this.ClientSize = new System.Drawing.Size(1354, 733);
			this.Controls.Add(this.panel1);
			this.Name = "MainForm";
			this.Text = "MainForm";
			this.Activated += new System.EventHandler(this.MainFormActivated);
			this.Shown += new System.EventHandler(this.MainFormShown);
			this.Enter += new System.EventHandler(this.MainFormEnter);
			this.panel1.ResumeLayout(false);
			this.panel3.ResumeLayout(false);
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
	}
}
