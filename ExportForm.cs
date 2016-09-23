using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using RKLib.ExportData;
using System.Configuration;

namespace ExportDataGrid
{
	public class ExportForm : System.Windows.Forms.Form
	{
        private string constr = "";
        private SqlConnection con = null;
        private DataSet ds = null;
		private System.Windows.Forms.DataGrid dgrid;
		private System.Windows.Forms.Button btnExportCSV;
        private System.Windows.Forms.Button btnExportExcel;
		private System.Windows.Forms.Label lblMessage;
        private TextBox textBoxFolder;
        private Label label1;
        private TextBox textBoxSQL;
        private Button button1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ExportForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.dgrid = new System.Windows.Forms.DataGrid();
            this.btnExportCSV = new System.Windows.Forms.Button();
            this.btnExportExcel = new System.Windows.Forms.Button();
            this.lblMessage = new System.Windows.Forms.Label();
            this.textBoxFolder = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxSQL = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgrid)).BeginInit();
            this.SuspendLayout();
            // 
            // dgrid
            // 
            this.dgrid.BackgroundColor = System.Drawing.Color.White;
            this.dgrid.CaptionBackColor = System.Drawing.Color.MidnightBlue;
            this.dgrid.DataMember = "";
            this.dgrid.Font = new System.Drawing.Font("Garamond", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgrid.HeaderForeColor = System.Drawing.Color.Black;
            this.dgrid.Location = new System.Drawing.Point(24, 127);
            this.dgrid.Name = "dgrid";
            this.dgrid.Size = new System.Drawing.Size(730, 308);
            this.dgrid.TabIndex = 0;
            // 
            // btnExportCSV
            // 
            this.btnExportCSV.BackColor = System.Drawing.Color.DarkGray;
            this.btnExportCSV.Font = new System.Drawing.Font("Garamond", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportCSV.Location = new System.Drawing.Point(24, 487);
            this.btnExportCSV.Name = "btnExportCSV";
            this.btnExportCSV.Size = new System.Drawing.Size(112, 23);
            this.btnExportCSV.TabIndex = 1;
            this.btnExportCSV.Text = "Export to CSV";
            this.btnExportCSV.UseVisualStyleBackColor = false;
            this.btnExportCSV.Click += new System.EventHandler(this.btnExportCSV_Click);
            // 
            // btnExportExcel
            // 
            this.btnExportExcel.BackColor = System.Drawing.Color.DarkGray;
            this.btnExportExcel.Font = new System.Drawing.Font("Garamond", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportExcel.Location = new System.Drawing.Point(162, 486);
            this.btnExportExcel.Name = "btnExportExcel";
            this.btnExportExcel.Size = new System.Drawing.Size(128, 24);
            this.btnExportExcel.TabIndex = 3;
            this.btnExportExcel.Text = "Export to Excel";
            this.btnExportExcel.UseVisualStyleBackColor = false;
            this.btnExportExcel.Click += new System.EventHandler(this.btnExportExcel_Click);
            // 
            // lblMessage
            // 
            this.lblMessage.Font = new System.Drawing.Font("Garamond", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMessage.ForeColor = System.Drawing.Color.Red;
            this.lblMessage.Location = new System.Drawing.Point(21, 526);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(544, 40);
            this.lblMessage.TabIndex = 4;
            this.lblMessage.Text = "Error Message";
            // 
            // textBoxFolder
            // 
            this.textBoxFolder.Location = new System.Drawing.Point(123, 446);
            this.textBoxFolder.Name = "textBoxFolder";
            this.textBoxFolder.Size = new System.Drawing.Size(437, 22);
            this.textBoxFolder.TabIndex = 5;
            this.textBoxFolder.Text = "E:/tmp/data/";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Garamond", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(21, 449);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 16);
            this.label1.TabIndex = 6;
            this.label1.Text = "Folder to save:";
            // 
            // textBoxSQL
            // 
            this.textBoxSQL.Location = new System.Drawing.Point(24, 12);
            this.textBoxSQL.Multiline = true;
            this.textBoxSQL.Name = "textBoxSQL";
            this.textBoxSQL.Size = new System.Drawing.Size(638, 109);
            this.textBoxSQL.TabIndex = 7;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.DarkGray;
            this.button1.Font = new System.Drawing.Font("Garamond", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(668, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(86, 25);
            this.button1.TabIndex = 9;
            this.button1.Text = "Load";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // ExportForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(784, 562);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBoxSQL);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxFolder);
            this.Controls.Add(this.lblMessage);
            this.Controls.Add(this.btnExportExcel);
            this.Controls.Add(this.btnExportCSV);
            this.Controls.Add(this.dgrid);
            this.Font = new System.Drawing.Font("Garamond", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "ExportForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ExportForm - WindowsForms C#";
            this.Load += new System.EventHandler(this.ExportForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new ExportForm());
		}


		private void ExportForm_Load(object sender, System.EventArgs e)
		{
            //constr = @"Data Source=HUNGLE\SQLEXPRESS;Initial Catalog=illusions;Persist Security Info=True;User ID=sa;Password=123457";
            constr = ConfigurationSettings.AppSettings["MSSQLConStr"];
            con = new SqlConnection(constr);
		}

        private void button1_Click(object sender, EventArgs e)
        {
            SqlDataAdapter da = new SqlDataAdapter(textBoxSQL.Text, con);
            ds = new DataSet();
            da.Fill(ds, "Table1");
            dgrid.DataSource = ds.Tables["Table1"];
        }

		private void btnExportCSV_Click(object sender, System.EventArgs e)
		{
            lblMessage.Text = "";
            DataTable dt = ds.Tables["Table1"].Copy();
            int nCol = ds.Tables["Table1"].Columns.Count;
            int[] iColumns = new int[nCol];
            for (int i = 1; i < nCol; i++) iColumns[i] = i;
            try
            {
                string file = textBoxFolder.Text + "table1.csv";
                RKLib.ExportData.Export objExport = new RKLib.ExportData.Export("Win");
                objExport.ExportDetails(dt, iColumns, Export.ExportFormat.CSV, file);
                lblMessage.Text = "Successfully exported to " + file;
            }
            catch (Exception Ex)
            {
                lblMessage.Text = Ex.Message;
            }
		}

		private void btnExportExcel_Click(object sender, System.EventArgs e)
		{
            lblMessage.Text = "";
            DataTable dt = ds.Tables["Table1"].Copy();
            try
            {
                string file = textBoxFolder.Text + "table1.xls";
                RKLib.ExportData.Export objExport = new RKLib.ExportData.Export("Win");
                objExport.ExportDetails(dt, Export.ExportFormat.Excel, file);
                lblMessage.Text = "Successfully exported to " + file;
            }
            catch(Exception Ex)
            {
                lblMessage.Text = Ex.Message;
            }
		}
	}
}
