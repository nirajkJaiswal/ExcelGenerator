using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            tabControl1.Selecting += new TabControlCancelEventHandler(tabControl1_Selecting);
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs args)
        {
            TabPage current = (sender as TabControl).SelectedTab;

            if (string.IsNullOrEmpty(MyExcel.DB_PATH))
            {
                MessageBox.Show(" Please provide excel file ..", "Error !!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                args.Cancel = true;
            }
        }   

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        protected  override void OnFormClosing(System.Windows.Forms.FormClosingEventArgs e)
        {
            if(!string.IsNullOrEmpty(MyExcel.DB_PATH))
            MyExcel.CloseExcel();
        }     


        private void btnLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog ExcelDialog = new OpenFileDialog();
            ExcelDialog.Filter = "Excel Files (*.xlsx) | *.xlsx";
            ExcelDialog.InitialDirectory = @"C:\";
            ExcelDialog.Title = "Select your master excel";
            if (ExcelDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                MyExcel.DB_PATH = ExcelDialog.FileName;
                txtFileName.Text = ExcelDialog.FileName;
                txtFileName.ReadOnly = true;
                txtFileName.Click -= btnLoad_Click;
                tabControl1.Selecting -= tabControl1_Selecting;
                btnLoad.Enabled = false;
                
                MyExcel.InitializeExcel();
                MyExcel.ReadMyExcel();
                MyExcel.GenerateExcel();

            }
            Form.ActiveForm.Close();
        }

    }
}
