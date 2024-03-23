namespace ExcelGenerator
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.pnlFileHeader = new System.Windows.Forms.Panel();
            this.btnLoad = new System.Windows.Forms.Button();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.empConstantsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.pnlFileHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.empConstantsBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(604, 363);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.pnlFileHeader);
            this.tabPage1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(596, 337);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Excel Generator";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // pnlFileHeader
            // 
            this.pnlFileHeader.Controls.Add(this.btnLoad);
            this.pnlFileHeader.Controls.Add(this.txtFileName);
            this.pnlFileHeader.Controls.Add(this.label13);
            this.pnlFileHeader.Location = new System.Drawing.Point(8, 6);
            this.pnlFileHeader.Name = "pnlFileHeader";
            this.pnlFileHeader.Size = new System.Drawing.Size(563, 44);
            this.pnlFileHeader.TabIndex = 3;
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(485, 10);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(75, 23);
            this.btnLoad.TabIndex = 2;
            this.btnLoad.Text = "Load";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(185, 10);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(294, 23);
            this.txtFileName.TabIndex = 1;
            this.txtFileName.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(26, 10);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(169, 17);
            this.label13.TabIndex = 0;
            this.label13.Text = "Provide master excel: ";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(619, 132);
            this.Controls.Add(this.tabControl1);
            this.Name = "Form1";
            this.Text = "Excel Generator";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.pnlFileHeader.ResumeLayout(false);
            this.pnlFileHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.empConstantsBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.BindingSource empConstantsBindingSource;
        private System.Windows.Forms.Panel pnlFileHeader;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Label label13;

    }
}

