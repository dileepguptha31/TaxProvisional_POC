namespace ItrCalc
{
    partial class ItrCalc
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
            this.btnProcess = new System.Windows.Forms.Button();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtoutput = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbProvisionalMonths = new System.Windows.Forms.ComboBox();
            this.btnloadfiles = new System.Windows.Forms.Button();
            this.lstboxFiles = new System.Windows.Forms.ListBox();
            this.stsstrip = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.tstripstatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.filestatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.processingStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.killOpenExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label4 = new System.Windows.Forms.Label();
            this.stsstrip.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(355, 193);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(146, 35);
            this.btnProcess.TabIndex = 0;
            this.btnProcess.Text = "Process";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Visible = false;
            this.btnProcess.Click += new System.EventHandler(this.Load_Click);
            // 
            // txtPath
            // 
            this.txtPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPath.Location = new System.Drawing.Point(190, 61);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(311, 22);
            this.txtPath.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(54, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 16);
            this.label1.TabIndex = 3;
            this.label1.Text = "Files Path :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(57, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 16);
            this.label2.TabIndex = 4;
            this.label2.Text = "Output Path :";
            // 
            // txtoutput
            // 
            this.txtoutput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtoutput.Location = new System.Drawing.Point(190, 98);
            this.txtoutput.Name = "txtoutput";
            this.txtoutput.Size = new System.Drawing.Size(311, 22);
            this.txtoutput.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(60, 141);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(119, 16);
            this.label3.TabIndex = 6;
            this.label3.Text = "Provisional Month :";
            // 
            // cmbProvisionalMonths
            // 
            this.cmbProvisionalMonths.FormattingEnabled = true;
            this.cmbProvisionalMonths.Items.AddRange(new object[] {
            "Aug",
            "Sep",
            "Oct",
            "Nov",
            "Dec"});
            this.cmbProvisionalMonths.Location = new System.Drawing.Point(190, 141);
            this.cmbProvisionalMonths.Name = "cmbProvisionalMonths";
            this.cmbProvisionalMonths.Size = new System.Drawing.Size(311, 24);
            this.cmbProvisionalMonths.TabIndex = 7;
            // 
            // btnloadfiles
            // 
            this.btnloadfiles.Location = new System.Drawing.Point(190, 193);
            this.btnloadfiles.Name = "btnloadfiles";
            this.btnloadfiles.Size = new System.Drawing.Size(144, 35);
            this.btnloadfiles.TabIndex = 8;
            this.btnloadfiles.Text = "Load Files";
            this.btnloadfiles.UseVisualStyleBackColor = true;
            this.btnloadfiles.Click += new System.EventHandler(this.btnloadfiles_Click);
            // 
            // lstboxFiles
            // 
            this.lstboxFiles.FormattingEnabled = true;
            this.lstboxFiles.ItemHeight = 16;
            this.lstboxFiles.Location = new System.Drawing.Point(190, 260);
            this.lstboxFiles.Name = "lstboxFiles";
            this.lstboxFiles.Size = new System.Drawing.Size(311, 148);
            this.lstboxFiles.TabIndex = 10;
            this.lstboxFiles.Visible = false;
            // 
            // stsstrip
            // 
            this.stsstrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.stsstrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.tstripstatus,
            this.filestatus,
            this.processingStatus});
            this.stsstrip.Location = new System.Drawing.Point(0, 423);
            this.stsstrip.Name = "stsstrip";
            this.stsstrip.Size = new System.Drawing.Size(619, 22);
            this.stsstrip.TabIndex = 11;
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 16);
            // 
            // tstripstatus
            // 
            this.tstripstatus.Name = "tstripstatus";
            this.tstripstatus.Size = new System.Drawing.Size(0, 16);
            // 
            // filestatus
            // 
            this.filestatus.Name = "filestatus";
            this.filestatus.Size = new System.Drawing.Size(0, 16);
            // 
            // processingStatus
            // 
            this.processingStatus.Name = "processingStatus";
            this.processingStatus.Size = new System.Drawing.Size(0, 16);
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(619, 28);
            this.menuStrip1.TabIndex = 12;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.killOpenExcelToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(46, 24);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // killOpenExcelToolStripMenuItem
            // 
            this.killOpenExcelToolStripMenuItem.Name = "killOpenExcelToolStripMenuItem";
            this.killOpenExcelToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.killOpenExcelToolStripMenuItem.Text = "Kill Open Excel";
            this.killOpenExcelToolStripMenuItem.Click += new System.EventHandler(this.killOpenExcelToolStripMenuItem_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(290, 241);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 16);
            this.label4.TabIndex = 13;
            this.label4.Text = "label4";
            this.label4.Visible = false;
            // 
            // ItrCalc
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(619, 445);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.stsstrip);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.lstboxFiles);
            this.Controls.Add(this.btnloadfiles);
            this.Controls.Add(this.cmbProvisionalMonths);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtoutput);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.btnProcess);
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ItrCalc";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ItrCalc";
            this.Load += new System.EventHandler(this.ItrCalc_Load);
            this.stsstrip.ResumeLayout(false);
            this.stsstrip.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtoutput;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbProvisionalMonths;
        private System.Windows.Forms.Button btnloadfiles;
        private System.Windows.Forms.ListBox lstboxFiles;
        private System.Windows.Forms.StatusStrip stsstrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel tstripstatus;
        private System.Windows.Forms.ToolStripStatusLabel filestatus;
        private System.Windows.Forms.ToolStripStatusLabel processingStatus;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem killOpenExcelToolStripMenuItem;
        private System.Windows.Forms.Label label4;
    }
}