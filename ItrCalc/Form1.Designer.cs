namespace ItrCalc
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
            this.LoadItr = new System.Windows.Forms.Button();
            this.npoi = new System.Windows.Forms.Button();
            this.ms = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menuToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importFilesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.footerPanel = new System.Windows.Forms.Panel();
            this.executionProgress = new System.Windows.Forms.ProgressBar();
            this.footerStatusValue = new System.Windows.Forms.Label();
            this.footerStatus = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.footerPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // LoadItr
            // 
            this.LoadItr.Location = new System.Drawing.Point(591, 215);
            this.LoadItr.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.LoadItr.Name = "LoadItr";
            this.LoadItr.Size = new System.Drawing.Size(115, 41);
            this.LoadItr.TabIndex = 0;
            this.LoadItr.Text = "Load";
            this.LoadItr.UseVisualStyleBackColor = true;
            this.LoadItr.Click += new System.EventHandler(this.button1_Click);
            // 
            // npoi
            // 
            this.npoi.Location = new System.Drawing.Point(591, 330);
            this.npoi.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.npoi.Name = "npoi";
            this.npoi.Size = new System.Drawing.Size(115, 46);
            this.npoi.TabIndex = 1;
            this.npoi.Text = "npoi";
            this.npoi.UseVisualStyleBackColor = true;
            this.npoi.Click += new System.EventHandler(this.npoi_Click);
            // 
            // ms
            // 
            this.ms.Location = new System.Drawing.Point(591, 272);
            this.ms.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ms.Name = "ms";
            this.ms.Size = new System.Drawing.Size(115, 41);
            this.ms.TabIndex = 2;
            this.ms.Text = "Ms";
            this.ms.UseVisualStyleBackColor = true;
            this.ms.Click += new System.EventHandler(this.ms_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(947, 28);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // menuToolStripMenuItem
            // 
            this.menuToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importFilesToolStripMenuItem});
            this.menuToolStripMenuItem.Name = "menuToolStripMenuItem";
            this.menuToolStripMenuItem.Size = new System.Drawing.Size(60, 24);
            this.menuToolStripMenuItem.Text = "Menu";
            // 
            // importFilesToolStripMenuItem
            // 
            this.importFilesToolStripMenuItem.Name = "importFilesToolStripMenuItem";
            this.importFilesToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.importFilesToolStripMenuItem.Text = "Import Files";
            this.importFilesToolStripMenuItem.Click += new System.EventHandler(this.importFilesMenuItem_Click);
            // 
            // footerPanel
            // 
            this.footerPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.footerPanel.Controls.Add(this.executionProgress);
            this.footerPanel.Controls.Add(this.footerStatusValue);
            this.footerPanel.Controls.Add(this.footerStatus);
            this.footerPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.footerPanel.Location = new System.Drawing.Point(0, 488);
            this.footerPanel.Margin = new System.Windows.Forms.Padding(4);
            this.footerPanel.Name = "footerPanel";
            this.footerPanel.Size = new System.Drawing.Size(947, 28);
            this.footerPanel.TabIndex = 4;
            // 
            // executionProgress
            // 
            this.executionProgress.Location = new System.Drawing.Point(756, 5);
            this.executionProgress.Margin = new System.Windows.Forms.Padding(4);
            this.executionProgress.Name = "executionProgress";
            this.executionProgress.Size = new System.Drawing.Size(187, 23);
            this.executionProgress.TabIndex = 2;
            // 
            // footerStatusValue
            // 
            this.footerStatusValue.BackColor = System.Drawing.Color.Transparent;
            this.footerStatusValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.footerStatusValue.Location = new System.Drawing.Point(63, 5);
            this.footerStatusValue.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.footerStatusValue.Name = "footerStatusValue";
            this.footerStatusValue.Size = new System.Drawing.Size(683, 22);
            this.footerStatusValue.TabIndex = 1;
            // 
            // footerStatus
            // 
            this.footerStatus.AutoSize = true;
            this.footerStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.footerStatus.Location = new System.Drawing.Point(0, 6);
            this.footerStatus.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.footerStatus.Name = "footerStatus";
            this.footerStatus.Size = new System.Drawing.Size(59, 17);
            this.footerStatus.TabIndex = 0;
            this.footerStatus.Text = "Status:";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(840, 199);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "DUMMYEXCEL";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(947, 516);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.footerPanel);
            this.Controls.Add(this.ms);
            this.Controls.Add(this.npoi);
            this.Controls.Add(this.LoadItr);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.Text = "Tax Report Generator";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.footerPanel.ResumeLayout(false);
            this.footerPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button LoadItr;
        private System.Windows.Forms.Button npoi;
        private System.Windows.Forms.Button ms;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menuToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importFilesToolStripMenuItem;
        private System.Windows.Forms.Panel footerPanel;
        private System.Windows.Forms.Label footerStatus;
        private System.Windows.Forms.Label footerStatusValue;
        private System.Windows.Forms.ProgressBar executionProgress;
        private System.Windows.Forms.Button button1;
    }
}

