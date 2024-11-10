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
            this.SuspendLayout();
            // 
            // LoadItr
            // 
            this.LoadItr.Location = new System.Drawing.Point(256, 331);
            this.LoadItr.Name = "LoadItr";
            this.LoadItr.Size = new System.Drawing.Size(115, 41);
            this.LoadItr.TabIndex = 0;
            this.LoadItr.Text = "Load";
            this.LoadItr.UseVisualStyleBackColor = true;
            this.LoadItr.Click += new System.EventHandler(this.button1_Click);
            // 
            // npoi
            // 
            this.npoi.Location = new System.Drawing.Point(645, 372);
            this.npoi.Name = "npoi";
            this.npoi.Size = new System.Drawing.Size(75, 23);
            this.npoi.TabIndex = 1;
            this.npoi.Text = "npoi";
            this.npoi.UseVisualStyleBackColor = true;
            this.npoi.Click += new System.EventHandler(this.npoi_Click);
            // 
            // ms
            // 
            this.ms.Location = new System.Drawing.Point(458, 372);
            this.ms.Name = "ms";
            this.ms.Size = new System.Drawing.Size(75, 23);
            this.ms.TabIndex = 2;
            this.ms.Text = "Ms";
            this.ms.UseVisualStyleBackColor = true;
            this.ms.Click += new System.EventHandler(this.ms_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.ms);
            this.Controls.Add(this.npoi);
            this.Controls.Add(this.LoadItr);
            this.Name = "Form1";
            this.Text = "Itr Calc";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button LoadItr;
        private System.Windows.Forms.Button npoi;
        private System.Windows.Forms.Button ms;
    }
}

