namespace ADWTG
{
    partial class Main
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 61);
            this.progressBar1.MarqueeAnimationSpeed = 40;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(366, 22);
            this.progressBar1.Step = 25;
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(366, 46);
            this.label1.TabIndex = 1;
            this.label1.Text = "Please wait until ADWTG is done with the creation of the template...";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BackgroundWorker_DoWork);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BackgroundWorker_RunWorkerCompleted);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(390, 100);
            this.ControlBox = false;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Main";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ADWTG - Generate Template...";
            this.Load += new System.EventHandler(this.Main_Load);
            this.Shown += new System.EventHandler(this.Main_Shown);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label1;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
    }
}

