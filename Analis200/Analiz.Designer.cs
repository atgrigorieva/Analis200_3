namespace Analis200
{
    partial class Analiz
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
            this.WLNO1 = new System.Windows.Forms.ComboBox();
            this.label0 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // WLNO1
            // 
            this.WLNO1.FormattingEnabled = true;
            this.WLNO1.Location = new System.Drawing.Point(89, 64);
            this.WLNO1.Name = "WLNO1";
            this.WLNO1.Size = new System.Drawing.Size(121, 21);
            this.WLNO1.TabIndex = 0;
            this.WLNO1.SelectedIndexChanged += new System.EventHandler(this.WLNO1_SelectedIndexChanged);
            // 
            // label0
            // 
            this.label0.AutoSize = true;
            this.label0.Location = new System.Drawing.Point(12, 67);
            this.label0.Name = "label0";
            this.label0.Size = new System.Drawing.Size(41, 13);
            this.label0.TabIndex = 1;
            this.label0.Text = "WL No";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(539, 410);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(93, 24);
            this.button1.TabIndex = 2;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Analiz
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(769, 461);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label0);
            this.Controls.Add(this.WLNO1);
            this.Name = "Analiz";
            this.Text = "Analiz";
            this.Load += new System.EventHandler(this.Analiz_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox WLNO1;
        private System.Windows.Forms.Label label0;
        private System.Windows.Forms.Button button1;
    }
}