namespace Analis200
{
    partial class SWForm
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
            this.SWText = new System.Windows.Forms.TextBox();
            this.SWButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // SWText
            // 
            this.SWText.Location = new System.Drawing.Point(151, 60);
            this.SWText.Name = "SWText";
            this.SWText.Size = new System.Drawing.Size(100, 20);
            this.SWText.TabIndex = 0;
            // 
            // SWButton
            // 
            this.SWButton.Location = new System.Drawing.Point(308, 57);
            this.SWButton.Name = "SWButton";
            this.SWButton.Size = new System.Drawing.Size(75, 23);
            this.SWButton.TabIndex = 1;
            this.SWButton.Text = "Переход";
            this.SWButton.UseVisualStyleBackColor = true;
            this.SWButton.Click += new System.EventHandler(this.SWButton_Click_1);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Задать длину";
            // 
            // SWForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(658, 282);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.SWButton);
            this.Controls.Add(this.SWText);
            this.Name = "SWForm";
            this.Text = "SWForm";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SWForm_FormClosed);
            this.Load += new System.EventHandler(this.SWForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox SWText;
        private System.Windows.Forms.Button SWButton;
        private System.Windows.Forms.Label label1;
        private System.IO.Ports.SerialPort serialPort1;
        private System.Windows.Forms.Timer timer1;
    }
}