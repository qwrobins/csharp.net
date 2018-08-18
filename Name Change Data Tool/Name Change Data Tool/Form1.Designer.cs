namespace Name_Change_Data_Tool
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.fOpenButton = new System.Windows.Forms.Button();
            this.dirOpenButton = new System.Windows.Forms.Button();
            this.fTextBox = new System.Windows.Forms.TextBox();
            this.dirTextBox = new System.Windows.Forms.TextBox();
            this.parseTextBox = new System.Windows.Forms.TextBox();
            this.convertButton = new System.Windows.Forms.Button();
            this.pg1 = new System.Windows.Forms.ProgressBar();
            this.statusLabel = new System.Windows.Forms.Label();
            this.bgw1 = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(107, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Choose a file to open";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(145, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Choose a directory to save to";
            // 
            // fOpenButton
            // 
            this.fOpenButton.Location = new System.Drawing.Point(571, 36);
            this.fOpenButton.Name = "fOpenButton";
            this.fOpenButton.Size = new System.Drawing.Size(75, 23);
            this.fOpenButton.TabIndex = 2;
            this.fOpenButton.Text = "Browse";
            this.fOpenButton.UseVisualStyleBackColor = true;
            this.fOpenButton.Click += new System.EventHandler(this.fOpenButton_Click);
            // 
            // dirOpenButton
            // 
            this.dirOpenButton.Location = new System.Drawing.Point(571, 77);
            this.dirOpenButton.Name = "dirOpenButton";
            this.dirOpenButton.Size = new System.Drawing.Size(75, 23);
            this.dirOpenButton.TabIndex = 3;
            this.dirOpenButton.Text = "Browse";
            this.dirOpenButton.UseVisualStyleBackColor = true;
            this.dirOpenButton.Click += new System.EventHandler(this.dirOpenButton_Click);
            // 
            // fTextBox
            // 
            this.fTextBox.Location = new System.Drawing.Point(12, 38);
            this.fTextBox.Name = "fTextBox";
            this.fTextBox.Size = new System.Drawing.Size(553, 20);
            this.fTextBox.TabIndex = 4;
            // 
            // dirTextBox
            // 
            this.dirTextBox.Location = new System.Drawing.Point(12, 79);
            this.dirTextBox.Name = "dirTextBox";
            this.dirTextBox.Size = new System.Drawing.Size(553, 20);
            this.dirTextBox.TabIndex = 5;
            // 
            // parseTextBox
            // 
            this.parseTextBox.Location = new System.Drawing.Point(12, 212);
            this.parseTextBox.Multiline = true;
            this.parseTextBox.Name = "parseTextBox";
            this.parseTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.parseTextBox.Size = new System.Drawing.Size(1446, 430);
            this.parseTextBox.TabIndex = 6;
            this.parseTextBox.Visible = false;
            // 
            // convertButton
            // 
            this.convertButton.Location = new System.Drawing.Point(652, 36);
            this.convertButton.Name = "convertButton";
            this.convertButton.Size = new System.Drawing.Size(75, 23);
            this.convertButton.TabIndex = 7;
            this.convertButton.Text = "Convert";
            this.convertButton.UseVisualStyleBackColor = true;
            this.convertButton.Click += new System.EventHandler(this.convertButton_Click);
            // 
            // pg1
            // 
            this.pg1.Location = new System.Drawing.Point(12, 158);
            this.pg1.Name = "pg1";
            this.pg1.Size = new System.Drawing.Size(715, 23);
            this.pg1.TabIndex = 8;
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(9, 121);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(0, 13);
            this.statusLabel.TabIndex = 9;
            // 
            // bgw1
            // 
            this.bgw1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgw1_DoWork);
            this.bgw1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgw1_RunWorkerCompleted);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(769, 199);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.pg1);
            this.Controls.Add(this.convertButton);
            this.Controls.Add(this.parseTextBox);
            this.Controls.Add(this.dirTextBox);
            this.Controls.Add(this.fTextBox);
            this.Controls.Add(this.dirOpenButton);
            this.Controls.Add(this.fOpenButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Name Change Data Tool v1.0.5";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button fOpenButton;
        private System.Windows.Forms.Button dirOpenButton;
        private System.Windows.Forms.TextBox fTextBox;
        private System.Windows.Forms.TextBox dirTextBox;
        private System.Windows.Forms.TextBox parseTextBox;
        private System.Windows.Forms.Button convertButton;
        public System.Windows.Forms.ProgressBar pg1;
        private System.Windows.Forms.Label statusLabel;
        private System.ComponentModel.BackgroundWorker bgw1;
    }
}

