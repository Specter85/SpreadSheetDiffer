namespace SpreadSheetDiffer
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
            this.mBook1File = new System.Windows.Forms.TextBox();
            this.mBook2File = new System.Windows.Forms.TextBox();
            this.mBook1Load = new System.Windows.Forms.Button();
            this.mBook2Load = new System.Windows.Forms.Button();
            this.mSheet1Lbl = new System.Windows.Forms.Label();
            this.mSheet2Lbl = new System.Windows.Forms.Label();
            this.mBook1Sheet = new System.Windows.Forms.ComboBox();
            this.mBook2Sheet = new System.Windows.Forms.ComboBox();
            this.mOutputFileLbl = new System.Windows.Forms.Label();
            this.mOutFileName = new System.Windows.Forms.TextBox();
            this.mCreate = new System.Windows.Forms.Button();
            this.mDiff = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // mBook1File
            // 
            this.mBook1File.Location = new System.Drawing.Point(12, 29);
            this.mBook1File.Name = "mBook1File";
            this.mBook1File.ReadOnly = true;
            this.mBook1File.Size = new System.Drawing.Size(191, 20);
            this.mBook1File.TabIndex = 0;
            // 
            // mBook2File
            // 
            this.mBook2File.Location = new System.Drawing.Point(12, 72);
            this.mBook2File.Name = "mBook2File";
            this.mBook2File.ReadOnly = true;
            this.mBook2File.Size = new System.Drawing.Size(191, 20);
            this.mBook2File.TabIndex = 1;
            // 
            // mBook1Load
            // 
            this.mBook1Load.Location = new System.Drawing.Point(209, 28);
            this.mBook1Load.Name = "mBook1Load";
            this.mBook1Load.Size = new System.Drawing.Size(75, 23);
            this.mBook1Load.TabIndex = 2;
            this.mBook1Load.Text = "Browse";
            this.mBook1Load.UseVisualStyleBackColor = true;
            this.mBook1Load.Click += new System.EventHandler(this.mBook1Load_Click);
            // 
            // mBook2Load
            // 
            this.mBook2Load.Location = new System.Drawing.Point(209, 71);
            this.mBook2Load.Name = "mBook2Load";
            this.mBook2Load.Size = new System.Drawing.Size(75, 23);
            this.mBook2Load.TabIndex = 3;
            this.mBook2Load.Text = "Browse";
            this.mBook2Load.UseVisualStyleBackColor = true;
            this.mBook2Load.Click += new System.EventHandler(this.mBook2Load_Click);
            // 
            // mSheet1Lbl
            // 
            this.mSheet1Lbl.AutoSize = true;
            this.mSheet1Lbl.Location = new System.Drawing.Point(13, 13);
            this.mSheet1Lbl.Name = "mSheet1Lbl";
            this.mSheet1Lbl.Size = new System.Drawing.Size(58, 13);
            this.mSheet1Lbl.TabIndex = 4;
            this.mSheet1Lbl.Text = "Sheet One";
            // 
            // mSheet2Lbl
            // 
            this.mSheet2Lbl.AutoSize = true;
            this.mSheet2Lbl.Location = new System.Drawing.Point(13, 56);
            this.mSheet2Lbl.Name = "mSheet2Lbl";
            this.mSheet2Lbl.Size = new System.Drawing.Size(59, 13);
            this.mSheet2Lbl.TabIndex = 5;
            this.mSheet2Lbl.Text = "Sheet Two";
            // 
            // mBook1Sheet
            // 
            this.mBook1Sheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.mBook1Sheet.Enabled = false;
            this.mBook1Sheet.FormattingEnabled = true;
            this.mBook1Sheet.Location = new System.Drawing.Point(291, 29);
            this.mBook1Sheet.Name = "mBook1Sheet";
            this.mBook1Sheet.Size = new System.Drawing.Size(135, 21);
            this.mBook1Sheet.TabIndex = 6;
            // 
            // mBook2Sheet
            // 
            this.mBook2Sheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.mBook2Sheet.Enabled = false;
            this.mBook2Sheet.FormattingEnabled = true;
            this.mBook2Sheet.Location = new System.Drawing.Point(291, 72);
            this.mBook2Sheet.Name = "mBook2Sheet";
            this.mBook2Sheet.Size = new System.Drawing.Size(135, 21);
            this.mBook2Sheet.TabIndex = 7;
            // 
            // mOutputFileLbl
            // 
            this.mOutputFileLbl.AutoSize = true;
            this.mOutputFileLbl.Location = new System.Drawing.Point(13, 99);
            this.mOutputFileLbl.Name = "mOutputFileLbl";
            this.mOutputFileLbl.Size = new System.Drawing.Size(58, 13);
            this.mOutputFileLbl.TabIndex = 8;
            this.mOutputFileLbl.Text = "Output File";
            // 
            // mOutFileName
            // 
            this.mOutFileName.Location = new System.Drawing.Point(12, 116);
            this.mOutFileName.Name = "mOutFileName";
            this.mOutFileName.ReadOnly = true;
            this.mOutFileName.Size = new System.Drawing.Size(190, 20);
            this.mOutFileName.TabIndex = 9;
            // 
            // mCreate
            // 
            this.mCreate.Location = new System.Drawing.Point(209, 115);
            this.mCreate.Name = "mCreate";
            this.mCreate.Size = new System.Drawing.Size(75, 23);
            this.mCreate.TabIndex = 10;
            this.mCreate.Text = "Browse";
            this.mCreate.UseVisualStyleBackColor = true;
            this.mCreate.Click += new System.EventHandler(this.mCreate_Click);
            // 
            // mDiff
            // 
            this.mDiff.Location = new System.Drawing.Point(351, 115);
            this.mDiff.Name = "mDiff";
            this.mDiff.Size = new System.Drawing.Size(75, 23);
            this.mDiff.TabIndex = 11;
            this.mDiff.Text = "Diff";
            this.mDiff.UseVisualStyleBackColor = true;
            this.mDiff.Click += new System.EventHandler(this.mDiff_Click);
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(434, 142);
            this.Controls.Add(this.mDiff);
            this.Controls.Add(this.mCreate);
            this.Controls.Add(this.mOutFileName);
            this.Controls.Add(this.mOutputFileLbl);
            this.Controls.Add(this.mBook2Sheet);
            this.Controls.Add(this.mBook1Sheet);
            this.Controls.Add(this.mSheet2Lbl);
            this.Controls.Add(this.mSheet1Lbl);
            this.Controls.Add(this.mBook2Load);
            this.Controls.Add(this.mBook1Load);
            this.Controls.Add(this.mBook2File);
            this.Controls.Add(this.mBook1File);
            this.MaximumSize = new System.Drawing.Size(450, 180);
            this.MinimumSize = new System.Drawing.Size(450, 180);
            this.Name = "Form1";
            this.Text = "SpreadSheet Differ";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox mBook1File;
        private System.Windows.Forms.TextBox mBook2File;
        private System.Windows.Forms.Button mBook1Load;
        private System.Windows.Forms.Button mBook2Load;
        private System.Windows.Forms.Label mSheet1Lbl;
        private System.Windows.Forms.Label mSheet2Lbl;
        private System.Windows.Forms.ComboBox mBook1Sheet;
        private System.Windows.Forms.ComboBox mBook2Sheet;
        private System.Windows.Forms.Label mOutputFileLbl;
        private System.Windows.Forms.TextBox mOutFileName;
        private System.Windows.Forms.Button mCreate;
        private System.Windows.Forms.Button mDiff;
    }
}

