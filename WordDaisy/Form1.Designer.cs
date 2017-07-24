namespace WordDaisy {
    partial class Form1 {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.openDocButton = new System.Windows.Forms.Button();
            this.curDocTextLabel = new System.Windows.Forms.Label();
            this.curDocNameLabel = new System.Windows.Forms.Label();
            this.processDocButton = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.doneLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openDocButton
            // 
            this.openDocButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.openDocButton.Location = new System.Drawing.Point(38, 150);
            this.openDocButton.Name = "openDocButton";
            this.openDocButton.Size = new System.Drawing.Size(149, 35);
            this.openDocButton.TabIndex = 0;
            this.openDocButton.Text = "Open DOCX";
            this.openDocButton.UseVisualStyleBackColor = true;
            this.openDocButton.Click += new System.EventHandler(this.openDocButton_Click);
            // 
            // curDocTextLabel
            // 
            this.curDocTextLabel.AutoSize = true;
            this.curDocTextLabel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.curDocTextLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.curDocTextLabel.Location = new System.Drawing.Point(34, 33);
            this.curDocTextLabel.Name = "curDocTextLabel";
            this.curDocTextLabel.Size = new System.Drawing.Size(180, 24);
            this.curDocTextLabel.TabIndex = 1;
            this.curDocTextLabel.Text = "Current Document";
            // 
            // curDocNameLabel
            // 
            this.curDocNameLabel.AutoSize = true;
            this.curDocNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.curDocNameLabel.Location = new System.Drawing.Point(34, 57);
            this.curDocNameLabel.Name = "curDocNameLabel";
            this.curDocNameLabel.Size = new System.Drawing.Size(57, 24);
            this.curDocNameLabel.TabIndex = 2;
            this.curDocNameLabel.Text = "None";
            // 
            // processDocButton
            // 
            this.processDocButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.processDocButton.Location = new System.Drawing.Point(38, 224);
            this.processDocButton.Name = "processDocButton";
            this.processDocButton.Size = new System.Drawing.Size(149, 37);
            this.processDocButton.TabIndex = 3;
            this.processDocButton.Text = "Process DOCX";
            this.processDocButton.UseVisualStyleBackColor = true;
            this.processDocButton.Click += new System.EventHandler(this.processDocButton_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // doneLabel
            // 
            this.doneLabel.AutoSize = true;
            this.doneLabel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.doneLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.doneLabel.ForeColor = System.Drawing.Color.Navy;
            this.doneLabel.Location = new System.Drawing.Point(39, 264);
            this.doneLabel.Name = "doneLabel";
            this.doneLabel.Size = new System.Drawing.Size(66, 24);
            this.doneLabel.TabIndex = 4;
            this.doneLabel.Text = "Done!";
            this.doneLabel.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(254, 312);
            this.Controls.Add(this.doneLabel);
            this.Controls.Add(this.processDocButton);
            this.Controls.Add(this.curDocNameLabel);
            this.Controls.Add(this.curDocTextLabel);
            this.Controls.Add(this.openDocButton);
            this.Name = "Form1";
            this.Text = "WordDaisy";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button openDocButton;
        private System.Windows.Forms.Label curDocTextLabel;
        private System.Windows.Forms.Label curDocNameLabel;
        private System.Windows.Forms.Button processDocButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label doneLabel;
    }
}

