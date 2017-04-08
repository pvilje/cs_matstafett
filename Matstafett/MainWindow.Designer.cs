namespace Matstafett
{
    partial class MainWindow
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
            this.browseButton = new System.Windows.Forms.Button();
            this.fileBox = new System.Windows.Forms.TextBox();
            this.hiddenBoxFileName = new System.Windows.Forms.TextBox();
            this.generateLetters = new System.Windows.Forms.CheckBox();
            this.Start = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // browseButton
            // 
            this.browseButton.Location = new System.Drawing.Point(462, 58);
            this.browseButton.Name = "browseButton";
            this.browseButton.Size = new System.Drawing.Size(75, 23);
            this.browseButton.TabIndex = 0;
            this.browseButton.Text = "Browse";
            this.browseButton.UseVisualStyleBackColor = true;
            this.browseButton.Click += new System.EventHandler(this.browseButton_Click);
            // 
            // fileBox
            // 
            this.fileBox.Location = new System.Drawing.Point(40, 60);
            this.fileBox.Name = "fileBox";
            this.fileBox.ReadOnly = true;
            this.fileBox.Size = new System.Drawing.Size(416, 20);
            this.fileBox.TabIndex = 1;
            // 
            // hiddenBoxFileName
            // 
            this.hiddenBoxFileName.Location = new System.Drawing.Point(40, 13);
            this.hiddenBoxFileName.Name = "hiddenBoxFileName";
            this.hiddenBoxFileName.Size = new System.Drawing.Size(100, 20);
            this.hiddenBoxFileName.TabIndex = 2;
            this.hiddenBoxFileName.Visible = false;
            // 
            // generateLetters
            // 
            this.generateLetters.AutoSize = true;
            this.generateLetters.Checked = true;
            this.generateLetters.CheckState = System.Windows.Forms.CheckState.Checked;
            this.generateLetters.Location = new System.Drawing.Point(40, 106);
            this.generateLetters.Name = "generateLetters";
            this.generateLetters.Size = new System.Drawing.Size(159, 17);
            this.generateLetters.TabIndex = 3;
            this.generateLetters.Text = "Generera brev till deltagarna";
            this.generateLetters.UseVisualStyleBackColor = true;
            // 
            // Start
            // 
            this.Start.Enabled = false;
            this.Start.Location = new System.Drawing.Point(462, 185);
            this.Start.Name = "Start";
            this.Start.Size = new System.Drawing.Size(75, 56);
            this.Start.TabIndex = 4;
            this.Start.Text = "Kör!";
            this.Start.UseVisualStyleBackColor = true;
            this.Start.Click += new System.EventHandler(this.start_Click);
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(618, 281);
            this.Controls.Add(this.Start);
            this.Controls.Add(this.generateLetters);
            this.Controls.Add(this.hiddenBoxFileName);
            this.Controls.Add(this.fileBox);
            this.Controls.Add(this.browseButton);
            this.Name = "MainWindow";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button browseButton;
        private System.Windows.Forms.TextBox fileBox;
        private System.Windows.Forms.TextBox hiddenBoxFileName;
        private System.Windows.Forms.CheckBox generateLetters;
        private System.Windows.Forms.Button Start;
    }
}

