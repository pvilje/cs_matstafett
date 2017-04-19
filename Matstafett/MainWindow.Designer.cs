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
            this.generateLetters = new System.Windows.Forms.CheckBox();
            this.Start = new System.Windows.Forms.Button();
            this.log = new System.Windows.Forms.TextBox();
            this.clearLog = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.hjälpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.instruktionerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.kravPåFilenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.openFolder = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
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
            this.browseButton.Click += new System.EventHandler(this.BrowseButton_Click);
            // 
            // fileBox
            // 
            this.fileBox.Location = new System.Drawing.Point(40, 60);
            this.fileBox.Name = "fileBox";
            this.fileBox.ReadOnly = true;
            this.fileBox.Size = new System.Drawing.Size(416, 20);
            this.fileBox.TabIndex = 1;
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
            this.Start.Location = new System.Drawing.Point(462, 106);
            this.Start.Name = "Start";
            this.Start.Size = new System.Drawing.Size(75, 56);
            this.Start.TabIndex = 4;
            this.Start.Text = "Kör!";
            this.Start.UseVisualStyleBackColor = true;
            this.Start.Click += new System.EventHandler(this.Start_Click);
            // 
            // log
            // 
            this.log.Location = new System.Drawing.Point(40, 167);
            this.log.Multiline = true;
            this.log.Name = "log";
            this.log.ReadOnly = true;
            this.log.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.log.Size = new System.Drawing.Size(497, 164);
            this.log.TabIndex = 5;
            // 
            // clearLog
            // 
            this.clearLog.Location = new System.Drawing.Point(544, 307);
            this.clearLog.Name = "clearLog";
            this.clearLog.Size = new System.Drawing.Size(87, 23);
            this.clearLog.TabIndex = 6;
            this.clearLog.Text = "Töm Loggen";
            this.clearLog.UseVisualStyleBackColor = true;
            this.clearLog.Click += new System.EventHandler(this.ClearLog_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.hjälpToolStripMenuItem,
            this.toolStripMenuItem1});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(648, 24);
            this.menuStrip1.TabIndex = 7;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // hjälpToolStripMenuItem
            // 
            this.hjälpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.instruktionerToolStripMenuItem,
            this.kravPåFilenToolStripMenuItem});
            this.hjälpToolStripMenuItem.Name = "hjälpToolStripMenuItem";
            this.hjälpToolStripMenuItem.Size = new System.Drawing.Size(47, 20);
            this.hjälpToolStripMenuItem.Text = "Hjälp";
            // 
            // instruktionerToolStripMenuItem
            // 
            this.instruktionerToolStripMenuItem.Name = "instruktionerToolStripMenuItem";
            this.instruktionerToolStripMenuItem.Size = new System.Drawing.Size(141, 22);
            this.instruktionerToolStripMenuItem.Text = "Instruktioner";
            this.instruktionerToolStripMenuItem.Click += new System.EventHandler(this.InstruktionerToolStripMenuItem_Click);
            // 
            // kravPåFilenToolStripMenuItem
            // 
            this.kravPåFilenToolStripMenuItem.Name = "kravPåFilenToolStripMenuItem";
            this.kravPåFilenToolStripMenuItem.Size = new System.Drawing.Size(141, 22);
            this.kravPåFilenToolStripMenuItem.Text = "Krav på filen";
            this.kravPåFilenToolStripMenuItem.Click += new System.EventHandler(this.KravPåFilenToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(24, 20);
            this.toolStripMenuItem1.Text = "?";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.ToolStripMenuItem1_Click);
            // 
            // openFolder
            // 
            this.openFolder.Location = new System.Drawing.Point(544, 278);
            this.openFolder.Name = "openFolder";
            this.openFolder.Size = new System.Drawing.Size(87, 23);
            this.openFolder.TabIndex = 8;
            this.openFolder.Text = "Öppna Mapp";
            this.openFolder.UseVisualStyleBackColor = true;
            this.openFolder.Visible = false;
            this.openFolder.Click += new System.EventHandler(this.OpenFolder_Click);
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(648, 344);
            this.Controls.Add(this.openFolder);
            this.Controls.Add(this.clearLog);
            this.Controls.Add(this.log);
            this.Controls.Add(this.Start);
            this.Controls.Add(this.generateLetters);
            this.Controls.Add(this.fileBox);
            this.Controls.Add(this.browseButton);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainWindow";
            this.Text = "MatStafett";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button browseButton;
        private System.Windows.Forms.TextBox fileBox;
        private System.Windows.Forms.CheckBox generateLetters;
        private System.Windows.Forms.Button Start;
        private System.Windows.Forms.TextBox log;
        private System.Windows.Forms.Button clearLog;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem hjälpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem instruktionerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem kravPåFilenToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.Button openFolder;
    }
}

