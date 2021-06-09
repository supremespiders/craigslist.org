
namespace craigslist.org
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.panel3 = new System.Windows.Forms.Panel();
            this.ProgressB = new System.Windows.Forms.ProgressBar();
            this.displayT = new System.Windows.Forms.Label();
            this.metroTabControl1 = new MetroFramework.Controls.MetroTabControl();
            this.metroTabPage1 = new MetroFramework.Controls.MetroTabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.GoogleSheetIdI = new MetroFramework.Controls.MetroTextBox();
            this.startB = new MetroFramework.Controls.MetroButton();
            this.loadInputB = new MetroFramework.Controls.MetroButton();
            this.inputI = new MetroFramework.Controls.MetroTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.threadsI = new System.Windows.Forms.NumericUpDown();
            this.metroTabPage2 = new MetroFramework.Controls.MetroTabPage();
            this.metroPanel2 = new MetroFramework.Controls.MetroPanel();
            this.DebugT = new System.Windows.Forms.RichTextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel3.SuspendLayout();
            this.metroTabControl1.SuspendLayout();
            this.metroTabPage1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.threadsI)).BeginInit();
            this.metroTabPage2.SuspendLayout();
            this.metroPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.SystemColors.HighlightText;
            this.panel3.Controls.Add(this.ProgressB);
            this.panel3.Controls.Add(this.displayT);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(20, 533);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(727, 57);
            this.panel3.TabIndex = 15;
            // 
            // ProgressB
            // 
            this.ProgressB.Location = new System.Drawing.Point(4, 35);
            this.ProgressB.Name = "ProgressB";
            this.ProgressB.Size = new System.Drawing.Size(719, 14);
            this.ProgressB.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.ProgressB.TabIndex = 4;
            // 
            // displayT
            // 
            this.displayT.AutoSize = true;
            this.displayT.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.displayT.ForeColor = System.Drawing.Color.Black;
            this.displayT.Location = new System.Drawing.Point(22, 10);
            this.displayT.Name = "displayT";
            this.displayT.Size = new System.Drawing.Size(91, 16);
            this.displayT.TabIndex = 2;
            this.displayT.Text = "Bot Started";
            // 
            // metroTabControl1
            // 
            this.metroTabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.metroTabControl1.Controls.Add(this.metroTabPage1);
            this.metroTabControl1.Controls.Add(this.metroTabPage2);
            this.metroTabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.metroTabControl1.Location = new System.Drawing.Point(20, 60);
            this.metroTabControl1.Name = "metroTabControl1";
            this.metroTabControl1.SelectedIndex = 0;
            this.metroTabControl1.Size = new System.Drawing.Size(727, 473);
            this.metroTabControl1.Style = MetroFramework.MetroColorStyle.Orange;
            this.metroTabControl1.TabIndex = 16;
            this.metroTabControl1.Theme = MetroFramework.MetroThemeStyle.Light;
            this.metroTabControl1.UseSelectable = true;
            this.metroTabControl1.UseStyleColors = true;
            // 
            // metroTabPage1
            // 
            this.metroTabPage1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.metroTabPage1.Controls.Add(this.panel2);
            this.metroTabPage1.ForeColor = System.Drawing.Color.Black;
            this.metroTabPage1.HorizontalScrollbarBarColor = true;
            this.metroTabPage1.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage1.HorizontalScrollbarSize = 0;
            this.metroTabPage1.Location = new System.Drawing.Point(4, 41);
            this.metroTabPage1.Name = "metroTabPage1";
            this.metroTabPage1.Size = new System.Drawing.Size(719, 428);
            this.metroTabPage1.TabIndex = 0;
            this.metroTabPage1.Text = "Options";
            this.metroTabPage1.VerticalScrollbarBarColor = true;
            this.metroTabPage1.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage1.VerticalScrollbarSize = 0;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.GoogleSheetIdI);
            this.panel2.Controls.Add(this.startB);
            this.panel2.Controls.Add(this.loadInputB);
            this.panel2.Controls.Add(this.inputI);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.threadsI);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(719, 428);
            this.panel2.TabIndex = 14;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(29, 194);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 13);
            this.label3.TabIndex = 28;
            this.label3.Text = "Google sheet id:";
            // 
            // GoogleSheetIdI
            // 
            // 
            // 
            // 
            this.GoogleSheetIdI.CustomButton.Image = null;
            this.GoogleSheetIdI.CustomButton.Location = new System.Drawing.Point(399, 1);
            this.GoogleSheetIdI.CustomButton.Name = "";
            this.GoogleSheetIdI.CustomButton.Size = new System.Drawing.Size(21, 21);
            this.GoogleSheetIdI.CustomButton.Style = MetroFramework.MetroColorStyle.Blue;
            this.GoogleSheetIdI.CustomButton.TabIndex = 1;
            this.GoogleSheetIdI.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light;
            this.GoogleSheetIdI.CustomButton.UseSelectable = true;
            this.GoogleSheetIdI.CustomButton.Visible = false;
            this.GoogleSheetIdI.Lines = new string[0];
            this.GoogleSheetIdI.Location = new System.Drawing.Point(148, 184);
            this.GoogleSheetIdI.MaxLength = 32767;
            this.GoogleSheetIdI.Name = "GoogleSheetIdI";
            this.GoogleSheetIdI.PasswordChar = '\0';
            this.GoogleSheetIdI.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.GoogleSheetIdI.SelectedText = "";
            this.GoogleSheetIdI.SelectionLength = 0;
            this.GoogleSheetIdI.SelectionStart = 0;
            this.GoogleSheetIdI.ShortcutsEnabled = true;
            this.GoogleSheetIdI.Size = new System.Drawing.Size(421, 23);
            this.GoogleSheetIdI.Style = MetroFramework.MetroColorStyle.Orange;
            this.GoogleSheetIdI.TabIndex = 27;
            this.GoogleSheetIdI.UseSelectable = true;
            this.GoogleSheetIdI.WaterMarkColor = System.Drawing.Color.FromArgb(((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.GoogleSheetIdI.WaterMarkFont = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel);
            // 
            // startB
            // 
            this.startB.Location = new System.Drawing.Point(277, 352);
            this.startB.Name = "startB";
            this.startB.Size = new System.Drawing.Size(111, 43);
            this.startB.Style = MetroFramework.MetroColorStyle.Orange;
            this.startB.TabIndex = 23;
            this.startB.Text = "Start";
            this.startB.UseSelectable = true;
            this.startB.UseStyleColors = true;
            this.startB.Click += new System.EventHandler(this.startB_Click_1);
            // 
            // loadInputB
            // 
            this.loadInputB.Location = new System.Drawing.Point(21, 120);
            this.loadInputB.Name = "loadInputB";
            this.loadInputB.Size = new System.Drawing.Size(111, 23);
            this.loadInputB.Style = MetroFramework.MetroColorStyle.Orange;
            this.loadInputB.TabIndex = 22;
            this.loadInputB.Text = "Input File";
            this.loadInputB.UseSelectable = true;
            this.loadInputB.UseStyleColors = true;
            this.loadInputB.Click += new System.EventHandler(this.loadInputB_Click_1);
            // 
            // inputI
            // 
            // 
            // 
            // 
            this.inputI.CustomButton.Image = null;
            this.inputI.CustomButton.Location = new System.Drawing.Point(399, 1);
            this.inputI.CustomButton.Name = "";
            this.inputI.CustomButton.Size = new System.Drawing.Size(21, 21);
            this.inputI.CustomButton.Style = MetroFramework.MetroColorStyle.Blue;
            this.inputI.CustomButton.TabIndex = 1;
            this.inputI.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light;
            this.inputI.CustomButton.UseSelectable = true;
            this.inputI.CustomButton.Visible = false;
            this.inputI.Lines = new string[0];
            this.inputI.Location = new System.Drawing.Point(148, 120);
            this.inputI.MaxLength = 32767;
            this.inputI.Name = "inputI";
            this.inputI.PasswordChar = '\0';
            this.inputI.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.inputI.SelectedText = "";
            this.inputI.SelectionLength = 0;
            this.inputI.SelectionStart = 0;
            this.inputI.ShortcutsEnabled = true;
            this.inputI.Size = new System.Drawing.Size(421, 23);
            this.inputI.Style = MetroFramework.MetroColorStyle.Orange;
            this.inputI.TabIndex = 20;
            this.inputI.UseSelectable = true;
            this.inputI.WaterMarkColor = System.Drawing.Color.FromArgb(((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.inputI.WaterMarkFont = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(29, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(130, 16);
            this.label1.TabIndex = 15;
            this.label1.Text = "Number of threads";
            // 
            // threadsI
            // 
            this.threadsI.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.threadsI.Location = new System.Drawing.Point(165, 38);
            this.threadsI.Maximum = new decimal(new int[] {
            200,
            0,
            0,
            0});
            this.threadsI.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.threadsI.Name = "threadsI";
            this.threadsI.Size = new System.Drawing.Size(58, 21);
            this.threadsI.TabIndex = 6;
            this.threadsI.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // metroTabPage2
            // 
            this.metroTabPage2.Controls.Add(this.metroPanel2);
            this.metroTabPage2.HorizontalScrollbarBarColor = false;
            this.metroTabPage2.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage2.HorizontalScrollbarSize = 0;
            this.metroTabPage2.Location = new System.Drawing.Point(4, 41);
            this.metroTabPage2.Name = "metroTabPage2";
            this.metroTabPage2.Size = new System.Drawing.Size(719, 428);
            this.metroTabPage2.TabIndex = 1;
            this.metroTabPage2.Text = "Logs";
            this.metroTabPage2.VerticalScrollbarBarColor = false;
            this.metroTabPage2.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage2.VerticalScrollbarSize = 0;
            // 
            // metroPanel2
            // 
            this.metroPanel2.Controls.Add(this.DebugT);
            this.metroPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.metroPanel2.HorizontalScrollbarBarColor = true;
            this.metroPanel2.HorizontalScrollbarHighlightOnWheel = false;
            this.metroPanel2.HorizontalScrollbarSize = 10;
            this.metroPanel2.Location = new System.Drawing.Point(0, 0);
            this.metroPanel2.Name = "metroPanel2";
            this.metroPanel2.Size = new System.Drawing.Size(719, 428);
            this.metroPanel2.TabIndex = 2;
            this.metroPanel2.VerticalScrollbarBarColor = true;
            this.metroPanel2.VerticalScrollbarHighlightOnWheel = false;
            this.metroPanel2.VerticalScrollbarSize = 10;
            // 
            // DebugT
            // 
            this.DebugT.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.DebugT.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.DebugT.Cursor = System.Windows.Forms.Cursors.Default;
            this.DebugT.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DebugT.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DebugT.Location = new System.Drawing.Point(0, 0);
            this.DebugT.Margin = new System.Windows.Forms.Padding(4);
            this.DebugT.Name = "DebugT";
            this.DebugT.ReadOnly = true;
            this.DebugT.Size = new System.Drawing.Size(719, 428);
            this.DebugT.TabIndex = 1;
            this.DebugT.Text = "";
            this.DebugT.WordWrap = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.White;
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Image = global::craigslist.org.Properties.Resources.clipart196740;
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(20, 11);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(48, 42);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 17;
            this.pictureBox1.TabStop = false;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 610);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.metroTabControl1);
            this.Controls.Add(this.panel3);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Style = MetroFramework.MetroColorStyle.Orange;
            this.Text = "         Craigslist.org scraper 1.03";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.metroTabControl1.ResumeLayout(false);
            this.metroTabPage1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.threadsI)).EndInit();
            this.metroTabPage2.ResumeLayout(false);
            this.metroPanel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.ProgressBar ProgressB;
        private System.Windows.Forms.Label displayT;
        private MetroFramework.Controls.MetroTabControl metroTabControl1;
        private MetroFramework.Controls.MetroTabPage metroTabPage1;
        private System.Windows.Forms.Panel panel2;
        private MetroFramework.Controls.MetroTabPage metroTabPage2;
        private MetroFramework.Controls.MetroPanel metroPanel2;
        internal System.Windows.Forms.RichTextBox DebugT;
        private System.Windows.Forms.PictureBox pictureBox1;
        private MetroFramework.Controls.MetroButton startB;
        private MetroFramework.Controls.MetroButton loadInputB;
        private MetroFramework.Controls.MetroTextBox inputI;
        private System.Windows.Forms.Label label1;
        internal System.Windows.Forms.NumericUpDown threadsI;
        private System.Windows.Forms.Label label3;
        private MetroFramework.Controls.MetroTextBox GoogleSheetIdI;
    }
}

