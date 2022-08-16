namespace GETDATA
{
    partial class GETDATA
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GETDATA));
            this.comlist1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.comlist2 = new System.Windows.Forms.ComboBox();
            this.comlist3 = new System.Windows.Forms.ComboBox();
            this.comlist4 = new System.Windows.Forms.ComboBox();
            this.ScanCOM = new System.Windows.Forms.Button();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.Start = new System.Windows.Forms.Button();
            this.Clear = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.Save = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // comlist1
            // 
            this.comlist1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comlist1.FormattingEnabled = true;
            this.comlist1.Location = new System.Drawing.Point(50, 16);
            this.comlist1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comlist1.Name = "comlist1";
            this.comlist1.Size = new System.Drawing.Size(75, 22);
            this.comlist1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 14);
            this.label1.TabIndex = 1;
            this.label1.Text = "Port-1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(143, 19);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 14);
            this.label2.TabIndex = 2;
            this.label2.Text = "Port-2";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 53);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(38, 14);
            this.label3.TabIndex = 3;
            this.label3.Text = "Port-3";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(143, 53);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(38, 14);
            this.label4.TabIndex = 4;
            this.label4.Text = "Port-4";
            // 
            // comlist2
            // 
            this.comlist2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comlist2.FormattingEnabled = true;
            this.comlist2.Location = new System.Drawing.Point(184, 16);
            this.comlist2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comlist2.Name = "comlist2";
            this.comlist2.Size = new System.Drawing.Size(75, 22);
            this.comlist2.TabIndex = 1;
            // 
            // comlist3
            // 
            this.comlist3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comlist3.FormattingEnabled = true;
            this.comlist3.Location = new System.Drawing.Point(50, 50);
            this.comlist3.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comlist3.Name = "comlist3";
            this.comlist3.Size = new System.Drawing.Size(75, 22);
            this.comlist3.TabIndex = 2;
            // 
            // comlist4
            // 
            this.comlist4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comlist4.FormattingEnabled = true;
            this.comlist4.Location = new System.Drawing.Point(184, 50);
            this.comlist4.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comlist4.Name = "comlist4";
            this.comlist4.Size = new System.Drawing.Size(75, 22);
            this.comlist4.TabIndex = 3;
            // 
            // ScanCOM
            // 
            this.ScanCOM.Location = new System.Drawing.Point(293, 8);
            this.ScanCOM.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.ScanCOM.Name = "ScanCOM";
            this.ScanCOM.Size = new System.Drawing.Size(70, 70);
            this.ScanCOM.TabIndex = 5;
            this.ScanCOM.Text = "ScanCOM";
            this.ScanCOM.UseVisualStyleBackColor = true;
            this.ScanCOM.Click += new System.EventHandler(this.ScanCOM_Click);
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.AllowUserToResizeColumns = false;
            this.dataGridView.AllowUserToResizeRows = false;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(9, 85);
            this.dataGridView.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.ReadOnly = true;
            this.dataGridView.RowHeadersVisible = false;
            this.dataGridView.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridView.RowTemplate.Height = 24;
            this.dataGridView.Size = new System.Drawing.Size(806, 125);
            this.dataGridView.TabIndex = 4;
            // 
            // Start
            // 
            this.Start.Location = new System.Drawing.Point(371, 8);
            this.Start.Name = "Start";
            this.Start.Size = new System.Drawing.Size(70, 70);
            this.Start.TabIndex = 9;
            this.Start.Text = "Start";
            this.Start.UseVisualStyleBackColor = true;
            this.Start.Click += new System.EventHandler(this.Start_Click);
            // 
            // Clear
            // 
            this.Clear.Location = new System.Drawing.Point(449, 8);
            this.Clear.Name = "Clear";
            this.Clear.Size = new System.Drawing.Size(70, 70);
            this.Clear.TabIndex = 10;
            this.Clear.Text = "Clear";
            this.Clear.UseVisualStyleBackColor = true;
            this.Clear.Click += new System.EventHandler(this.Clear_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.richTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.richTextBox1.Location = new System.Drawing.Point(9, 217);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(806, 158);
            this.richTextBox1.TabIndex = 13;
            this.richTextBox1.Text = "";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(614, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(201, 66);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 8;
            this.pictureBox1.TabStop = false;
            // 
            // serialPort1
            // 
            this.serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(this.sp_DataReceived);
            // 
            // Save
            // 
            this.Save.Location = new System.Drawing.Point(527, 8);
            this.Save.Name = "Save";
            this.Save.Size = new System.Drawing.Size(70, 70);
            this.Save.TabIndex = 14;
            this.Save.Text = "Save";
            this.Save.UseVisualStyleBackColor = true;
            this.Save.Click += new System.EventHandler(this.Save_to_Excel_Click);
            // 
            // GETDATA
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(824, 383);
            this.Controls.Add(this.Save);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.Clear);
            this.Controls.Add(this.Start);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.comlist4);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.comlist3);
            this.Controls.Add(this.ScanCOM);
            this.Controls.Add(this.comlist2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.comlist1);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "GETDATA";
            this.Text = "GETDATA_A0";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comlist1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comlist4;
        private System.Windows.Forms.ComboBox comlist3;
        private System.Windows.Forms.ComboBox comlist2;
        private System.Windows.Forms.Button ScanCOM;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Button Start;
        private System.Windows.Forms.Button Clear;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.IO.Ports.SerialPort serialPort1;
        private System.Windows.Forms.Button Save;
    }
}

