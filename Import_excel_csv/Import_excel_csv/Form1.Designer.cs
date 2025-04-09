namespace Tools_LayHinh
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btn_import = new System.Windows.Forms.Button();
            this.cb_sheet = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_exportimage = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(2, 63);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1164, 492);
            this.dataGridView1.TabIndex = 0;
            // 
            // btn_import
            // 
            this.btn_import.Location = new System.Drawing.Point(528, 12);
            this.btn_import.Name = "btn_import";
            this.btn_import.Size = new System.Drawing.Size(101, 32);
            this.btn_import.TabIndex = 1;
            this.btn_import.Text = "Imports excel";
            this.btn_import.UseVisualStyleBackColor = true;
            this.btn_import.Click += new System.EventHandler(this.btn_import_Click);
            // 
            // cb_sheet
            // 
            this.cb_sheet.FormattingEnabled = true;
            this.cb_sheet.Location = new System.Drawing.Point(377, 19);
            this.cb_sheet.Name = "cb_sheet";
            this.cb_sheet.Size = new System.Drawing.Size(121, 21);
            this.cb_sheet.TabIndex = 2;
            this.cb_sheet.SelectedIndexChanged += new System.EventHandler(this.cb_sheet_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(307, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Chọn sheet:";
            // 
            // btn_exportimage
            // 
            this.btn_exportimage.Location = new System.Drawing.Point(656, 12);
            this.btn_exportimage.Name = "btn_exportimage";
            this.btn_exportimage.Size = new System.Drawing.Size(101, 32);
            this.btn_exportimage.TabIndex = 4;
            this.btn_exportimage.Text = "Lấy hình - Export";
            this.btn_exportimage.UseVisualStyleBackColor = true;
            this.btn_exportimage.Click += new System.EventHandler(this.btn_exportimage_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1168, 552);
            this.Controls.Add(this.btn_exportimage);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cb_sheet);
            this.Controls.Add(this.btn_import);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Image - Hai.NH";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btn_import;
        private System.Windows.Forms.ComboBox cb_sheet;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_exportimage;
    }
}

