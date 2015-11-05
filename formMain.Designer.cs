namespace ExcelExposer
{
    partial class formMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formMain));
            this.buttonExpose = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxSourceFile = new System.Windows.Forms.TextBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.checkBoxShowHiddenRows = new System.Windows.Forms.CheckBox();
            this.checkBoxShowHiddenColumns = new System.Windows.Forms.CheckBox();
            this.textBoxStatus = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // buttonExpose
            // 
            this.buttonExpose.Location = new System.Drawing.Point(173, 38);
            this.buttonExpose.Name = "buttonExpose";
            this.buttonExpose.Size = new System.Drawing.Size(293, 41);
            this.buttonExpose.TabIndex = 0;
            this.buttonExpose.Text = "Expose!";
            this.buttonExpose.UseVisualStyleBackColor = true;
            this.buttonExpose.Click += new System.EventHandler(this.buttonExpose_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Source:";
            // 
            // textBoxSourceFile
            // 
            this.textBoxSourceFile.Location = new System.Drawing.Point(62, 12);
            this.textBoxSourceFile.Name = "textBoxSourceFile";
            this.textBoxSourceFile.Size = new System.Drawing.Size(404, 20);
            this.textBoxSourceFile.TabIndex = 2;
            this.textBoxSourceFile.Text = "Double click to select file...";
            this.textBoxSourceFile.DoubleClick += new System.EventHandler(this.textBoxSourceFile_DoubleClick);
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "Excel|*.xls;*.xlsx;*.xlsm";
            // 
            // checkBoxShowHiddenRows
            // 
            this.checkBoxShowHiddenRows.AutoSize = true;
            this.checkBoxShowHiddenRows.Location = new System.Drawing.Point(15, 38);
            this.checkBoxShowHiddenRows.Name = "checkBoxShowHiddenRows";
            this.checkBoxShowHiddenRows.Size = new System.Drawing.Size(139, 17);
            this.checkBoxShowHiddenRows.TabIndex = 6;
            this.checkBoxShowHiddenRows.Text = "Show all Hidden Rows?";
            this.checkBoxShowHiddenRows.UseVisualStyleBackColor = true;
            // 
            // checkBoxShowHiddenColumns
            // 
            this.checkBoxShowHiddenColumns.AutoSize = true;
            this.checkBoxShowHiddenColumns.Location = new System.Drawing.Point(15, 62);
            this.checkBoxShowHiddenColumns.Name = "checkBoxShowHiddenColumns";
            this.checkBoxShowHiddenColumns.Size = new System.Drawing.Size(152, 17);
            this.checkBoxShowHiddenColumns.TabIndex = 7;
            this.checkBoxShowHiddenColumns.Text = "Show all Hidden Columns?";
            this.checkBoxShowHiddenColumns.UseVisualStyleBackColor = true;
            // 
            // textBoxStatus
            // 
            this.textBoxStatus.Location = new System.Drawing.Point(15, 85);
            this.textBoxStatus.Multiline = true;
            this.textBoxStatus.Name = "textBoxStatus";
            this.textBoxStatus.ReadOnly = true;
            this.textBoxStatus.Size = new System.Drawing.Size(451, 102);
            this.textBoxStatus.TabIndex = 8;
            this.textBoxStatus.Text = resources.GetString("textBoxStatus.Text");
            // 
            // formMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(478, 192);
            this.Controls.Add(this.textBoxStatus);
            this.Controls.Add(this.checkBoxShowHiddenColumns);
            this.Controls.Add(this.checkBoxShowHiddenRows);
            this.Controls.Add(this.textBoxSourceFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonExpose);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "formMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel Exposer";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonExpose;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxSourceFile;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.CheckBox checkBoxShowHiddenRows;
        private System.Windows.Forms.CheckBox checkBoxShowHiddenColumns;
        private System.Windows.Forms.TextBox textBoxStatus;
    }
}

