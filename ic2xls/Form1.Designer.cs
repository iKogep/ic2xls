namespace ic2xls
{
    partial class Form1
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.textBox_FileName = new System.Windows.Forms.TextBox();
            this.button_Open = new System.Windows.Forms.Button();
            this.button_Convert = new System.Windows.Forms.Button();
            this.button_Exit = new System.Windows.Forms.Button();
            this.textBox_FileExport = new System.Windows.Forms.TextBox();
            this.button_Save = new System.Windows.Forms.Button();
            this.numericUpDown_Sheet = new System.Windows.Forms.NumericUpDown();
            this.checkBox_DateConvert = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_Sheet)).BeginInit();
            this.SuspendLayout();
            // 
            // textBox_FileName
            // 
            this.textBox_FileName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_FileName.Enabled = false;
            this.textBox_FileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox_FileName.Location = new System.Drawing.Point(12, 12);
            this.textBox_FileName.Name = "textBox_FileName";
            this.textBox_FileName.Size = new System.Drawing.Size(379, 24);
            this.textBox_FileName.TabIndex = 0;
            // 
            // button_Open
            // 
            this.button_Open.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Open.Location = new System.Drawing.Point(397, 12);
            this.button_Open.Name = "button_Open";
            this.button_Open.Size = new System.Drawing.Size(75, 23);
            this.button_Open.TabIndex = 1;
            this.button_Open.Text = "Открыть";
            this.button_Open.UseVisualStyleBackColor = true;
            this.button_Open.Click += new System.EventHandler(this.button_Open_Click);
            // 
            // button_Convert
            // 
            this.button_Convert.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Convert.Enabled = false;
            this.button_Convert.Location = new System.Drawing.Point(12, 72);
            this.button_Convert.Name = "button_Convert";
            this.button_Convert.Size = new System.Drawing.Size(460, 23);
            this.button_Convert.TabIndex = 2;
            this.button_Convert.Text = "Преобразовать";
            this.button_Convert.UseVisualStyleBackColor = true;
            this.button_Convert.Click += new System.EventHandler(this.button_Convert_Click);
            // 
            // button_Exit
            // 
            this.button_Exit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Exit.Location = new System.Drawing.Point(397, 102);
            this.button_Exit.Name = "button_Exit";
            this.button_Exit.Size = new System.Drawing.Size(75, 23);
            this.button_Exit.TabIndex = 3;
            this.button_Exit.Text = "Выход";
            this.button_Exit.UseVisualStyleBackColor = true;
            this.button_Exit.Click += new System.EventHandler(this.button_Exit_Click);
            // 
            // textBox_FileExport
            // 
            this.textBox_FileExport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_FileExport.Enabled = false;
            this.textBox_FileExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox_FileExport.Location = new System.Drawing.Point(12, 42);
            this.textBox_FileExport.Name = "textBox_FileExport";
            this.textBox_FileExport.Size = new System.Drawing.Size(336, 24);
            this.textBox_FileExport.TabIndex = 4;
            // 
            // button_Save
            // 
            this.button_Save.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Save.Enabled = false;
            this.button_Save.Location = new System.Drawing.Point(397, 43);
            this.button_Save.Name = "button_Save";
            this.button_Save.Size = new System.Drawing.Size(75, 23);
            this.button_Save.TabIndex = 5;
            this.button_Save.Text = "Сохранить";
            this.button_Save.UseVisualStyleBackColor = true;
            this.button_Save.Click += new System.EventHandler(this.button_Save_Click);
            // 
            // numericUpDown_Sheet
            // 
            this.numericUpDown_Sheet.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.numericUpDown_Sheet.Enabled = false;
            this.numericUpDown_Sheet.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.numericUpDown_Sheet.Location = new System.Drawing.Point(354, 42);
            this.numericUpDown_Sheet.Maximum = new decimal(new int[] {
            99,
            0,
            0,
            0});
            this.numericUpDown_Sheet.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDown_Sheet.Name = "numericUpDown_Sheet";
            this.numericUpDown_Sheet.Size = new System.Drawing.Size(37, 24);
            this.numericUpDown_Sheet.TabIndex = 6;
            this.numericUpDown_Sheet.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // checkBox_DateConvert
            // 
            this.checkBox_DateConvert.AutoSize = true;
            this.checkBox_DateConvert.Location = new System.Drawing.Point(12, 101);
            this.checkBox_DateConvert.Name = "checkBox_DateConvert";
            this.checkBox_DateConvert.Size = new System.Drawing.Size(273, 17);
            this.checkBox_DateConvert.TabIndex = 7;
            this.checkBox_DateConvert.Text = "автоматически преобразовывать поля с датами";
            this.checkBox_DateConvert.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 137);
            this.Controls.Add(this.checkBox_DateConvert);
            this.Controls.Add(this.numericUpDown_Sheet);
            this.Controls.Add(this.button_Save);
            this.Controls.Add(this.textBox_FileExport);
            this.Controls.Add(this.button_Exit);
            this.Controls.Add(this.button_Convert);
            this.Controls.Add(this.button_Open);
            this.Controls.Add(this.textBox_FileName);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(200, 175);
            this.Name = "Form1";
            this.Text = "ИЦ 2 xls";
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_Sheet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox_FileName;
        private System.Windows.Forms.Button button_Open;
        private System.Windows.Forms.Button button_Convert;
        private System.Windows.Forms.Button button_Exit;
        private System.Windows.Forms.TextBox textBox_FileExport;
        private System.Windows.Forms.Button button_Save;
        private System.Windows.Forms.NumericUpDown numericUpDown_Sheet;
        private System.Windows.Forms.CheckBox checkBox_DateConvert;
    }
}

