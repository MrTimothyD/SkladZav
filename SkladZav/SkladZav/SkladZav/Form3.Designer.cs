namespace SkladZav
{
    partial class Form3
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form3));
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.skladZavDBDataSet4 = new SkladZav.SkladZavDBDataSet4();
            this.lastMonthBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.lastMonthTableAdapter = new SkladZav.SkladZavDBDataSet4TableAdapters.LastMonthTableAdapter();
            this.halfRecordBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.halfRecordTableAdapter = new SkladZav.SkladZavDBDataSet4TableAdapters.HalfRecordTableAdapter();
            this.button4 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.skladZavDBDataSet4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lastMonthBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.halfRecordBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(80, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(427, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "Добро пожаловать в БД \"Заявка по складу\"";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(12, 115);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(135, 50);
            this.button1.TabIndex = 1;
            this.button1.Text = "Редактирование";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button2.Location = new System.Drawing.Point(153, 115);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(135, 50);
            this.button2.TabIndex = 2;
            this.button2.Text = "Просмотр заявки";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(413, 177);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(157, 15);
            this.label2.TabIndex = 3;
            this.label2.Text = "Автор Тимофей Девятериков";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(220, 80);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(152, 19);
            this.label3.TabIndex = 4;
            this.label3.Text = "Выберите действие";
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button3.Location = new System.Drawing.Point(435, 115);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(135, 50);
            this.button3.TabIndex = 5;
            this.button3.Text = "Выход";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // bindingSource1
            // 
            this.bindingSource1.DataSource = this.skladZavDBDataSet4;
            this.bindingSource1.Position = 0;
            // 
            // skladZavDBDataSet4
            // 
            this.skladZavDBDataSet4.DataSetName = "SkladZavDBDataSet4";
            this.skladZavDBDataSet4.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // lastMonthBindingSource
            // 
            this.lastMonthBindingSource.DataMember = "LastMonth";
            this.lastMonthBindingSource.DataSource = this.bindingSource1;
            // 
            // lastMonthTableAdapter
            // 
            this.lastMonthTableAdapter.ClearBeforeFill = true;
            // 
            // halfRecordBindingSource
            // 
            this.halfRecordBindingSource.DataMember = "HalfRecord";
            this.halfRecordBindingSource.DataSource = this.bindingSource1;
            // 
            // halfRecordTableAdapter
            // 
            this.halfRecordTableAdapter.ClearBeforeFill = true;
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button4.Location = new System.Drawing.Point(294, 115);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(135, 50);
            this.button4.TabIndex = 6;
            this.button4.Text = "Просмотр отчетов";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.MediumSeaGreen;
            this.ClientSize = new System.Drawing.Size(582, 201);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form3";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Заявка по складу";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form3_FormClosing);
            this.Load += new System.EventHandler(this.Form3_Load);
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.skladZavDBDataSet4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lastMonthBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.halfRecordBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.BindingSource bindingSource1;
        private SkladZavDBDataSet4 skladZavDBDataSet4;
        private System.Windows.Forms.BindingSource lastMonthBindingSource;
        private SkladZavDBDataSet4TableAdapters.LastMonthTableAdapter lastMonthTableAdapter;
        private System.Windows.Forms.BindingSource halfRecordBindingSource;
        private SkladZavDBDataSet4TableAdapters.HalfRecordTableAdapter halfRecordTableAdapter;
        private System.Windows.Forms.Button button4;
    }
}