namespace WinFormsApp1
{
    partial class EXCEL_TALEP
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
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EXCEL_TALEP));
            groupBox2 = new GroupBox();
            textBox2 = new TextBox();
            label2 = new Label();
            button2 = new Button();
            advancedDataGridView1 = new Zuby.ADGV.AdvancedDataGridView();
            button1 = new Button();
            button3 = new Button();
            groupBox1 = new GroupBox();
            textBox1 = new TextBox();
            toolStrip1 = new ToolStrip();
            toolStripLabel1 = new ToolStripLabel();
            toolStripButton1 = new ToolStripButton();
            checkBox1 = new CheckBox();
            groupBox3 = new GroupBox();
            dateTimePicker2 = new DateTimePicker();
            dateTimePicker1 = new DateTimePicker();
            toolTip1 = new ToolTip(components);
            colorDialog1 = new ColorDialog();
            groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).BeginInit();
            groupBox1.SuspendLayout();
            toolStrip1.SuspendLayout();
            groupBox3.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(textBox2);
            groupBox2.Controls.Add(label2);
            groupBox2.Controls.Add(button2);
            groupBox2.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            groupBox2.Location = new Point(12, 34);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(400, 117);
            groupBox2.TabIndex = 10;
            groupBox2.TabStop = false;
            // 
            // textBox2
            // 
            textBox2.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            textBox2.Location = new Point(146, 14);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(237, 30);
            textBox2.TabIndex = 6;

            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            label2.Location = new Point(11, 14);
            label2.Name = "label2";
            label2.Size = new Size(129, 23);
            label2.TabIndex = 7;
            label2.Text = "Exel Sayfa Adı:";
            // 
            // button2
            // 
            button2.FlatAppearance.BorderColor = Color.Gray;
            button2.FlatAppearance.BorderSize = 0;
            button2.FlatAppearance.MouseDownBackColor = Color.White;
            button2.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 192, 0);
            button2.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            button2.Location = new Point(146, 50);
            button2.Name = "button2";
            button2.Size = new Size(237, 57);
            button2.TabIndex = 5;
            button2.Text = "Excel Tablosunu Aktar";
            button2.TextImageRelation = TextImageRelation.TextAboveImage;
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // advancedDataGridView1
            // 
            advancedDataGridView1.AllowUserToAddRows = false;
            advancedDataGridView1.AllowUserToDeleteRows = false;
            advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            advancedDataGridView1.BackgroundColor = Color.White;
            advancedDataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            advancedDataGridView1.FilterAndSortEnabled = true;
            advancedDataGridView1.FilterStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.Location = new Point(12, 171);
            advancedDataGridView1.Name = "advancedDataGridView1";
            advancedDataGridView1.RightToLeft = RightToLeft.No;
            advancedDataGridView1.RowHeadersWidth = 51;
            advancedDataGridView1.RowTemplate.Height = 29;
            advancedDataGridView1.Size = new Size(1298, 783);
            advancedDataGridView1.SortStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.TabIndex = 11;
            // 
            // button1
            // 
            button1.FlatAppearance.BorderSize = 0;
            button1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            button1.Location = new Point(418, 48);
            button1.Name = "button1";
            button1.Size = new Size(94, 103);
            button1.TabIndex = 12;
            button1.Text = "GÜNCELLE";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button3
            // 
            button3.FlatAppearance.BorderSize = 0;
            button3.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            button3.Location = new Point(1203, 48);
            button3.Name = "button3";
            button3.Size = new Size(107, 103);
            button3.TabIndex = 13;
            button3.Text = "TALEP OLUŞTUR";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(textBox1);
            groupBox1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            groupBox1.Location = new Point(1042, 48);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(155, 55);
            groupBox1.TabIndex = 14;
            groupBox1.TabStop = false;
            groupBox1.Text = "Proje Kodu";
            // 
            // textBox1
            // 
            textBox1.Location = new Point(8, 22);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(141, 27);
            textBox1.TabIndex = 0;
            // 
            // toolStrip1
            // 
            toolStrip1.ImageScalingSize = new Size(20, 20);
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripLabel1, toolStripButton1 });
            toolStrip1.Location = new Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(1322, 27);
            toolStrip1.TabIndex = 15;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel1
            // 
            toolStripLabel1.Image = SİPARİŞ_GİRİŞ.Properties.Resources.icons8_question_mark_64;
            toolStripLabel1.Name = "toolStripLabel1";
            toolStripLabel1.Size = new Size(20, 24);
            // 
            // toolStripButton1
            // 
            toolStripButton1.DisplayStyle = ToolStripItemDisplayStyle.Image;
            toolStripButton1.Image = (Image)resources.GetObject("toolStripButton1.Image");
            toolStripButton1.ImageTransparentColor = Color.Magenta;
            toolStripButton1.Name = "toolStripButton1";
            toolStripButton1.Size = new Size(29, 24);
            toolStripButton1.Text = "toolStripButton1";
            toolStripButton1.Click += toolStripButton1_Click;
            // 
            // checkBox1
            // 
            checkBox1.AutoSize = true;
            checkBox1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            checkBox1.Location = new Point(1042, 108);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(160, 24);
            checkBox1.TabIndex = 16;
            checkBox1.Text = "HEPSİNİ İŞARETLE";
            checkBox1.UseVisualStyleBackColor = true;
            checkBox1.Click += checkBox1_Click;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(dateTimePicker2);
            groupBox3.Controls.Add(dateTimePicker1);
            groupBox3.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            groupBox3.Location = new Point(518, 40);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(250, 111);
            groupBox3.TabIndex = 17;
            groupBox3.TabStop = false;
            groupBox3.Text = "TARİH ARALIĞI";

            // 
            // dateTimePicker2
            // 
            dateTimePicker2.Location = new Point(6, 64);
            dateTimePicker2.Name = "dateTimePicker2";
            dateTimePicker2.Size = new Size(238, 27);
            dateTimePicker2.TabIndex = 1;
            // 
            // dateTimePicker1
            // 
            dateTimePicker1.Location = new Point(6, 28);
            dateTimePicker1.Name = "dateTimePicker1";
            dateTimePicker1.Size = new Size(238, 27);
            dateTimePicker1.TabIndex = 0;
            // 
            // EXCEL_TALEP
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Gray;
            ClientSize = new Size(1322, 971);
            Controls.Add(groupBox3);
            Controls.Add(checkBox1);
            Controls.Add(toolStrip1);
            Controls.Add(groupBox1);
            Controls.Add(button3);
            Controls.Add(button1);
            Controls.Add(advancedDataGridView1);
            Controls.Add(groupBox2);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "EXCEL_TALEP";
            StartPosition = FormStartPosition.CenterScreen;
            Tag = "";
            Text = "EXCEL_TALEP";
            Load += EXCEL_TALEP_Load;
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            groupBox3.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private GroupBox groupBox2;
        private TextBox textBox2;
        private Label label2;
        private Button button2;
        private Zuby.ADGV.AdvancedDataGridView advancedDataGridView1;
        private Button button1;
        private Button button3;
        private GroupBox groupBox1;
        private TextBox textBox1;
        private ToolStrip toolStrip1;
        private ToolStripButton toolStripButton1;
        private CheckBox checkBox1;
        private GroupBox groupBox3;
        private DateTimePicker dateTimePicker2;
        private DateTimePicker dateTimePicker1;
        private ToolTip toolTip1;
        private ToolStripLabel toolStripLabel1;
        private ColorDialog colorDialog1;
    }
}