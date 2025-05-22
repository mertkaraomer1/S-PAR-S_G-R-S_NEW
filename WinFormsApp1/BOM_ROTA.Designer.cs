namespace WinFormsApp1
{
    partial class BOM_ROTA
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BOM_ROTA));
            advancedDataGridView1 = new Zuby.ADGV.AdvancedDataGridView();
            toolStrip1 = new ToolStrip();
            toolStripLabel1 = new ToolStripLabel();
            toolStripButton1 = new ToolStripButton();
            groupBox2 = new GroupBox();
            button1 = new Button();
            textBox2 = new TextBox();
            label2 = new Label();
            button5 = new Button();
            groupBox5 = new GroupBox();
            button6 = new Button();
            button2 = new Button();
            button3 = new Button();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).BeginInit();
            toolStrip1.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox5.SuspendLayout();
            SuspendLayout();
            // 
            // advancedDataGridView1
            // 
            advancedDataGridView1.AllowUserToAddRows = false;
            advancedDataGridView1.AllowUserToDeleteRows = false;
            advancedDataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            advancedDataGridView1.BackgroundColor = Color.White;
            advancedDataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            advancedDataGridView1.FilterAndSortEnabled = true;
            advancedDataGridView1.FilterStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.Location = new Point(12, 177);
            advancedDataGridView1.Name = "advancedDataGridView1";
            advancedDataGridView1.ReadOnly = true;
            advancedDataGridView1.RightToLeft = RightToLeft.No;
            advancedDataGridView1.RowHeadersWidth = 51;
            advancedDataGridView1.RowTemplate.Height = 29;
            advancedDataGridView1.Size = new Size(1878, 817);
            advancedDataGridView1.SortStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.TabIndex = 3;
            advancedDataGridView1.CellContentClick += advancedDataGridView1_CellContentClick;
            // 
            // toolStrip1
            // 
            toolStrip1.ImageScalingSize = new Size(20, 20);
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripLabel1, toolStripButton1 });
            toolStrip1.Location = new Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(1902, 27);
            toolStrip1.TabIndex = 4;
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
            // groupBox2
            // 
            groupBox2.Controls.Add(button1);
            groupBox2.Controls.Add(textBox2);
            groupBox2.Controls.Add(label2);
            groupBox2.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            groupBox2.Location = new Point(12, 50);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(400, 117);
            groupBox2.TabIndex = 11;
            groupBox2.TabStop = false;
            // 
            // button1
            // 
            button1.Location = new Point(236, 54);
            button1.Name = "button1";
            button1.Size = new Size(147, 57);
            button1.TabIndex = 8;
            button1.Text = "AKTAR";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // textBox2
            // 
            textBox2.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            textBox2.Location = new Point(146, 14);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(237, 30);
            textBox2.TabIndex = 6;
            textBox2.Text = "Sheet1";
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
            // button5
            // 
            button5.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            button5.Location = new Point(137, 26);
            button5.Name = "button5";
            button5.Size = new Size(106, 85);
            button5.TabIndex = 25;
            button5.Text = "MONTAJ SARF MLZM SİL";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // groupBox5
            // 
            groupBox5.Controls.Add(button6);
            groupBox5.Controls.Add(button5);
            groupBox5.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            groupBox5.Location = new Point(436, 49);
            groupBox5.Name = "groupBox5";
            groupBox5.Size = new Size(259, 118);
            groupBox5.TabIndex = 26;
            groupBox5.TabStop = false;
            groupBox5.Text = "BOM ÜSTÜNDE İŞLEMLER";
            // 
            // button6
            // 
            button6.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            button6.Location = new Point(6, 26);
            button6.Name = "button6";
            button6.Size = new Size(94, 84);
            button6.TabIndex = 26;
            button6.Text = "FASON EVET İSE SİL";
            button6.UseVisualStyleBackColor = true;
            button6.Click += button6_Click;
            // 
            // button2
            // 
            button2.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            button2.Location = new Point(797, 76);
            button2.Name = "button2";
            button2.Size = new Size(106, 85);
            button2.TabIndex = 27;
            button2.Text = "RESİMLİ SATINALMA ÇEK\r\n";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click_1;
            // 
            // button3
            // 
            button3.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            button3.Location = new Point(961, 74);
            button3.Name = "button3";
            button3.Size = new Size(106, 85);
            button3.TabIndex = 28;
            button3.Text = "RESİMLİ SATINALMA kAPAT\r\n";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // BOM_ROTA
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Gray;
            ClientSize = new Size(1902, 1033);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(groupBox5);
            Controls.Add(groupBox2);
            Controls.Add(toolStrip1);
            Controls.Add(advancedDataGridView1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "BOM_ROTA";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "BOM_ROTA";
            WindowState = FormWindowState.Maximized;
            Load += BOM_ROTA_Load;
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).EndInit();
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            groupBox5.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Zuby.ADGV.AdvancedDataGridView advancedDataGridView1;
        private ToolStrip toolStrip1;
        private ToolStripButton toolStripButton1;
        private GroupBox groupBox2;
        private TextBox textBox2;
        private Label label2;
        private Button button1;
        private ToolStripLabel toolStripLabel1;
        private Button button5;
        private GroupBox groupBox5;
        private Button button6;
        private Button button2;
        private Button button3;
    }
}