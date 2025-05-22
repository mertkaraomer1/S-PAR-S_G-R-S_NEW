namespace SİPARİŞ_GİRİŞ
{
    partial class VERİLEN_TEKLİF_MALİYET_HESABI
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(VERİLEN_TEKLİF_MALİYET_HESABI));
            toolStrip1 = new ToolStrip();
            toolStripLabel1 = new ToolStripLabel();
            toolStripButton1 = new ToolStripButton();
            toolStripButton2 = new ToolStripButton();
            groupBox1 = new GroupBox();
            button1 = new Button();
            advancedDataGridView1 = new Zuby.ADGV.AdvancedDataGridView();
            button2 = new Button();
            label2 = new Label();
            textBox2 = new TextBox();
            label1 = new Label();
            textBox1 = new TextBox();
            advancedDataGridView2 = new Zuby.ADGV.AdvancedDataGridView();
            toolStrip1.SuspendLayout();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView2).BeginInit();
            SuspendLayout();
            // 
            // toolStrip1
            // 
            toolStrip1.ImageScalingSize = new Size(20, 20);
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripLabel1, toolStripButton1, toolStripButton2 });
            toolStrip1.Location = new Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(1902, 27);
            toolStrip1.TabIndex = 2;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel1
            // 
            toolStripLabel1.Image = Properties.Resources.icons8_question_mark_64;
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
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(button1);
            groupBox1.Controls.Add(advancedDataGridView1);
            groupBox1.Controls.Add(button2);
            groupBox1.Controls.Add(label2);
            groupBox1.Controls.Add(textBox2);
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(textBox1);
            groupBox1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            groupBox1.Location = new Point(12, 30);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(339, 138);
            groupBox1.TabIndex = 3;
            groupBox1.TabStop = false;
            // 
            // button1
            // 
            button1.Location = new Point(222, 76);
            button1.Name = "button1";
            button1.Size = new Size(108, 55);
            button1.TabIndex = 7;
            button1.Text = "Yeni Kod Getir";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // advancedDataGridView1
            // 
            advancedDataGridView1.AllowUserToAddRows = false;
            advancedDataGridView1.AllowUserToDeleteRows = false;
            advancedDataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            advancedDataGridView1.BackgroundColor = Color.White;
            advancedDataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            advancedDataGridView1.FilterAndSortEnabled = true;
            advancedDataGridView1.FilterStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.Location = new Point(-799, -333);
            advancedDataGridView1.Name = "advancedDataGridView1";
            advancedDataGridView1.ReadOnly = true;
            advancedDataGridView1.RightToLeft = RightToLeft.No;
            advancedDataGridView1.RowHeadersWidth = 51;
            advancedDataGridView1.RowTemplate.Height = 29;
            advancedDataGridView1.Size = new Size(312, 0);
            advancedDataGridView1.SortStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.TabIndex = 6;
            // 
            // button2
            // 
            button2.Location = new Point(32, 76);
            button2.Name = "button2";
            button2.Size = new Size(108, 55);
            button2.TabIndex = 5;
            button2.Text = "Eski Kod Getir";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(6, 53);
            label2.Name = "label2";
            label2.Size = new Size(137, 20);
            label2.TabIndex = 4;
            label2.Text = "İşçi Maliyeti (DK) :";
            // 
            // textBox2
            // 
            textBox2.Location = new Point(146, 50);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(70, 27);
            textBox2.TabIndex = 3;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(63, 20);
            label1.Name = "label1";
            label1.Size = new Size(80, 20);
            label1.TabIndex = 1;
            label1.Text = "Parça No :";
            // 
            // textBox1
            // 
            textBox1.Location = new Point(146, 17);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(187, 27);
            textBox1.TabIndex = 0;
            // 
            // advancedDataGridView2
            // 
            advancedDataGridView2.AllowUserToAddRows = false;
            advancedDataGridView2.AllowUserToDeleteRows = false;
            advancedDataGridView2.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            advancedDataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            advancedDataGridView2.BackgroundColor = Color.White;
            advancedDataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            advancedDataGridView2.FilterAndSortEnabled = true;
            advancedDataGridView2.FilterStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView2.Location = new Point(12, 174);
            advancedDataGridView2.Name = "advancedDataGridView2";
            advancedDataGridView2.ReadOnly = true;
            advancedDataGridView2.RightToLeft = RightToLeft.No;
            advancedDataGridView2.RowHeadersWidth = 51;
            advancedDataGridView2.RowTemplate.Height = 29;
            advancedDataGridView2.Size = new Size(1878, 804);
            advancedDataGridView2.SortStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView2.TabIndex = 4;
            // 
            // VERİLEN_TEKLİF_MALİYET_HESABI
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Gray;
            ClientSize = new Size(1902, 1033);
            Controls.Add(advancedDataGridView2);
            Controls.Add(groupBox1);
            Controls.Add(toolStrip1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "VERİLEN_TEKLİF_MALİYET_HESABI";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "VERİLEN_TEKLİF_MALİYET_HESABI";
            WindowState = FormWindowState.Maximized;
            Load += VERİLEN_TEKLİF_MALİYET_HESABI_Load;
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView2).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ToolStrip toolStrip1;
        private ToolStripLabel toolStripLabel1;
        private ToolStripButton toolStripButton1;
        private GroupBox groupBox1;
        private Button button2;
        private Label label2;
        private TextBox textBox2;
        private Label label1;
        private TextBox textBox1;
        private Zuby.ADGV.AdvancedDataGridView advancedDataGridView1;
        private Zuby.ADGV.AdvancedDataGridView advancedDataGridView2;
        private Button button1;
        private ToolStripButton toolStripButton2;
    }
}