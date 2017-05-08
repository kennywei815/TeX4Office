namespace TeX4Office_WindowsFormsApplication
{
    partial class EditorForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EditorForm));
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.ribbon1 = new System.Windows.Forms.Ribbon();
            this.ribbonSeparator1 = new System.Windows.Forms.RibbonSeparator();
            this.ribbonOrbMenuItem1 = new System.Windows.Forms.RibbonOrbMenuItem();
            this.ribbonTextBox2 = new System.Windows.Forms.RibbonTextBox();
            this.ribbonTextBox1 = new System.Windows.Forms.RibbonTextBox();
            this.ribbonTab1 = new System.Windows.Forms.RibbonTab();
            this.outputPanel = new System.Windows.Forms.RibbonPanel();
            this.generateButton1 = new System.Windows.Forms.RibbonButton();
            this.engineComboBox1 = new System.Windows.Forms.RibbonComboBox();
            this.engineButton_pdflatex = new System.Windows.Forms.RibbonButton();
            this.engineButton_lualatex = new System.Windows.Forms.RibbonButton();
            this.engineButton_xelatex = new System.Windows.Forms.RibbonButton();
            this.dpiComboBox1 = new System.Windows.Forms.RibbonComboBox();
            this.dpiButton600 = new System.Windows.Forms.RibbonButton();
            this.dpiButton1200 = new System.Windows.Forms.RibbonButton();
            this.dpiButton2400 = new System.Windows.Forms.RibbonButton();
            this.dpiButton4800 = new System.Windows.Forms.RibbonButton();
            this.helpPanel = new System.Windows.Forms.RibbonPanel();
            this.helpButton1 = new System.Windows.Forms.RibbonButton();
            this.defaultsPanel = new System.Windows.Forms.RibbonPanel();
            this.loadButton1 = new System.Windows.Forms.RibbonButton();
            this.saveButton1 = new System.Windows.Forms.RibbonButton();
            this.generateButton = new System.Windows.Forms.RibbonButton();
            this.SuspendLayout();
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(-1, 158);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(895, 492);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = "";
            // 
            // ribbon1
            // 
            this.ribbon1.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.ribbon1.Location = new System.Drawing.Point(0, 0);
            this.ribbon1.Minimized = false;
            this.ribbon1.Name = "ribbon1";
            // 
            // 
            // 
            this.ribbon1.OrbDropDown.BorderRoundness = 8;
            this.ribbon1.OrbDropDown.Location = new System.Drawing.Point(0, 0);
            this.ribbon1.OrbDropDown.MenuItems.Add(this.ribbonSeparator1);
            this.ribbon1.OrbDropDown.MenuItems.Add(this.ribbonOrbMenuItem1);
            this.ribbon1.OrbDropDown.Name = "";
            this.ribbon1.OrbDropDown.Size = new System.Drawing.Size(527, 119);
            this.ribbon1.OrbDropDown.TabIndex = 0;
            this.ribbon1.OrbImage = null;
            this.ribbon1.OrbStyle = System.Windows.Forms.RibbonOrbStyle.Office_2013;
            this.ribbon1.OrbText = "檔案";
            this.ribbon1.OrbVisible = false;
            // 
            // 
            // 
            this.ribbon1.QuickAcessToolbar.DropDownButtonItems.Add(this.ribbonTextBox2);
            this.ribbon1.QuickAcessToolbar.Items.Add(this.ribbonTextBox1);
            this.ribbon1.RibbonTabFont = new System.Drawing.Font("Trebuchet MS", 9F);
            this.ribbon1.Size = new System.Drawing.Size(894, 160);
            this.ribbon1.TabIndex = 6;
            this.ribbon1.Tabs.Add(this.ribbonTab1);
            this.ribbon1.TabsMargin = new System.Windows.Forms.Padding(12, 26, 20, 0);
            this.ribbon1.Text = "ribbon1";
            this.ribbon1.ThemeColor = System.Windows.Forms.RibbonTheme.Blue;
            this.ribbon1.Click += new System.EventHandler(this.ribbon_main_botton_click);
            // 
            // ribbonOrbMenuItem1
            // 
            this.ribbonOrbMenuItem1.DropDownArrowDirection = System.Windows.Forms.RibbonArrowDirection.Left;
            this.ribbonOrbMenuItem1.Image = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItem1.Image")));
            this.ribbonOrbMenuItem1.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItem1.SmallImage")));
            this.ribbonOrbMenuItem1.Text = "ribbonOrbMenuItem1";
            // 
            // ribbonTextBox2
            // 
            this.ribbonTextBox2.TextBoxText = "";
            // 
            // ribbonTextBox1
            // 
            this.ribbonTextBox1.MaxSizeMode = System.Windows.Forms.RibbonElementSizeMode.Compact;
            this.ribbonTextBox1.Text = "";
            this.ribbonTextBox1.TextBoxText = "";
            this.ribbonTextBox1.TextBoxWidth = 300;
            this.ribbonTextBox1.Value = "";
            // 
            // ribbonTab1
            // 
            this.ribbonTab1.Panels.Add(this.outputPanel);
            this.ribbonTab1.Panels.Add(this.helpPanel);
            this.ribbonTab1.Panels.Add(this.defaultsPanel);
            this.ribbonTab1.Text = "常用";
            // 
            // outputPanel
            // 
            this.outputPanel.Items.Add(this.generateButton1);
            this.outputPanel.Items.Add(this.engineComboBox1);
            this.outputPanel.Items.Add(this.dpiComboBox1);
            this.outputPanel.Text = "輸出";
            // 
            // generateButton1
            // 
            this.generateButton1.Image = global::TeX4Office_WindowsFormsApplication.Properties.Resources.synaptic_64px_1174960_easyicon_net;
            this.generateButton1.SmallImage = ((System.Drawing.Image)(resources.GetObject("generateButton1.SmallImage")));
            this.generateButton1.Text = "輸出並關閉";
            this.generateButton1.Click += new System.EventHandler(this.output_botton_click);
            // 
            // engineComboBox1
            // 
            this.engineComboBox1.AllowTextEdit = false;
            this.engineComboBox1.DropDownItems.Add(this.engineButton_pdflatex);
            this.engineComboBox1.DropDownItems.Add(this.engineButton_lualatex);
            this.engineComboBox1.DropDownItems.Add(this.engineButton_xelatex);
            this.engineComboBox1.LabelWidth = 34;
            this.engineComboBox1.Text = "引擎:";
            this.engineComboBox1.TextBoxText = "PDFLaTeX";
            this.engineComboBox1.TextBoxWidth = 110;
            this.engineComboBox1.Value = "pdflatex";
            // 
            // engineButton_pdflatex
            // 
            this.engineButton_pdflatex.DrawIconsBar = false;
            this.engineButton_pdflatex.Image = ((System.Drawing.Image)(resources.GetObject("engineButton_pdflatex.Image")));
            this.engineButton_pdflatex.SmallImage = ((System.Drawing.Image)(resources.GetObject("engineButton_pdflatex.SmallImage")));
            this.engineButton_pdflatex.Tag = "PDFLaTeX";
            this.engineButton_pdflatex.Text = "PDFLaTeX";
            this.engineButton_pdflatex.Value = "PDFLaTeX";
            // 
            // engineButton_lualatex
            // 
            this.engineButton_lualatex.DrawIconsBar = false;
            this.engineButton_lualatex.Image = ((System.Drawing.Image)(resources.GetObject("engineButton_lualatex.Image")));
            this.engineButton_lualatex.SmallImage = ((System.Drawing.Image)(resources.GetObject("engineButton_lualatex.SmallImage")));
            this.engineButton_lualatex.Tag = "LuaLaTeX";
            this.engineButton_lualatex.Text = "LuaLaTeX";
            this.engineButton_lualatex.Value = "LuaLaTeX";
            // 
            // engineButton_xelatex
            // 
            this.engineButton_xelatex.DrawIconsBar = false;
            this.engineButton_xelatex.Image = ((System.Drawing.Image)(resources.GetObject("engineButton_xelatex.Image")));
            this.engineButton_xelatex.SmallImage = ((System.Drawing.Image)(resources.GetObject("engineButton_xelatex.SmallImage")));
            this.engineButton_xelatex.Tag = "XeLaTeX";
            this.engineButton_xelatex.Text = "XeLaTeX";
            this.engineButton_xelatex.Value = "XeLaTeX";
            // 
            // dpiComboBox1
            // 
            this.dpiComboBox1.AllowTextEdit = false;
            this.dpiComboBox1.DropDownItems.Add(this.dpiButton600);
            this.dpiComboBox1.DropDownItems.Add(this.dpiButton1200);
            this.dpiComboBox1.DropDownItems.Add(this.dpiButton2400);
            this.dpiComboBox1.DropDownItems.Add(this.dpiButton4800);
            this.dpiComboBox1.LabelWidth = 34;
            this.dpiComboBox1.Text = "DPI:";
            this.dpiComboBox1.TextBoxText = "600";
            this.dpiComboBox1.TextBoxWidth = 110;
            this.dpiComboBox1.Value = "1200";
            // 
            // dpiButton600
            // 
            this.dpiButton600.Image = ((System.Drawing.Image)(resources.GetObject("dpiButton600.Image")));
            this.dpiButton600.SmallImage = ((System.Drawing.Image)(resources.GetObject("dpiButton600.SmallImage")));
            this.dpiButton600.Tag = "600";
            this.dpiButton600.Text = "600";
            this.dpiButton600.Value = "600";
            // 
            // dpiButton1200
            // 
            this.dpiButton1200.Image = ((System.Drawing.Image)(resources.GetObject("dpiButton1200.Image")));
            this.dpiButton1200.SmallImage = ((System.Drawing.Image)(resources.GetObject("dpiButton1200.SmallImage")));
            this.dpiButton1200.Tag = "1200";
            this.dpiButton1200.Text = "1200";
            this.dpiButton1200.Value = "1200";
            // 
            // dpiButton2400
            // 
            this.dpiButton2400.Image = ((System.Drawing.Image)(resources.GetObject("dpiButton2400.Image")));
            this.dpiButton2400.SmallImage = ((System.Drawing.Image)(resources.GetObject("dpiButton2400.SmallImage")));
            this.dpiButton2400.Tag = "2400";
            this.dpiButton2400.Text = "2400";
            this.dpiButton2400.Value = "2400";
            // 
            // dpiButton4800
            // 
            this.dpiButton4800.Image = ((System.Drawing.Image)(resources.GetObject("dpiButton4800.Image")));
            this.dpiButton4800.SmallImage = ((System.Drawing.Image)(resources.GetObject("dpiButton4800.SmallImage")));
            this.dpiButton4800.Tag = "4800";
            this.dpiButton4800.Text = "4800";
            this.dpiButton4800.Value = "4800";
            // 
            // helpPanel
            // 
            this.helpPanel.Items.Add(this.helpButton1);
            this.helpPanel.Text = "說明";
            // 
            // helpButton1
            // 
            this.helpButton1.Image = global::TeX4Office_WindowsFormsApplication.Properties.Resources.help_info_64px_1174848_easyicon_net;
            this.helpButton1.SmallImage = ((System.Drawing.Image)(resources.GetObject("helpButton1.SmallImage")));
            this.helpButton1.Text = "說明";
            this.helpButton1.Click += new System.EventHandler(this.help_botton_click);
            // 
            // defaultsPanel
            // 
            this.defaultsPanel.Items.Add(this.loadButton1);
            this.defaultsPanel.Items.Add(this.saveButton1);
            this.defaultsPanel.Text = "範本";
            // 
            // loadButton1
            // 
            this.loadButton1.Image = global::TeX4Office_WindowsFormsApplication.Properties.Resources.softwarecenter_64px_1174877_easyicon_net;
            this.loadButton1.SmallImage = ((System.Drawing.Image)(resources.GetObject("loadButton1.SmallImage")));
            this.loadButton1.Text = "載入";
            this.loadButton1.Click += new System.EventHandler(this.load_botton_click);
            // 
            // saveButton1
            // 
            this.saveButton1.Image = global::TeX4Office_WindowsFormsApplication.Properties.Resources.software_properties_64px_1174892_easyicon_net;
            this.saveButton1.SmallImage = ((System.Drawing.Image)(resources.GetObject("saveButton1.SmallImage")));
            this.saveButton1.Text = "儲存";
            this.saveButton1.Click += new System.EventHandler(this.save_botton_click);
            // 
            // generateButton
            // 
            this.generateButton.Image = ((System.Drawing.Image)(resources.GetObject("generateButton.Image")));
            this.generateButton.SmallImage = ((System.Drawing.Image)(resources.GetObject("generateButton.SmallImage")));
            this.generateButton.Text = "輸出";
            // 
            // EditorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(894, 648);
            this.Controls.Add(this.ribbon1);
            this.Controls.Add(this.richTextBox1);
            this.Name = "EditorForm";
            this.Text = "TeX4Office Editor";
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Ribbon ribbon1;
        private System.Windows.Forms.RibbonTextBox ribbonTextBox2;
        private System.Windows.Forms.RibbonTextBox ribbonTextBox1;
        private System.Windows.Forms.RibbonSeparator ribbonSeparator1;
        private System.Windows.Forms.RibbonOrbMenuItem ribbonOrbMenuItem1;
        private System.Windows.Forms.RibbonTab ribbonTab1;
        private System.Windows.Forms.RibbonPanel outputPanel;
        private System.Windows.Forms.RibbonButton generateButton;
        private System.Windows.Forms.RibbonButton generateButton1;
        private System.Windows.Forms.RibbonPanel helpPanel;
        private System.Windows.Forms.RibbonButton helpButton1;
        private System.Windows.Forms.RibbonPanel defaultsPanel;
        private System.Windows.Forms.RibbonButton loadButton1;
        private System.Windows.Forms.RibbonButton saveButton1;
        private System.Windows.Forms.RibbonComboBox engineComboBox1;
        private System.Windows.Forms.RibbonButton engineButton_pdflatex;
        private System.Windows.Forms.RibbonComboBox dpiComboBox1;
        private System.Windows.Forms.RibbonButton engineButton_lualatex;
        private System.Windows.Forms.RibbonButton engineButton_xelatex;
        private System.Windows.Forms.RibbonButton dpiButton600;
        private System.Windows.Forms.RibbonButton dpiButton1200;
        private System.Windows.Forms.RibbonButton dpiButton2400;
        private System.Windows.Forms.RibbonButton dpiButton4800;
    }
}

