namespace CevaplamaTarziHesaplamasi
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.editBox3 = this.Factory.CreateRibbonEditBox();
            this.editBox4 = this.Factory.CreateRibbonEditBox();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.button3 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Label = "Cevaplama Tarzı Hesaplaması";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.editBox2);
            this.group1.Items.Add(this.editBox3);
            this.group1.Items.Add(this.editBox4);
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Veri Giriş Ekranı";
            this.group1.Name = "group1";
            // 
            // editBox2
            // 
            this.editBox2.Label = "Ölçek Nokta Sayısı";
            this.editBox2.Name = "editBox2";
            this.editBox2.Text = null;
            // 
            // editBox3
            // 
            this.editBox3.Label = "Madde Sayısı";
            this.editBox3.Name = "editBox3";
            this.editBox3.Text = null;
            // 
            // editBox4
            // 
            this.editBox4.Label = "Cevaplayıcı Sayısı";
            this.editBox4.Name = "editBox4";
            this.editBox4.Text = null;
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::CevaplamaTarziHesaplamasi.Properties.Resources.verigirisi;
            this.button1.Label = "Veri Giriş Ekranı Oluştur";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button8);
            this.group2.Items.Add(this.button9);
            this.group2.Label = "Veri Temizleme Ekranı";
            this.group2.Name = "group2";
            // 
            // button8
            // 
            this.button8.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button8.Enabled = false;
            this.button8.Image = global::CevaplamaTarziHesaplamasi.Properties.Resources.clearicon;
            this.button8.Label = "Missing Temizleme";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            // 
            // button9
            // 
            this.button9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button9.Enabled = false;
            this.button9.Image = global::CevaplamaTarziHesaplamasi.Properties.Resources.clearicon;
            this.button9.Label = "Kontrol Sorusu Ekle";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button9_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button2);
            this.group3.Items.Add(this.button4);
            this.group3.Label = "Anket İşlemleri";
            this.group3.Name = "group3";
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Enabled = false;
            this.button2.Image = global::CevaplamaTarziHesaplamasi.Properties.Resources.hesapla;
            this.button2.Label = "Hesapla";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Enabled = false;
            this.button4.Image = global::CevaplamaTarziHesaplamasi.Properties.Resources.guvenilirlik_icon;
            this.button4.Label = "Güvenilirlik Analizi";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.button7);
            this.group4.Items.Add(this.button5);
            this.group4.Label = "Yardım";
            this.group4.Name = "group4";
            // 
            // button7
            // 
            this.button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button7.Image = global::CevaplamaTarziHesaplamasi.Properties.Resources.yardim;
            this.button7.Label = " ";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // button5
            // 
            this.button5.Label = "";
            this.button5.Name = "button5";
            // 
            // group5
            // 
            this.group5.Items.Add(this.button3);
            this.group5.Label = "Hakkımızda";
            this.group5.Name = "group5";
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = global::CevaplamaTarziHesaplamasi.Properties.Resources.hakkimizda;
            this.button3.Label = " ";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
