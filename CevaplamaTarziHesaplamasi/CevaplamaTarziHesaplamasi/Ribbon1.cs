using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Collections;
using System.Windows.Forms;
using System.Drawing;

namespace CevaplamaTarziHesaplamasi
{
    public partial class Ribbon1
    {
        MissingTemizlemeFormu MTemizlemeFormu;
        KontrolSorusuEklemeFormu KSoruFormu;
        //Global Değişkenler.
        public int Anket_Tipi, Madde_Sayisi, Katilimci_Sayisi, ProgressBarMaxValue = 0;
        double EKOCT = 0, EKOCT_TOPLAM = 0, EKCT = 0, EKCT_TOPLAM = 0, ECT = 0, KCT = 0, KCT_TOPLAM = 0, KOCT = 0, KOCT_TOPLAM = 0, NKCT = 0, NKCT_TOPLAM = 0, ONCT = 0, ONCT_TOPLAM = 0, CVPSZCT = 0, CVPSZCT_TOPLAM = 0, ARADCT = 0, ARADCT_TOPLAM = 0;
        double[,] Sonuclar;
        public ArrayList Secilecek_Liste = new ArrayList();
        public ArrayList Secilen_Liste = new ArrayList();
        ArrayList harfler = new ArrayList();
        ArrayList SilinecekIndisler, Bos_Satirlar;
        string IslemSayfasi = "";
        public Boolean FormKontrol = false;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            HarfleriTanımla();
        }

        public void HarfleriTanımla()
        {
            harfler.Add("A"); harfler.Add("B"); harfler.Add("C"); harfler.Add("D"); harfler.Add("E"); harfler.Add("F"); harfler.Add("G"); harfler.Add("H"); harfler.Add("I"); harfler.Add("J"); harfler.Add("K"); harfler.Add("L"); harfler.Add("M"); harfler.Add("N"); harfler.Add("O"); harfler.Add("P"); harfler.Add("Q"); harfler.Add("R");
            harfler.Add("S"); harfler.Add("T"); harfler.Add("U"); harfler.Add("V"); harfler.Add("W"); harfler.Add("X"); harfler.Add("Y"); harfler.Add("Z"); harfler.Add("AA"); harfler.Add("AB"); harfler.Add("AC"); harfler.Add("AD"); harfler.Add("AE"); harfler.Add("AF"); harfler.Add("AG"); harfler.Add("AH"); harfler.Add("AI"); harfler.Add("AJ");
            harfler.Add("AK"); harfler.Add("AL"); harfler.Add("AM"); harfler.Add("AN"); harfler.Add("AO"); harfler.Add("AP"); harfler.Add("AQ"); harfler.Add("AR"); harfler.Add("AS"); harfler.Add("AT"); harfler.Add("AU"); harfler.Add("AV"); harfler.Add("AW"); harfler.Add("AX"); harfler.Add("AY"); harfler.Add("AZ"); harfler.Add("BA"); harfler.Add("BB");
            harfler.Add("BC"); harfler.Add("BD"); harfler.Add("BE"); harfler.Add("BF"); harfler.Add("BG"); harfler.Add("BH"); harfler.Add("BI"); harfler.Add("BJ"); harfler.Add("BK"); harfler.Add("BL"); harfler.Add("BM"); harfler.Add("BN"); harfler.Add("BO"); harfler.Add("BP"); harfler.Add("BQ"); harfler.Add("BR"); harfler.Add("BS"); harfler.Add("BT");
        }

        public string Hucrelerden_Okuma_Metodu(int satir, int sutun)
        {//Bu metodun amacı içerisine aldığı satır ve sütun numaralı hucreyi bulup onun değerini göndermek.
            //Eğer gönderilecek değer boş ise Sıfır değerini geri gönderir.
            Excel.Worksheet currentW = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range cell = currentW.Cells.get_Item(satir, sutun) as Excel.Range;
            string siradaki = cell.Text;
            if (siradaki.Trim() == "" || siradaki.Trim() == "0") return "0";
            return siradaki.Trim();
        }

        public void Hucreye_Yazdirma_Metodu(int satir, int sutun, string mesaj)
        {//Bu metodun amacı içerisine aldığı satir ve sütun numaralı hücreyi bulup o hücreye değer yazdırır.
            Excel.Application xlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook xlWorkbook = xlApp.ActiveWorkbook;
            Excel.Worksheet xlWorksheet = xlWorkbook.ActiveSheet;
            xlWorksheet.Cells[satir, sutun] = mesaj;
            xlWorksheet.Columns.AutoFit();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //ilk önce ölçek nokta sayısı madde sayısı ve cevaplayıcı sayısı doğru girilip girilmediği kontrol edilir.
            if (editBox2.Text != "" && Convert.ToInt32(editBox2.Text) >= 1)
            {
                if (editBox3.Text != "" && Convert.ToInt32(editBox3.Text) >= 1)
                {
                    if (editBox4.Text != "" && Convert.ToInt32(editBox4.Text) >= 1)
                    {
                        Globals.ThisAddIn.Application.Cells.Clear();//sayfayı temizler
                        Anket_Tipi = Convert.ToInt32(editBox2.Text);

                        Madde_Sayisi = Convert.ToInt32(editBox3.Text);

                        Secilecek_Liste.Clear();
                        Secilen_Liste.Clear();
                        Hucreye_Yazdirma_Metodu(1, 1, "ID");
                        for (int l = 1; l <= Madde_Sayisi; l++)//veri temizleme işleminde kullanılması için maddeler güncellenir.
                            Secilecek_Liste.Add(l);
                        int i;
                        for (i = 1; i <= Madde_Sayisi; i++)//Veri giriş ekranı oluşturulur.
                        {
                            Hucreye_Yazdirma_Metodu(1, (i + 1), "Madde" + i.ToString());
                        }
                        Hucreye_Yazdirma_Metodu(1, (i + 2), "ECT");
                        Hucreye_Yazdirma_Metodu(1, (i + 3), "EKCT");
                        Hucreye_Yazdirma_Metodu(1, (i + 4), "EKOCT");
                        Hucreye_Yazdirma_Metodu(1, (i + 5), "KCT");
                        Hucreye_Yazdirma_Metodu(1, (i + 6), "KOCT");
                        Hucreye_Yazdirma_Metodu(1, (i + 7), "NKCT");
                        Hucreye_Yazdirma_Metodu(1, (i + 8), "ONCT");
                        Hucreye_Yazdirma_Metodu(1, (i + 9), "CVPSZCT");
                        Hucreye_Yazdirma_Metodu(1, (i + 10), "ARADCT");

                        Katilimci_Sayisi = Convert.ToInt32(editBox4.Text);

                        Hucreye_Yazdirma_Metodu(Katilimci_Sayisi + 2, Madde_Sayisi + 2, "Ortalamalar");

                        button2.Enabled = true;
                        button8.Enabled = true;
                        button9.Enabled = true;
                        button2.Enabled = true;
                        button4.Enabled = true;
                        Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
                        IslemSayfasi = activeWorksheet.Name;//hesaplama işleminde kullanılması için işlem sayfası tutulur.
                    }
                    else MessageBox.Show("Cevaplayıcı Sayısı Boş Bırakılamaz Ve 0 ' dan Büyük Olması Gerekir.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else MessageBox.Show("Madde Sayısı Boş Bırakılamaz Ve 0 ' dan Büyük Olması Gerekir.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else MessageBox.Show("Ölçek Tipi Boş Bırakılamaz Ve 0 ' dan Büyük Olması Gerekir.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            string AktifSayfa = activeWorksheet.Name;
            if (IslemSayfasi == AktifSayfa)//Doğru sayfada olup olmadığı kontrol edilir.
            {//veriler sıfırlanır.
                bool Hata_Yakalama = false;
                EKOCT_TOPLAM = 0;
                EKCT_TOPLAM = 0;
                KCT_TOPLAM = 0;
                KOCT_TOPLAM = 0;
                NKCT_TOPLAM = 0;
                ONCT_TOPLAM = 0;
                CVPSZCT_TOPLAM = 0;
                ARADCT_TOPLAM = 0;
                //bütün soruları işaretlemeyen cevaplayıcılar tespit edilir.
                Bos_Satirlar = new ArrayList();
                int Bos_Satir_Sayisi = 0;
                for (int i = 1; i <= Katilimci_Sayisi; i++)
                {//Bütün boş satırların sayısını ve indis numaralarını bulur.
                    int sayac = 0;
                    for (int j = 1; j <= Madde_Sayisi; j++)
                    {
                        string siradaki = Hucrelerden_Okuma_Metodu((i + 1), (j + 1));
                        if (siradaki == "" || siradaki == " " || siradaki == "0") sayac++;
                    }
                    if (sayac == Madde_Sayisi)
                    {
                        Bos_Satir_Sayisi++;
                        Bos_Satirlar.Add(i + 1);
                    }
                }

                Excel.Worksheet wsSil = Globals.ThisAddIn.Application.ActiveSheet;
                Excel.Range row;
                int[] rows = new int[Bos_Satirlar.Count];
                int SilinenSayac = 0;
                foreach (var item in Bos_Satirlar)
                {//Bütün soruları işaretlemeyen cevaplayıcılar silinir.
                    row = (Excel.Range)wsSil.Rows[Convert.ToInt32(item.ToString()) - SilinenSayac];
                    row.Delete();
                    SilinenSayac++;
                }

                Katilimci_Sayisi = Katilimci_Sayisi - Bos_Satir_Sayisi;
                Sonuclar = new double[Katilimci_Sayisi, 9];

                for (int i = 1; i <= Katilimci_Sayisi; i++)
                {
                    Hata_Yakalama = false;
                    ArrayList Katilimci_Cevaplari = new ArrayList();//Her katılımcı için liste oluşturulur.
                    for (int j = 1; j <= Madde_Sayisi; j++)
                    {
                        string siradaki = Hucrelerden_Okuma_Metodu((i + 1), (j + 1));
                        if (siradaki == "" || siradaki == " ") Katilimci_Cevaplari.Add(siradaki);
                        else
                        {
                            try
                            {//Doğru aralıkta veri girişi yapılmış ise katılımcının cevapları bir listeye eklenir.
                                //yanlış veri girişi yapılmış ise program hata mesajı verir.
                                if (Convert.ToInt32(siradaki) > Anket_Tipi || Convert.ToInt32(siradaki) < 0)
                                { i = Madde_Sayisi + 1; Hata_Yakalama = true; MessageBox.Show("Hatalı Veri Girişi Yapılmıştır.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error); break; }
                                else Katilimci_Cevaplari.Add(siradaki);
                            }
                            catch (Exception ex) { i = Madde_Sayisi + 1; Hata_Yakalama = true; MessageBox.Show("Hatalı Veri Girişi Yapılmıştır.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error); break; }
                        }
                    }

                    //formüllerin hesaplanması için ilgili metod çağırılır.
                    EKOCT_Metod(Katilimci_Cevaplari, i);

                    EKCT_Metod(Katilimci_Cevaplari, i);

                    ECT = EKOCT + EKCT;//Ekstrem Cevaplama Tarzı.

                    Sonuclar[(i - 1), 0] = ECT;

                    if (Anket_Tipi % 2 == 1)
                    {//Orta Noktası Bulunan Anket Türleri İçin.

                        KCT_Metod(Katilimci_Cevaplari, i);
                        KCT_TOPLAM += KCT;


                        KOCT_Metod(Katilimci_Cevaplari, i);
                        KOCT_TOPLAM += KOCT;


                        ONCT_Metod(Katilimci_Cevaplari, i);
                        ONCT_TOPLAM += ONCT;
                    }
                    NKCT = KCT - KOCT;
                    NKCT_TOPLAM += NKCT;

                    Sonuclar[(i - 1), 5] = NKCT;


                    CVPSZCT_Metod(Katilimci_Cevaplari, i);
                    CVPSZCT_TOPLAM += CVPSZCT;

                    ARADCT = 1 - ECT;
                    ARADCT_TOPLAM += ARADCT;

                    Sonuclar[(i - 1), 8] = ARADCT;

                    EKOCT_TOPLAM += EKOCT;
                    EKCT_TOPLAM += EKCT;
                }
                if (Hata_Yakalama == false)//bir hata ile karşılaşılmadı ise devam eder.
                {
                    Excel.Worksheet ws1 = Globals.ThisAddIn.Application.ActiveSheet;
                    Excel.Range range = (Excel.Range)ws1.Range[ws1.Cells[2, Madde_Sayisi + 3], ws1.Cells[Katilimci_Sayisi + 1, Madde_Sayisi + 11]];
                    range.Value2 = Sonuclar;
                    //her formül için ortalama değerleri bulunur.
                    EKOCT_TOPLAM = EKOCT_TOPLAM / Katilimci_Sayisi;

                    Hucreye_Yazdirma_Metodu(Katilimci_Sayisi + 2, Madde_Sayisi + 5, EKOCT_TOPLAM.ToString("0.#####"));

                    EKCT_TOPLAM = EKCT_TOPLAM / Katilimci_Sayisi;

                    Hucreye_Yazdirma_Metodu(Katilimci_Sayisi + 2, Madde_Sayisi + 4, EKCT_TOPLAM.ToString("0.#####"));


                    Hucreye_Yazdirma_Metodu(Katilimci_Sayisi + 2, Madde_Sayisi + 3, (EKCT_TOPLAM + EKOCT_TOPLAM).ToString("0.#####"));

                    KCT_TOPLAM = KCT_TOPLAM / Katilimci_Sayisi;

                    Hucreye_Yazdirma_Metodu(Katilimci_Sayisi + 2, Madde_Sayisi + 6, KCT_TOPLAM.ToString("0.#####"));

                    KOCT_TOPLAM = KOCT_TOPLAM / Katilimci_Sayisi;

                    Hucreye_Yazdirma_Metodu(Katilimci_Sayisi + 2, Madde_Sayisi + 7, KOCT_TOPLAM.ToString("0.#####"));

                    NKCT_TOPLAM = NKCT_TOPLAM / Katilimci_Sayisi;

                    Hucreye_Yazdirma_Metodu(Katilimci_Sayisi + 2, Madde_Sayisi + 8, NKCT_TOPLAM.ToString("0.#####"));

                    ONCT_TOPLAM = ONCT_TOPLAM / Katilimci_Sayisi;

                    Hucreye_Yazdirma_Metodu(Katilimci_Sayisi + 2, Madde_Sayisi + 9, ONCT_TOPLAM.ToString("0.#####"));

                    CVPSZCT_TOPLAM = CVPSZCT_TOPLAM / Katilimci_Sayisi;

                    Hucreye_Yazdirma_Metodu(Katilimci_Sayisi + 2, Madde_Sayisi + 10, CVPSZCT_TOPLAM.ToString("0.#####"));

                    ARADCT_TOPLAM = ARADCT_TOPLAM / Katilimci_Sayisi;

                    Hucreye_Yazdirma_Metodu(Katilimci_Sayisi + 2, Madde_Sayisi + 11, ARADCT_TOPLAM.ToString("0.#####"));

                    //Hücreleri kilitleme ve renk ayarı.

                    //kullanılmayan alanlar kullanıma kapatılır.

                    string sutun = harfler[Madde_Sayisi + 10].ToString() + (Katilimci_Sayisi + 2).ToString();

                    Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
                    Excel.Range fr;

                    fr = ws.get_Range("A1", sutun);
                    fr.Borders.Weight = 2;
                    fr.Borders.Color = Color.Black;
                    fr.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LightBlue);

                    Globals.ThisAddIn.Application.Cells.Locked = true;
                    Globals.ThisAddIn.Application.get_Range("A1", sutun).Locked = false;
                    Excel.Worksheet Sheet1 = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                    Sheet1.Protect();
                }
            }
            else MessageBox.Show("Hesaplama İşlemini '" + IslemSayfasi + "' İçinde Yapabilirsiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void CVPSZCT_Metod(ArrayList liste, int SiraNo)
        {//Cevapsız Bırakma.
            double Cvpsz_Sayac = 0;
            foreach (var item in liste)
            {
                if (item == "" || item == " " || item == null || item == "0") Cvpsz_Sayac++;
            }
            CVPSZCT = Cvpsz_Sayac / Madde_Sayisi;
            Sonuclar[(SiraNo - 1), 7] = CVPSZCT;
        }

        public void ONCT_Metod(ArrayList liste, int SiraNo)
        {//Orta Nokta Cevaplama Tarzı.
            double Onct_Sayac = 0, Cevaplanmayan_Sayac = 0;
            int Orta_Nokta = (Anket_Tipi + 1) / 2;
            foreach (var item in liste)
            {
                if (Convert.ToInt32(item) == Orta_Nokta) Onct_Sayac++;
                else if (item == "" || item == " " || item == null || item == "0") Cevaplanmayan_Sayac++;

            }
            ONCT = Onct_Sayac / (Madde_Sayisi - Cevaplanmayan_Sayac);
            Sonuclar[(SiraNo - 1), 6] = ONCT;
        }

        public void KOCT_Metod(ArrayList liste, int SiraNo)
        {//Katılımcı Olmayan Cevaplama Tarzı.
            double Koct_Sayac = 0, Cevaplanmayan_Sayac = 0;
            int Orta_Nokta = (Anket_Tipi + 1) / 2;
            foreach (var item in liste)
            {
                if (Convert.ToInt32(item) < Orta_Nokta && item != "0") Koct_Sayac++;
                else if (item == "" || item == " " || item == null || item == "0") Cevaplanmayan_Sayac++;
            }
            KOCT = Koct_Sayac / (Madde_Sayisi - Cevaplanmayan_Sayac);
            Sonuclar[(SiraNo - 1), 4] = KOCT;
        }

        public void KCT_Metod(ArrayList liste, int SiraNo)
        {//Katılımcı Cevaplama Tarzı.
            double Kct_Sayac = 0, Cevaplanmayan_Sayac = 0;
            int Orta_Nokta = (Anket_Tipi + 1) / 2;
            foreach (var item in liste)
            {
                if (Convert.ToInt32(item) > Orta_Nokta && item != "0") Kct_Sayac++;
                else if (item == "" || item == " " || item == null || item == "0") Cevaplanmayan_Sayac++;
            }
            KCT = Kct_Sayac / (Madde_Sayisi - Cevaplanmayan_Sayac);
            Sonuclar[(SiraNo - 1), 3] = KCT;
        }

        public void EKCT_Metod(ArrayList liste, int SiraNo)
        {//Ekstrem Katılımcı Cevaplama Tarzı.
            double Ekct_Sayac = 0, Cevaplanmayan_Sayac = 0;
            foreach (var item in liste)
            {//cevaplayıcının boş soru sayısı bulur.
                //cevaplayıcının uç nokta sayısı bulur.
                //uç nokta sayısı ile cevapladığı soru sayısı oranlanır.
                string Siradaki = Convert.ToString(item);
                if (Siradaki == Anket_Tipi.ToString()) Ekct_Sayac++;
                else if (Siradaki == "" || Siradaki == " " || Siradaki == null || Convert.ToInt32(item) == 0) Cevaplanmayan_Sayac++;
            }
            EKCT = Ekct_Sayac / (Madde_Sayisi - Cevaplanmayan_Sayac);
            Sonuclar[(SiraNo - 1), 1] = EKCT;
        }

        public void EKOCT_Metod(ArrayList liste, int SiraNo)
        {//Ekstrem Katılımcı Olmayan Cevaplama Tarzı.
            double Ekoct_Sayac = 0, Cevaplanmayan_Sayac = 0;
            foreach (var item in liste)
            {
                string Siradaki = Convert.ToString(item);
                if (Siradaki == "1") Ekoct_Sayac++;
                else if (Siradaki == "" || Siradaki == " " || Siradaki == null || Siradaki == "0") Cevaplanmayan_Sayac++;
            }
            EKOCT = Ekoct_Sayac / (Madde_Sayisi - Cevaplanmayan_Sayac);
            Sonuclar[(SiraNo - 1), 2] = EKOCT;
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            if (FormKontrol == false)//aynı anda 2 tane veri temizleme penceresinin açık olmasını engellemek için kullanılır.
            {
                MTemizlemeFormu = new MissingTemizlemeFormu();
                foreach (var item in Secilecek_Liste)
                {//veri temizlemek için açılan pencereye maddeleri gönderir.
                    MTemizlemeFormu.listBox1.Items.Add("Madde " + item.ToString());
                }
                FormKontrol = true;
                MTemizlemeFormu.Show();
            }
            else MessageBox.Show("Açık Pencereleri Kapatınız.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            if (FormKontrol == false)
            {
                KSoruFormu = new KontrolSorusuEklemeFormu();
                foreach (var item in Secilecek_Liste)
                {
                    KSoruFormu.listBox1.Items.Add("Madde " + item.ToString());
                }
                FormKontrol = true;
                KSoruFormu.Show();
            }
            else MessageBox.Show("Açık Pencereleri Kapatınız.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void MissingTemizlemeMetodu()
        {
            MTemizlemeFormu.progressBar1.Maximum = Convert.ToInt32(Katilimci_Sayisi) * Convert.ToInt32(Madde_Sayisi);//progressbar ın maksimum değerini atar.
            if (Convert.ToInt32(Secilen_Liste.Count.ToString()) > 0)//Veri temizleme yapmak için en az bir tane soru seçilmesi gerekir.
            {
                MTemizlemeFormu.progressBar1.Value = 0;//işlem sürecini gösteren progressbar sıfırlanır.
                MTemizlemeFormu.Cursor = Cursors.No;
                Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
                string AktifSayfa = activeWorksheet.Name;//aktif sayfanın adına ulaşır.
                if (IslemSayfasi == AktifSayfa)//aktif sayfa işlem sayfası ise yani verilerin girildiği sayfa açık ise veri temizleme işlemine devam eder.
                {
                    SilinecekIndisler = new ArrayList();
                    for (int i = 1; i <= Katilimci_Sayisi; i++)
                    {//Her satır için.
                        for (int j = 1; j <= Madde_Sayisi; j++)//Her Sütun için.
                        {//Her hücredeki değer için tek tek bakar ve seçilen listede varsa indis numarasını tutar.
                            string siradaki = Hucrelerden_Okuma_Metodu((i + 1), (j + 1));
                            if (siradaki == "0" || siradaki == "")
                            {
                                int kolon = j, flag = 0;
                                foreach (var item in Secilen_Liste)
                                {
                                    if (kolon.ToString() == item.ToString())
                                    {
                                        flag++;
                                    }
                                }
                                if (flag > 0) { SilinecekIndisler.Add(i); break; }
                            }
                            try
                            {
                                MTemizlemeFormu.progressBar1.Value++;
                            }
                            catch { }
                        }
                    }
                    MTemizlemeFormu.progressBar1.Maximum += SilinecekIndisler.Count * Madde_Sayisi;//işlem sürecini gösteren progressbarı günceller.
                    int[,] sil = new int[Convert.ToInt32(SilinecekIndisler.Count), Madde_Sayisi];
                    for (int i = 0; i < SilinecekIndisler.Count; i++)
                    {
                        for (int j = 0; j < Madde_Sayisi; j++)
                        {//silinecek verileri bir dizide tutar.
                            sil[i, j] = Convert.ToInt32(Hucrelerden_Okuma_Metodu((Convert.ToInt32(SilinecekIndisler[i]) + 1), (j + 2)));
                            try
                            {
                                MTemizlemeFormu.progressBar1.Value++;
                            }
                            catch { }
                        }
                    }

                    MTemizlemeFormu.progressBar1.Maximum += SilinecekIndisler.Count;//işlem sürecini gösteren progressbarı günceller.
                    Excel.Worksheet wsSil = Globals.ThisAddIn.Application.ActiveSheet;//aktif sayfaya ulaşır.
                    Excel.Range row;
                    int[] rows = new int[SilinecekIndisler.Count];
                    int SilinenSayac = 0;//bu değişkenin amacı excelde bir satır silinince yerine bir alttaki satır geldiği için listedeki silinecek satırdan
                    //daha önce silinenlerin sayısını çıkartmamız gerekir.
                    foreach (var item in SilinecekIndisler)
                    {//silinecek indisleri bulup siler.
                        row = (Excel.Range)wsSil.Rows[Convert.ToInt32(item.ToString()) + 1 - SilinenSayac];
                        row.Delete();
                        SilinenSayac++;
                        try
                        {
                            MTemizlemeFormu.progressBar1.Value++;
                        }
                        catch { }
                    }
                    if (SilinenSayac > 0)//devam etmesi için en az 1 tane veri silinmesi gerekir.
                    {
                        Katilimci_Sayisi = Katilimci_Sayisi - SilinecekIndisler.Count;//katılımcı sayısı silinenler olduğu için güncellenir.

                        Excel.Workbook wsTemizlenenler = Globals.ThisAddIn.Application.ActiveWorkbook;
                        Excel.Worksheet x = wsTemizlenenler.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        x.Name = YeniTemizlenenlerSayfasiOlustur();//yeni bir temizlenenler sayfası oluşturulur.

                        for (int i = 0; i < SilinecekIndisler.Count; i++)
                        {
                            for (int j = 0; j < Madde_Sayisi; j++)
                            {//yeni temizlenenler sayfasına silinecek veriler taşınır.
                                x.Cells[i + 1, j + 1] = sil[i, j].ToString();
                                try
                                {
                                    MTemizlemeFormu.progressBar1.Value++;
                                }
                                catch { }
                            }
                        }
                        if (MTemizlemeFormu.progressBar1.Value < MTemizlemeFormu.progressBar1.Maximum)
                            for (int i = MTemizlemeFormu.progressBar1.Value; i < MTemizlemeFormu.progressBar1.Maximum; i++)
                            {
                                MTemizlemeFormu.progressBar1.Value++;
                            }
                        MTemizlemeFormu.Cursor = Cursors.Default;
                        //Bilgi mesajı verir.
                        MessageBox.Show(SilinecekIndisler.Count.ToString() + " Veri " + x.Name.ToString() + " Sayfasına Taşındı.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {//Hiç veri silinmedi ise bilgi verir.
                        MTemizlemeFormu.Cursor = Cursors.Default;
                        MessageBox.Show("Silinecek Hiç Veri Bulunamadı.");
                        if (MTemizlemeFormu.progressBar1.Value < MTemizlemeFormu.progressBar1.Maximum)
                            for (int i = MTemizlemeFormu.progressBar1.Value; i < MTemizlemeFormu.progressBar1.Maximum; i++)
                            {
                                MTemizlemeFormu.progressBar1.Value++;
                            }
                    }
                }
                else MessageBox.Show("Veri Temizleme İşlemini '" + IslemSayfasi + "' İçinde Yapabilirsiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else MessageBox.Show("Madde Listesinden Seçim Yapınız.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void KontrolSorusuMetodu()
        {//Bu metodun amacı kullanıcının seçtiği maddeleri işaretleyen kullanıcıları temizlemek.
            KSoruFormu.progressBar1.Maximum = Convert.ToInt32(Katilimci_Sayisi) * Convert.ToInt32(Madde_Sayisi);
            if (Secilen_Liste.Count > 0)//en az 1 tane madde seçilmesi gerekir.
            {
                KSoruFormu.progressBar1.Value = 0;
                KSoruFormu.Cursor = Cursors.No;
                Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
                string AktifSayfa = activeWorksheet.Name;
                if (AktifSayfa == IslemSayfasi)//Doğru Sayfada Hesaplama Yapması İçin Kullanılır.
                {
                    SilinecekIndisler = new ArrayList();
                    for (int i = 1; i <= Katilimci_Sayisi; i++)
                    {//Bütün satırlar için.
                        for (int j = 1; j <= Madde_Sayisi; j++)
                        {//Her bir satırın sütunları için.
                            string siradaki = Hucrelerden_Okuma_Metodu((i + 1), (j + 1));
                            if (siradaki != "0")
                            {//sıradaki hücre sıfırdan farklı ise yani kullanıcı soruyu işaretlemiş ise o satırın indis numarası tutulur.
                                int flag = 0;
                                foreach (var item in Secilen_Liste)
                                {
                                    if (j.ToString() == item.ToString()) flag++;
                                }
                                if (flag > 0) { SilinecekIndisler.Add(i); break; }
                            }
                            try
                            {
                                KSoruFormu.progressBar1.Value++;
                            }
                            catch { }
                        }
                    }
                    if (SilinecekIndisler.Count > 0)//silinecek satır var ise devam eder.
                    {
                        KSoruFormu.progressBar1.Maximum += SilinecekIndisler.Count * Madde_Sayisi + SilinecekIndisler.Count;
                        int[,] sil = new int[Convert.ToInt32(SilinecekIndisler.Count), Madde_Sayisi];
                        for (int i = 0; i < SilinecekIndisler.Count; i++)
                        {
                            for (int j = 0; j < Madde_Sayisi; j++)
                            {//silinecek satırların değerlerini tutar.
                                sil[i, j] = Convert.ToInt32(Hucrelerden_Okuma_Metodu((Convert.ToInt32(SilinecekIndisler[i]) + 1), (j + 2)));
                                try
                                {
                                    KSoruFormu.progressBar1.Value++;
                                }
                                catch { }
                            }
                        }
                        Excel.Worksheet wsSil = Globals.ThisAddIn.Application.ActiveSheet;
                        Excel.Range row;
                        int[] rows = new int[SilinecekIndisler.Count];
                        int SilinenSayac = 0;
                        foreach (var item in SilinecekIndisler)
                        {//silinmesi gereken satırları siler.
                            row = (Excel.Range)wsSil.Rows[Convert.ToInt32(item.ToString()) + 1 - SilinenSayac];
                            row.Delete();
                            SilinenSayac++;
                            try
                            {
                                KSoruFormu.progressBar1.Value++;
                            }
                            catch { }
                        }
                        Katilimci_Sayisi = Katilimci_Sayisi - SilinecekIndisler.Count;//katılımcı sayısını günceller.

                        Excel.Workbook wsTemizlenenler = Globals.ThisAddIn.Application.ActiveWorkbook;
                        Excel.Worksheet x = wsTemizlenenler.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        x.Name = YeniTemizlenenlerSayfasiOlustur();//yeni temizlenenler sayfası oluşturur.
                        KSoruFormu.progressBar1.Maximum += SilinecekIndisler.Count * Madde_Sayisi;
                        for (int i = 0; i < SilinecekIndisler.Count; i++)
                        {
                            for (int j = 0; j < Madde_Sayisi; j++)
                            {
                                x.Cells[i + 1, j + 1] = sil[i, j].ToString();//yeni oluşturulan sayfaya silinecek verileri yazdırır.
                                try
                                {
                                    KSoruFormu.progressBar1.Value++;
                                }
                                catch { }
                            }
                        }
                        KSoruFormu.Cursor = Cursors.Default;
                        if (KSoruFormu.progressBar1.Value < KSoruFormu.progressBar1.Maximum)
                            for (int i = KSoruFormu.progressBar1.Value; i < KSoruFormu.progressBar1.Maximum; i++)
                            {
                                KSoruFormu.progressBar1.Value++;
                            }
                        //bilgi mesajı.
                        MessageBox.Show(SilinecekIndisler.Count.ToString() + " Veri " + x.Name.ToString() + " Sayfasına Taşındı.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {//hiç veri silinmediyse.
                        KSoruFormu.Cursor = Cursors.Default;
                        MessageBox.Show("Silinecek Hiç Bir Veri Bulunamadı.");
                        if (KSoruFormu.progressBar1.Value < KSoruFormu.progressBar1.Maximum)
                            for (int i = KSoruFormu.progressBar1.Value; i < KSoruFormu.progressBar1.Maximum; i++)
                            {
                                KSoruFormu.progressBar1.Value++;
                            }
                    }

                }
                else MessageBox.Show("Veri Temizleme İşlemini '" + IslemSayfasi + "' İçinde Yapabilirsiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else MessageBox.Show("Madde Listesinden Seçim Yapınız.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public string YeniTemizlenenlerSayfasiOlustur()
        {
            int sayac = 0;
            string SayfaAdi = "Temizlenenler";
            //((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[1]).Name.ToString();
            Boolean Flag = true;
            while (Flag)
            {
                Flag = false;
                sayac++;
                SayfaAdi = "Temizlenenler" + sayac.ToString();
                for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Count; i++)
                {
                    if (SayfaAdi == ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[i]).Name.ToString())
                    {
                        Flag = true;
                    }
                }
            }
            return SayfaAdi;
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            YardimPaneli yardimpaneli = new YardimPaneli();
            yardimpaneli.AllowDrop = true;
            yardimpaneli.Show();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Hakkimizda hakkimizda = new Hakkimizda();
            hakkimizda.Show();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {//cevapları alır.
            ArrayList Guvenirlilik_Oranlari=new ArrayList();

            int[,] veriler = new int[Katilimci_Sayisi, Madde_Sayisi];
            for (int i = 0; i < Katilimci_Sayisi; i++)
            {
                for (int j = 0; j < Madde_Sayisi; j++)
                {
                    veriler[i, j] = Convert.ToInt32(Hucrelerden_Okuma_Metodu((i + 2), (j + 2)));
                }
            }

            Guvenirlilik_Oranlari.Add(Guvenilirlik_Orani_Hesapla(veriler));
            try
            {
                for (int k = 0; k < Madde_Sayisi; k++)
                {
                    int satir = 0, sutun = 0;
                    int[,] veriler2 = new int[Katilimci_Sayisi, Madde_Sayisi-1];

                    for (int i = 0; i < Katilimci_Sayisi; i++)
                    {
                        sutun = 0;
                        for (int j = 0; j < Madde_Sayisi; j++)
                        {
                            if (k != j)
                            {
                                veriler2[satir, sutun] = Convert.ToInt32(Hucrelerden_Okuma_Metodu((i + 2), (j + 2)));
                                sutun++;
                            }
                        }
                        satir++;
                    }
                    Guvenirlilik_Oranlari.Add(Guvenilirlik_Orani_Hesapla(veriler2));
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }

            Excel.Workbook wsTemizlenenler = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet x = wsTemizlenenler.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            x.Cells[1, 1] = "Güvenirlilik Oranı: " + Guvenirlilik_Oranlari[0].ToString();

            x.Cells[2, 1] = "Madde Silinirse Güvenilirlik Oranı";
            for (int i = 1; i < Guvenirlilik_Oranlari.Count; i++)
            {
                x.Cells[i+2,1]="Madde"+i.ToString();
                x.Cells[i + 2, 2] = Guvenirlilik_Oranlari[i].ToString();
            }
            x.Cells.AutoFit();
        }

        private double Guvenilirlik_Orani_Hesapla(int[,] veriler)
        {
            ArrayList Maddeler_Toplaminin_Varyansi = new ArrayList();
            ArrayList Maddeler_Toplami_Listesi = new ArrayList();

            for (int j = 0; j < veriler.GetLength(1); j++)
            {
                Maddeler_Toplami_Listesi.Clear();
                for (int i = 0; i < veriler.GetLength(0); i++)
                {
                    Maddeler_Toplami_Listesi.Add(veriler[i, j]);
                }
                Maddeler_Toplaminin_Varyansi.Add(Varyans(Maddeler_Toplami_Listesi));
            }

            double Varyanslar_Toplami = 0;
            foreach (var item in Maddeler_Toplaminin_Varyansi)
            {
                Varyanslar_Toplami += Convert.ToDouble(item);
            }


            ArrayList Test_Skor_Toplami = new ArrayList();

            for (int i = 0; i < veriler.GetLength(0); i++)
            {
                int toplam = 0;
                for (int j = 0; j < veriler.GetLength(1); j++)
                {
                    toplam += veriler[i, j];
                }
                Test_Skor_Toplami.Add(toplam);
            }
            double Test_Skor_Toplam_Varyansi = Varyans(Test_Skor_Toplami);


            double Guvenilirlik_Orani = (Convert.ToDouble(veriler.GetLength(1)) / (Convert.ToDouble(veriler.GetLength(1)) - 1)) * (1 - (Varyanslar_Toplami / Test_Skor_Toplam_Varyansi));
            return Guvenilirlik_Orani;
        }

       

        public double Varyans(ArrayList liste)
        {

            double ortalama = 0, KarelerToplami = 0;
            ArrayList SapmalarKaresi = new ArrayList();

            foreach (var item in liste)
            {
                ortalama += Convert.ToDouble(item);
            }
            ortalama /= liste.Count;

            foreach (var item in liste)
            {
                SapmalarKaresi.Add(Math.Pow(Convert.ToDouble(item) - ortalama, 2));
            }

            foreach (var item in SapmalarKaresi)
            {
                KarelerToplami += Convert.ToDouble(item);
            }
            return KarelerToplami / liste.Count;

        }
    }
}
