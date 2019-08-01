using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Collections;

namespace CevaplamaTarziHesaplamasi
{
    public partial class MissingTemizlemeFormu : Form
    {
        public MissingTemizlemeFormu()
        {
            InitializeComponent();
        }

        Ribbon1 r = Globals.Ribbons.Ribbon1;
        ArrayList list2 = new ArrayList();

        private void MissingTemizlemeFormu_Load(object sender, EventArgs e)
        {
            r.Secilen_Liste.Clear();
            Control.CheckForIllegalCrossThreadCalls = true;
            if (listBox1.Items.Count > 0)
                listBox1.SelectedIndex = 0;
            button2.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >= 0)
            {
                int SonSecilenIndis = listBox1.SelectedIndex;
                string secilen = listBox1.SelectedItem.ToString();
                string[] parcala = secilen.Split(' ');
                r.Secilen_Liste.Add(parcala[1].ToString());
                listBox2.Items.Add(secilen);
                listBox1.Items.Remove(secilen);
                if (listBox1.Items.Count > 0) listBox1.SelectedIndex = SonSecilenIndis - 1;
                button2.Enabled = true;
                ListboxSort(listBox2);
            }
            else MessageBox.Show("Bir Madde Seçiniz.");
            if (listBox1.Items.Count == 0) button1.Enabled = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex >= 0)
            {
                string secilen = listBox2.SelectedItem.ToString();
                string[] parcala = secilen.Split(' ');
                r.Secilen_Liste.Remove(parcala[1].ToString());
                listBox1.Items.Add(secilen);
                listBox2.Items.Remove(secilen);
                if (listBox2.Items.Count > 0) listBox2.SelectedIndex = 0;
                button1.Enabled = true;
                if (listBox2.Items.Count > 0) listBox2.SelectedIndex = listBox2.Items.Count - 1;
                ListboxSort(listBox1);
            }
            else MessageBox.Show("Bir Madde Seçiniz.");
            if (listBox2.Items.Count == 0) button2.Enabled = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            r.MissingTemizlemeMetodu();
        }

        public void ListboxSort(ListBox liste)
        {
            ArrayList GelenVeriler = new ArrayList();
            foreach (var item in liste.Items)
            {
                GelenVeriler.Add(Convert.ToInt32(item.ToString().Split(' ').ToList()[1]));
            }
            GelenVeriler.Sort();
            liste.Items.Clear();
            foreach (var item in GelenVeriler)
            {
                liste.Items.Add("Madde " + item.ToString());
            }
        }

        private void MissingTemizlemeFormu_FormClosed(object sender, FormClosedEventArgs e)
        {
            r.FormKontrol = false;
        }
    }
}
