using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CevaplamaTarziHesaplamasi
{
    public partial class KontrolSorusuEklemeFormu : Form
    {
        public KontrolSorusuEklemeFormu()
        {
            InitializeComponent();
        }

        Ribbon1 r = Globals.Ribbons.Ribbon1;

        private void KontrolSorusuEklemeFormu_Load(object sender, EventArgs e)
        {
            r.Secilen_Liste.Clear();
            Control.CheckForIllegalCrossThreadCalls = true;
            if (listBox1.Items.Count > 0)
                listBox1.SelectedIndex = 0;
            button2.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex > -1)
            {
                string secilen = listBox1.SelectedItem.ToString();
                string[] parcala = secilen.Split(' ');
                r.Secilen_Liste.Add(parcala[1].ToString());
                listBox2.Items.Add(secilen);
                listBox1.Items.Remove(secilen);
                button2.Enabled = true;
                ListboxSort(listBox2);
                listBox1.SelectedIndex = 0;
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
                ListboxSort(listBox1);
            }
            else MessageBox.Show("Bir Madde Seçiniz.");
            if (listBox2.Items.Count == 0) button2.Enabled = false;
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

        private void button3_Click(object sender, EventArgs e)
        {
            r.KontrolSorusuMetodu();
        }

        private void KontrolSorusuEklemeFormu_FormClosed(object sender, FormClosedEventArgs e)
        {
            r.FormKontrol = false;
        }
    }
}
