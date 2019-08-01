using CevaplamaTarziHesaplamasi.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CevaplamaTarziHesaplamasi
{
    public partial class YardimPaneli : Form
    {
        public YardimPaneli()
        {
            InitializeComponent();
        }

        int pivot = 0;

        private void YardimPaneli_Load(object sender, EventArgs e)
        {
            ResimGuncelle(true);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ResimGuncelle(false);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ResimGuncelle(true);
        }

        public void ResimGuncelle(bool islem)
        {//True => ileri gitme işlemi. False => Geri gitme işlemi.
            if (islem && pivot < 5)
                pivot++;
            else if (islem == false && pivot > 1)
                pivot--;

            if (pivot == 1)
            {
                label1.ForeColor = Color.Black;
                label2.ForeColor = label3.ForeColor = label4.ForeColor = label5.ForeColor = Color.White;
                pictureBox1.BackgroundImage = Resources.Adim1;
            }
            else if (pivot == 2)
            {
                label2.ForeColor = Color.Black;
                label1.ForeColor = label3.ForeColor = label4.ForeColor = label5.ForeColor = Color.White;
                pictureBox1.BackgroundImage = Resources.Adim2;
            }
            else if (pivot == 3)
            {
                label3.ForeColor = Color.Black;
                label2.ForeColor = label1.ForeColor = label4.ForeColor = label5.ForeColor = Color.White;
                pictureBox1.BackgroundImage = Resources.Adim3;
            }
            else if (pivot == 4)
            {
                label4.ForeColor = Color.Black;
                label2.ForeColor = label3.ForeColor = label1.ForeColor = label5.ForeColor = Color.White;
                pictureBox1.BackgroundImage = Resources.Adim4;
            }
            else if (pivot == 5)
            {
                label5.ForeColor = Color.Black;
                label2.ForeColor = label3.ForeColor = label4.ForeColor = label1.ForeColor = Color.White;
                pictureBox1.BackgroundImage = Resources.Adim5;
            }
        }
    }
}
