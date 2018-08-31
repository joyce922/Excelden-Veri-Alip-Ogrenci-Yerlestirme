using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;


namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        private List<string> dersler = new List<string>();
        private List<int> kontenjanlar = new List<int>();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dersler.Contains(textBoxDERS.Text))
            {
                errorProvider1.SetError(textBoxDERS, " Bu ders zaten kayıtlı ! ");
            }
            else
            {
                listBoxDers.Items.Add(textBoxDERS.Text + " Kontenjan :" + numericUpDownKontenjan.Value);
                dersler.Add(textBoxDERS.Text);
                kontenjanlar.Add(Convert.ToInt32(numericUpDownKontenjan.Value));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dersler.RemoveAt(listBoxDers.SelectedIndex);
            kontenjanlar.RemoveAt(listBoxDers.SelectedIndex);
            listBoxDers.Items.Remove(listBoxDers.SelectedItem);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog opfd = new OpenFileDialog();
            if (opfd.ShowDialog() == DialogResult.OK)
                textBoxSeç.Text = opfd.FileName;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string stringconn = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + textBoxSeç.Text + "; Extended Properties='Excel 12.0;HDR=YES'";
            OleDbConnection conn = new OleDbConnection(stringconn);
            if (textBoxSeç.Text != "")
            {
                OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + textBoxAktar.Text + "$]", conn);
                DataTable dt = new DataTable ();
                 da.Fill(dt);
                 DataTable grd = new DataTable();
                 grd.Columns.Add("ID", typeof(int));
                 grd.Columns.Add("Ad Soyad");
                 grd.Columns.Add("Ortalama", typeof(decimal));
                 grd.Columns.Add("Tercih");
                 grd.Columns.Add("Kayıtlı Ders");
                 DataRow[] drs = dt.Select("", " Ortalama DESC");
                 for (int i = 0; i < drs.Length; i++)
                 {
                     grd.Rows.Add(drs[i]["ID"], drs[i]["Ad Soyad"], drs[i]["Ortalama"], drs[i]["Tercih"], "");
                 }
                 dataGridView1.DataSource = grd;
                 int x = 0 ;
                 int toplam = 0;
                 string[] tercihler;

                 for (int i = 0; i < kontenjanlar.Count; i++)
                 {
                     toplam = toplam + kontenjanlar[i];                    
                 }

                 if (toplam != grd.Rows.Count)
                 {
                     errorProvider2.SetError(numericUpDownKontenjan, " Eksik veya fazla kontenjan sayısı girdiniz");
                     return;
                 }

                 for (int i = 0; i < grd.Rows.Count; i++)
                 {
                     tercihler = grd.Rows[i][3].ToString().Split(',');
                     if (dersler.Count < tercihler.Count())
                     {
                         for (int q = 0; q < dersler.Count; q++)
                         {
                             for (int w = 0; w < tercihler.Count(); w++)
                             {
                                 if (tercihler[w] != dersler[q])
                                 {
                                     dersler.Add(tercihler[w]);
                                     kontenjanlar.Add(0);
                                 }
                             }
                         }
                     }
                         for (int b = 0; b < dersler.Count ; b++)
                         {
                             if (tercihler[x] == dersler[b])
                             {
                                 if (kontenjanlar[b] > 0)
                                 {
                                     x = 0;
                                     grd.Rows[i]["Kayıtlı Ders"] = dersler[b];                                    
                                     kontenjanlar[b] = kontenjanlar[b] - 1;
                                     break;
                                     
                                 }
                                 else 
                                 {
                                     if (x < (tercihler.Count()-1))
                                     {
                                         x = x + 1;
                                         b = -1;
                                     }
                                 }
                             }
                         }
                 }
            }
            else
            {
                MessageBox.Show("HATA !");
            }   
        }
    }
}
