using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Range = System.Range;
using GroupBox = System.Windows.Forms.GroupBox;
using Rectangle = System.Drawing.Rectangle;
using Point = System.Drawing.Point;

namespace Ekrem_Gungor
{
    public partial class Form1 : Form
    {
        SqlConnection baglanti = new SqlConnection("Data Source=localhost\\SQLEXPRESS;Initial Catalog=ekremdb;Integrated Security=True");

        public Form1()
        {
            InitializeComponent();
        }
        void masrafGetir()
        {
            baglanti.Open();
            DataTable tablo = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM masraf", baglanti);
            da.Fill(tablo);
            dataGridView1.DataSource = tablo;
            baglanti.Close();

        }

        void masrafGetir2()
        {
            baglanti.Open();
            DataTable tablo = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM fatura", baglanti);
            da.Fill(tablo);
            dataGridView3.DataSource = tablo;
            baglanti.Close();

        }

        void masrafGetir3()
        {
            baglanti.Open();
            DataTable tablo = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM smasraf", baglanti);
            da.Fill(tablo);
            dataGridView4.DataSource = tablo;
            baglanti.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "" && textBox3.Text != "" && dateTimePicker1.Text != "")
            {
                string sql = "insert into masraf (mkalemi,mtutar,tarih) values(@mkalemi,@mtutar,@tarih)";
                SqlCommand komut = new SqlCommand(sql, baglanti);
                baglanti.Open();
                komut.Parameters.AddWithValue("@mkalemi", textBox2.Text);
                komut.Parameters.AddWithValue("@mtutar", double.Parse(textBox3.Text));
                komut.Parameters.AddWithValue("@tarih", dateTimePicker1.Value);
                komut.ExecuteNonQuery();
                baglanti.Close();
                masrafGetir();
            }
            else
            {
                MessageBox.Show("Tüm Alanlar Doldurulmalıdır");
            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            masrafGetir();
            renklendir();
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            //dataGridView1.Columns[0].Visible = false;

            dataGridView1.Columns[1].HeaderText = "Gider Kalemi";
            dataGridView1.Columns[2].HeaderText = "Tutar";
            dataGridView1.Columns[3].HeaderText = "Tarih";

            textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            dateTimePicker1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridView1.GetCellCount(DataGridViewElementStates.Selected) > 0)
            {
                String sorgu = "DELETE FROM masraf WHERE id=@id";
                SqlCommand komut = new SqlCommand(sorgu, baglanti);
                komut.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value));
                baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();
                masrafGetir();
            }
            else
            {
                MessageBox.Show("Silinecek Alanı Seçiniz");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "" && textBox3.Text != "" && dateTimePicker1.Text != "" && dataGridView1.GetCellCount(DataGridViewElementStates.Selected) > 0)
            {
                string sorgu = "UPDATE masraf SET mkalemi=@mkalemi, mtutar=@mtutar, tarih=@tarih WHERE id=@id";
                SqlCommand komut = new SqlCommand(sorgu, baglanti);
                komut.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value));
                komut.Parameters.AddWithValue("@mkalemi", textBox2.Text);
                komut.Parameters.AddWithValue("@mtutar", double.Parse(textBox3.Text));
                komut.Parameters.AddWithValue("@tarih", dateTimePicker1.Value);
                baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();
                masrafGetir();
            }
            else
            {
                MessageBox.Show("Önce Veri Eklenmelidir");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DateTime a, b;
            DateTime.TryParse(dateTimePicker3.Text, out a);
            DateTime.TryParse(dateTimePicker5.Text, out b);

            string sql = "SELECT mkalemi, mtutar, tarih FROM masraf WHERE tarih BETWEEN @tarih1 and @tarih2 and mkalemi=@aranan";
            DataTable dt2 = new DataTable();
            SqlDataAdapter daa = new SqlDataAdapter(sql, baglanti);
            baglanti.Open();
            daa.SelectCommand.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = a;
            daa.SelectCommand.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = b;
            daa.SelectCommand.Parameters.AddWithValue("@aranan", textBox4.Text);
            daa.Fill(dt2);
            dataGridView2.DataSource = dt2;
            baglanti.Close();
            dataGridView2.AllowUserToAddRows = false;
            double sum = 0;
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                sum = sum + double.Parse(dataGridView2.Rows[i].Cells[1].Value.ToString());

            }
            textBox6.Text = sum.ToString("0.00");
            datagridHeaderText();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DateTime a, b;
            DateTime.TryParse(dateTimePicker2.Text, out a);
            DateTime.TryParse(dateTimePicker4.Text, out b);

            string sql = "SELECT mkalemi, mtutar, tarih FROM masraf WHERE tarih BETWEEN @tarih1 and @tarih2";
            DataTable dt = new DataTable();
            SqlDataAdapter daa = new SqlDataAdapter(sql, baglanti);
            baglanti.Open();

            daa.SelectCommand.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = a;
            daa.SelectCommand.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = b;
            daa.Fill(dt);
            dataGridView2.DataSource = dt;
            baglanti.Close();
            dataGridView2.AllowUserToAddRows = false;
            double sum = 0;
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                sum = sum + double.Parse(dataGridView2.Rows[i].Cells[1].Value.ToString());

            }
            textBox7.Text = sum.ToString("0.00");
            datagridHeaderText();
        }

        void datagridHeaderText()
        {
            string[] str = { "Gider Kalemi", "Tutar", "Tarih" };
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
            {
                dataGridView2.Columns[i].HeaderText = str[i];
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                string sql = "insert into ciget (ciget,tarih) values(@ciget,@tarih)";
                SqlCommand komut = new SqlCommand(sql, baglanti);
                baglanti.Open();
                komut.Parameters.AddWithValue("@ciget", textBox1.Text);
                komut.Parameters.AddWithValue("@tarih", dateTimePicker7.Value);
                komut.ExecuteNonQuery();

                string sql2 = "insert into toplamciget (toplamciget) values(@toplamciget)";
                SqlCommand komut2 = new SqlCommand(sql2, baglanti);

                string sql3 = "SELECT * FROM ciget";
                DataTable dt9 = new DataTable();
                SqlDataAdapter daa = new SqlDataAdapter(sql3, baglanti);
                daa.Fill(dt9);
                double k = (double)dt9.Compute("SUM(ciget)", "");

                komut2.Parameters.AddWithValue("@toplamciget", k);
                komut2.ExecuteNonQuery();


                string sql6 = "SELECT * FROM pismiset";
                DataTable dt10 = new DataTable();
                SqlDataAdapter daa1 = new SqlDataAdapter(sql6, baglanti);
                daa1.Fill(dt10);
                if (dt10.Compute("SUM(pismiset)", "") != DBNull.Value && dt10.Compute("SUM(kalanet)", "") != DBNull.Value)
                {
                    double m = (double)dt10.Compute("SUM(pismiset)", "");
                    double n = (double)dt10.Compute("SUM(kalanet)", "");

                    double b = m + n;
                    double c = (k - b);

                    textBox12.Text = c.ToString();

                    baglanti.Close();

                }
                else
                {
                    textBox12.Text = k.ToString();
                    baglanti.Close();

                }

                MessageBox.Show("Miktar Eklendi", "Çiğ Et");
            }
            else
            {
                MessageBox.Show("Miktar Giriniz");
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            DateTime a, b;
            DateTime.TryParse(dateTimePicker6.Text, out a);
            DateTime.TryParse(dateTimePicker8.Text, out b);

            string sql = "SELECT ciget,tarih FROM ciget WHERE tarih BETWEEN @tarih1 and @tarih2";
            DataTable dt2 = new DataTable();
            SqlDataAdapter daa = new SqlDataAdapter(sql, baglanti);
            baglanti.Open();
            daa.SelectCommand.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = a;
            daa.SelectCommand.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = b;
            daa.Fill(dt2);
            textBox5.Text = dt2.Compute("SUM(ciget)", "").ToString();
            baglanti.Close();

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (textBox8.Text !="" && textBox9.Text != "") { 
            baglanti.Open();
            string sql2 = "SELECT TOP(1) toplamciget FROM toplamciget order by id desc";
            SqlCommand komut1 = new SqlCommand(sql2, baglanti);
            double a = (double)komut1.ExecuteScalar();


            string sql6 = "SELECT * FROM pismiset";
            DataTable dt10 = new DataTable();
            SqlDataAdapter daa1 = new SqlDataAdapter(sql6, baglanti);
            daa1.Fill(dt10);



            if (dt10.Compute("SUM(pismiset)", "") != DBNull.Value && dt10.Compute("SUM(kalanet)", "") != DBNull.Value)
            {

                double m = (double)dt10.Compute("SUM(pismiset)", "") + double.Parse(textBox8.Text);
                double n = (double)dt10.Compute("SUM(kalanet)", "") + double.Parse(textBox9.Text);

                double b = m + n;
                double c = (a - b);



                if (c < 0)
                {
                    MessageBox.Show("Çiğ Et Miktarı Yetersiz");
                    baglanti.Close();
                    return;
                }
                else
                {
                    if (c < 150)
                    {
                        MessageBox.Show("Çiğ Et Miktarı 150 kg ın Altında");
                    }
                    string sql = "insert into pismiset (pismiset, kalanet, tarih) values(@pismiset, @kalanet, @tarih)";
                    SqlCommand komut = new SqlCommand(sql, baglanti);
                    komut.Parameters.AddWithValue("@pismiset", double.Parse(textBox8.Text));
                    komut.Parameters.AddWithValue("@kalanet", double.Parse(textBox9.Text));
                    komut.Parameters.AddWithValue("@tarih", dateTimePicker9.Value);
                    komut.ExecuteNonQuery();

                    string sql5 = "insert into kalanciget (kalanciget) values(@kalanciget)";
                    SqlCommand komut5 = new SqlCommand(sql5, baglanti);
                    komut5.Parameters.AddWithValue("@kalanciget", c);
                    komut5.ExecuteNonQuery();
                    textBox12.Text = c.ToString();

                    baglanti.Close();

                }
            }
            else
            {
                double r = double.Parse(textBox8.Text);
                double s = double.Parse(textBox9.Text);
                double t = r + s;
                double p = a - t;
                if (p < 0)
                {
                    MessageBox.Show("Çiğ Et Miktarı Yetersiz");
                    baglanti.Close();
                    return;
                }
                else
                {
                    if (p < 150)
                    {
                        MessageBox.Show("Çiğ Et Miktarı 150 kg ın Altında");
                    }
                    string sql = "insert into pismiset (pismiset, kalanet, tarih) values(@pismiset, @kalanet, @tarih)";
                    SqlCommand komut = new SqlCommand(sql, baglanti);
                    komut.Parameters.AddWithValue("@pismiset", double.Parse(textBox8.Text));
                    komut.Parameters.AddWithValue("@kalanet", double.Parse(textBox9.Text));
                    komut.Parameters.AddWithValue("@tarih", dateTimePicker9.Value);
                    komut.ExecuteNonQuery();

                    string sql5 = "insert into kalanciget (kalanciget) values(@kalanciget)";
                    SqlCommand komut5 = new SqlCommand(sql5, baglanti);

                    komut5.Parameters.AddWithValue("@kalanciget", p);
                    komut5.ExecuteNonQuery();
                    textBox12.Text = p.ToString();

                    baglanti.Close();
                }
            }
        }else
            {
                MessageBox.Show("Alanlar boş olamaz");
            }
    }
        private void tabPage5_Enter(object sender, EventArgs e)
        {
            baglanti.Open();
            string sql2 = "SELECT TOP(1) toplamciget FROM toplamciget order by id desc";
            SqlCommand komut1 = new SqlCommand(sql2, baglanti);
            if (komut1.ExecuteScalar() != null)
            {
                double a = (double)komut1.ExecuteScalar();

                string sql6 = "SELECT * FROM pismiset";
                DataTable dt10 = new DataTable();
                SqlDataAdapter daa1 = new SqlDataAdapter(sql6, baglanti);
                daa1.Fill(dt10);
                if (dt10.Compute("SUM(pismiset)", "") != DBNull.Value && dt10.Compute("SUM(kalanet)", "") != DBNull.Value)
                {
                    double m = (double)dt10.Compute("SUM(pismiset)", "");
                    double n = (double)dt10.Compute("SUM(kalanet)", "");

                    double b = m + n;
                    double c = (a - b);

                    textBox12.Text = c.ToString();

                    baglanti.Close();

                }
                else
                {
                    textBox12.Text = a.ToString();
                    baglanti.Close();

                }

            }
            else
            {
                textBox12.Text = "0";
                baglanti.Close();
            }


        }

        private void button9_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            textBox10.Text = "";
            string sql7 = "SELECT * FROM pismiset";
            DataTable dt11 = new DataTable();
            SqlDataAdapter daa1 = new SqlDataAdapter(sql7, baglanti);
            daa1.Fill(dt11);

            if (comboBox1.SelectedIndex != -1) { 
            
            if (dt11.Compute("SUM(kalanet)", "") == DBNull.Value || dt11.Compute("SUM(pismiset)", "") == DBNull.Value)
            {
                baglanti.Close();
                MessageBox.Show("Pişmiş Et ve kalan Et Yok");
                return;
            }


            if (comboBox1.SelectedItem.ToString() == "kalanet" && dt11.Compute("SUM(kalanet)", "") != DBNull.Value)
            {
                DateTime t, r;
                DateTime.TryParse(dateTimePicker10.Text, out t);
                DateTime.TryParse(dateTimePicker11.Text, out r);
                string sql8 = "SELECT * FROM pismiset where tarih between @tarih1 and @tarih2";
                DataTable dt12 = new DataTable();
                SqlDataAdapter daa2 = new SqlDataAdapter(sql8, baglanti);
                daa2.SelectCommand.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = t;
                daa2.SelectCommand.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = r;
                daa2.Fill(dt12);

                double m = (double)dt12.Compute("SUM(kalanet)", "");
                label29.Text = "Kalan Et";
                textBox10.Text = m.ToString();
                baglanti.Close();
            }

            if (comboBox1.SelectedItem.ToString() == "pismiset" && dt11.Compute("SUM(pismiset)", "") != DBNull.Value)
            {
                DateTime k, l;
                DateTime.TryParse(dateTimePicker10.Text, out k);
                DateTime.TryParse(dateTimePicker11.Text, out l);
                string sql10 = "SELECT * FROM pismiset where tarih between @tarih1 and @tarih2";
                DataTable dt13 = new DataTable();
                SqlDataAdapter daa3 = new SqlDataAdapter(sql10, baglanti);
                daa3.SelectCommand.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = k;
                daa3.SelectCommand.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = l;
                daa3.Fill(dt13);

                double n = (double)dt13.Compute("SUM(pismiset)", "");
                label29.Text = "Pişmiş Et";
                textBox10.Text = n.ToString();
                baglanti.Close();
            }
            }
            else 
            {
                baglanti.Close();
                MessageBox.Show("Lütfen Seçim yapınız");
                return;
            }
           
        }

        private void button10_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            if (textBox15.Text == "")
            {
                baglanti.Close();
                MessageBox.Show("Lütfen Tutar Giriniz");
                return;
            }
            if (comboBox2.SelectedIndex != -1)
            {
                string sql33 = "insert into gelir (gelirturu, tutar, tarih) values(@gelirturu, @tutar, @tarih)";
                SqlCommand komut33 = new SqlCommand(sql33, baglanti);

                komut33.Parameters.AddWithValue("@gelirturu", comboBox2.SelectedItem.ToString());
                komut33.Parameters.AddWithValue("@tutar", double.Parse(textBox15.Text));
                komut33.Parameters.AddWithValue("@tarih", dateTimePicker7.Value);
                komut33.ExecuteNonQuery();
                MessageBox.Show("Gelir Eklendi", "Gelir");
                baglanti.Close();
            }
            else
            {
                baglanti.Close();
                MessageBox.Show("Lütfen Gelir Türü Seçin");
                return;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            baglanti.Open();


            int b = comboBox3.SelectedIndex;
            if (b == -1)
            {
                MessageBox.Show("Lütfen Gelir türü Seçiniz");
                baglanti.Close();
                return;
            }
            string a = comboBox3.SelectedItem.ToString();

            if (a == "kredi_karti")
            {
                DateTime h, i;
                DateTime.TryParse(dateTimePicker12.Text, out h);
                DateTime.TryParse(dateTimePicker14.Text, out i);
                string sql9 = "SELECT SUM(tutar) FROM gelir where tarih between @tarih1 and @tarih2 and gelirturu=@gelir";

                SqlCommand komut1 = new SqlCommand(sql9, baglanti);
                komut1.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = h;
                komut1.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = i;
                komut1.Parameters.AddWithValue("@gelir", a);
                textBox13.Text = komut1.ExecuteScalar().ToString();


                DateTime q, w;
                DateTime.TryParse(dateTimePicker12.Text, out q);
                DateTime.TryParse(dateTimePicker14.Text, out w);
                string sql10 = "SELECT SUM(tutar) FROM gelir where tarih between @tarih1 and @tarih2";

                SqlCommand komut2 = new SqlCommand(sql10, baglanti);
                komut2.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = q;
                komut2.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = w;
                textBox14.Text = komut2.ExecuteScalar().ToString();

                string sql11 = "SELECT SUM(tutar) FROM gelir";

                SqlCommand komut3 = new SqlCommand(sql11, baglanti);
                textBox16.Text = komut3.ExecuteScalar().ToString();
                
                baglanti.Close();
            }

            if (a == "nakit")
            {
                string sql9 = "SELECT SUM(tutar) FROM gelir where tarih between @tarih1 and @tarih2 and gelirturu=@gelir";

                DateTime h, i;
                DateTime.TryParse(dateTimePicker12.Text, out h);
                DateTime.TryParse(dateTimePicker14.Text, out i);

                SqlCommand komut1 = new SqlCommand(sql9, baglanti);
                komut1.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = h;
                komut1.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = i;
                komut1.Parameters.AddWithValue("@gelir", a);
                textBox13.Text = komut1.ExecuteScalar().ToString();

                string sql10 = "SELECT SUM(tutar) FROM gelir where tarih between @tarih1 and @tarih2";

                DateTime m, n;
                DateTime.TryParse(dateTimePicker12.Text, out m);
                DateTime.TryParse(dateTimePicker14.Text, out n);

                SqlCommand komut2 = new SqlCommand(sql10, baglanti);
                komut2.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = m;
                komut2.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = n;
                textBox14.Text = komut2.ExecuteScalar().ToString();

                string sql11 = "SELECT SUM(tutar) FROM gelir";

                SqlCommand komut3 = new SqlCommand(sql11, baglanti);
                textBox16.Text = komut3.ExecuteScalar().ToString();

                baglanti.Close();
            }

            if (a == "yemek_sepeti")
            {
                string sql9 = "SELECT SUM(tutar) FROM gelir where tarih between @tarih1 and @tarih2 and gelirturu=@gelir";

                DateTime m, n;
                DateTime.TryParse(dateTimePicker12.Text, out m);
                DateTime.TryParse(dateTimePicker14.Text, out n);

                SqlCommand komut1 = new SqlCommand(sql9, baglanti);
                komut1.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = m;
                komut1.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = n;
                komut1.Parameters.AddWithValue("@gelir", a);
                textBox13.Text = komut1.ExecuteScalar().ToString();

                string sql10 = "SELECT SUM(tutar) FROM gelir where tarih between @tarih1 and @tarih2";

                DateTime s, v;
                DateTime.TryParse(dateTimePicker12.Text, out s);
                DateTime.TryParse(dateTimePicker14.Text, out v);

                SqlCommand komut2 = new SqlCommand(sql10, baglanti);
                komut2.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = s;
                komut2.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = v;
                textBox14.Text = komut2.ExecuteScalar().ToString();

                string sql11 = "SELECT SUM(tutar) FROM gelir";

                SqlCommand komut3 = new SqlCommand(sql11, baglanti);
                textBox16.Text = komut3.ExecuteScalar().ToString();

                baglanti.Close();
            }



            baglanti.Close();
        }

        TimeSpan fark;
        double gunfark;
        private void button15_Click(object sender, EventArgs e)
        {
            if (textBox22.Text != "" && textBox19.Text != "" && dateTimePicker15.Text != "" && textBox20.Text != "" && dateTimePicker16.Text != "" && comboBox4.Text != "")
            {
                string sql = "insert into fatura (firmaAdi,faturano,faturatarih,tutar,vade,yuzde, kdv_mik, odendi) values(@firmaAdi, @faturano, @faturatarih, @tutar, @vade, @yuzde, @kdv_mik, @odendi)";
                SqlCommand komut = new SqlCommand(sql, baglanti);
                baglanti.Open();
                komut.Parameters.AddWithValue("@firmaAdi", textBox22.Text);
                komut.Parameters.AddWithValue("@faturano", textBox19.Text);
                komut.Parameters.AddWithValue("@faturatarih", dateTimePicker15.Value);
                komut.Parameters.AddWithValue("@tutar", double.Parse(textBox20.Text));
                komut.Parameters.AddWithValue("@vade", dateTimePicker16.Value);
                komut.Parameters.AddWithValue("@yuzde", double.Parse(comboBox4.Text.ToString()));
                komut.Parameters.AddWithValue("@kdv_mik", double.Parse(textBox20.Text)* double.Parse(comboBox4.Text.ToString()));
                komut.Parameters.AddWithValue("@odendi", false);
                komut.ExecuteNonQuery();
                baglanti.Close();
                masrafGetir2();
                renklendir();
                
               
                toplamFatura();
                tumGelir();
                if (textBox18.Text != "" && textBox17.Text != "")
                {
                    double z = double.Parse(textBox18.Text) - double.Parse(textBox17.Text);
                    textBox21.Text = z.ToString();
                }
                


            }
            else
            {
                MessageBox.Show("Tüm Bölümün Doldurulması Zorunludur", "Fatura");
                baglanti.Close();
                return;
            }
        }

        void tumGelir()
        {
            baglanti.Open();
            string sql11 = "SELECT SUM(tutar) FROM gelir";
            SqlCommand komut3 = new SqlCommand(sql11, baglanti);
            textBox18.Text = komut3.ExecuteScalar().ToString();
            baglanti.Close();
        }

       
        private void button13_Click(object sender, EventArgs e)
        {
            if (dataGridView1.GetCellCount(DataGridViewElementStates.Selected) > 0)
            {
                String sorgu = "DELETE FROM fatura WHERE id=@id";
                SqlCommand komut = new SqlCommand(sorgu, baglanti);
                komut.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView3.CurrentRow.Cells[0].Value));
                baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();
                masrafGetir2();
                renklendir();
                toplamFatura();
                tumGelir();
                if (textBox18.Text != "" && textBox17.Text != "")
                {
                    double z = double.Parse(textBox18.Text) - double.Parse(textBox17.Text);
                    textBox21.Text = z.ToString();
                }
            }
            else
            {
                MessageBox.Show("Silinecek Alan Seçiniz");
            }
            
        }

        private void tabPage4_Enter(object sender, EventArgs e)
        {
            masrafGetir2();
            renklendir();
            dataGridView3.Columns[0].Visible = false;
            dataGridView3.Columns[1].HeaderText = "Firma Adı";
            dataGridView3.Columns[2].HeaderText = "Fatura No";
            dataGridView3.Columns[3].HeaderText = "Fatura Tarihi";
            dataGridView3.Columns[4].HeaderText = "Tutar";
            dataGridView3.Columns[5].HeaderText = "Son Ödeme Tarihi";
            dataGridView3.Columns[6].HeaderText = "Oran";
            dataGridView3.Columns[7].HeaderText = "Kdv Miktarı";
            dataGridView3.Columns[8].HeaderText = "Ödendi";
            dataGridView3.Columns[9].HeaderText = "Ödeme Tarihi";
            toplamFatura();
            tumGelir();
            if (textBox18.Text != "" && textBox17.Text != "")
            {
                double z = double.Parse(textBox18.Text) - double.Parse(textBox17.Text);
                textBox21.Text = z.ToString();
            }
        }



        private void button14_Click(object sender, EventArgs e)
        {
            if (dataGridView3.GetCellCount(DataGridViewElementStates.Selected) > 0)
            {
                string sorgu = "UPDATE fatura SET firmaAdi=@firmaAdi, faturano=@faturano, faturatarih=@faturatarih, tutar=@tutar,vade=@vade, yuzde=@yuzde, kdv_mik=@kdv_mik WHERE id=@id";
                SqlCommand komut = new SqlCommand(sorgu, baglanti);
                komut.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView3.CurrentRow.Cells[0].Value));
                komut.Parameters.AddWithValue("@firmaAdi", textBox22.Text);
                komut.Parameters.AddWithValue("@faturano", int.Parse(textBox19.Text));
                komut.Parameters.AddWithValue("@faturatarih", dateTimePicker15.Value);
                komut.Parameters.AddWithValue("@tutar", double.Parse(textBox20.Text));
                komut.Parameters.AddWithValue("@vade", dateTimePicker16.Value);
                komut.Parameters.AddWithValue("@yuzde", double.Parse(comboBox4.Text.ToString()));
                komut.Parameters.AddWithValue("@kdv_mik", double.Parse(textBox20.Text) * double.Parse(comboBox4.Text.ToString()));
                baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();
                masrafGetir2();
                renklendir();
                toplamFatura();
                tumGelir();
                if (textBox18.Text != "" && textBox17.Text != "")
                {
                    double z = double.Parse(textBox18.Text) - double.Parse(textBox17.Text);
                    textBox21.Text = z.ToString();
                }
            }
            else
            {
                MessageBox.Show("Silecek Satır Seçiniz");
            }
        }

        private void dataGridView3_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            textBox22.Text = dataGridView3.CurrentRow.Cells[1].Value.ToString();
            textBox19.Text = dataGridView3.CurrentRow.Cells[2].Value.ToString();
            dateTimePicker15.Text = dataGridView3.CurrentRow.Cells[3].Value.ToString();
            textBox20.Text = dataGridView3.CurrentRow.Cells[4].Value.ToString();
            dateTimePicker16.Text = dataGridView3.CurrentRow.Cells[5].Value.ToString();
            comboBox4.Text = dataGridView3.CurrentRow.Cells[6].Value.ToString();
        }

       
        
        void renklendir()
        {

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
               
                    fark = Convert.ToDateTime(dataGridView3.Rows[i].Cells["vade"].Value.ToString()) - Convert.ToDateTime(DateTime.Now.ToShortDateString());
                    gunfark = fark.TotalDays;

                    bool odeme = Convert.ToBoolean(dataGridView3.Rows[i].Cells["odendi"].Value);
                    
                    if (gunfark <= 2 && odeme == false)
                    {
                        dataGridView3.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    }
                    else if (gunfark >= 3 && gunfark < 7 && odeme == false)
                    {
                        dataGridView3.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                    
                    }
                    else
                    {
                        dataGridView3.Rows[i].DefaultCellStyle.BackColor = Color.White;
                    }
                
            }
            
        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            masrafGetir2();
            renklendir();
            dataGridView3.Columns[0].Visible = false;
            dataGridView3.Columns[1].HeaderText = "Firma Adı";
            dataGridView3.Columns[2].HeaderText = "Fatura No";
            dataGridView3.Columns[3].HeaderText = "Fatura Tarihi";
            dataGridView3.Columns[4].HeaderText = "Tutar";
            dataGridView3.Columns[5].HeaderText = "Son Ödeme Tarihi";
            dataGridView3.Columns[6].HeaderText = "Oran";
            toplamFatura();
            tumGelir();
            
            if (textBox18.Text != "" && textBox17.Text != "")
            {
                double z = double.Parse(textBox18.Text) - double.Parse(textBox17.Text);
                textBox21.Text = z.ToString();
            }
        }

        void toplamFatura()
        {
           
                decimal k;
                decimal m;
               baglanti.Open();

                string sql12 = "SELECT SUM(tutar) FROM fatura";
                SqlCommand komut3 = new SqlCommand(sql12, baglanti);
                if (komut3.ExecuteScalar() != DBNull.Value)
                {
                    k = Convert.ToDecimal(komut3.ExecuteScalar());
                }
                else
                {
                    k = 0;
                }



                string sql11 = "SELECT SUM(mtutar) FROM masraf";
                SqlCommand komut4 = new SqlCommand(sql11, baglanti);
                if (komut4.ExecuteScalar() != DBNull.Value)
                {
                    m = Convert.ToDecimal(komut4.ExecuteScalar());
                }
                else
                {
                    m = 0;
                }
                
                baglanti.Close();

                textBox17.Text = (m+k).ToString();

            
        }

        private void groupBox1_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void DrawGroupBox(GroupBox box, Graphics g, Color textColor, Color borderColor)
        {
            if (box != null)
            {
                Brush textBrush = new SolidBrush(textColor);
                Brush borderBrush = new SolidBrush(borderColor);
                Pen borderPen = new Pen(borderBrush);
                SizeF strSize = g.MeasureString(box.Text, box.Font);
                Rectangle rect = new Rectangle(box.ClientRectangle.X,
                                               box.ClientRectangle.Y + (int)(strSize.Height / 2),
                                               box.ClientRectangle.Width - 1,
                                               box.ClientRectangle.Height - (int)(strSize.Height / 2) - 1);
                // Clear text and border
                g.Clear(this.BackColor);
                // Draw text
                g.DrawString(box.Text, box.Font, textBrush, box.Padding.Left, 0);
                // Drawing Border
                //Left
                g.DrawLine(borderPen, rect.Location, new Point(rect.X, rect.Y + rect.Height));
                //Right
                g.DrawLine(borderPen, new Point(rect.X + rect.Width, rect.Y), new Point(rect.X + rect.Width, rect.Y + rect.Height));
                //Bottom
                g.DrawLine(borderPen, new Point(rect.X, rect.Y + rect.Height), new Point(rect.X + rect.Width, rect.Y + rect.Height));
                //Top1
                g.DrawLine(borderPen, new Point(rect.X, rect.Y), new Point(rect.X + box.Padding.Left, rect.Y));
                //Top2
                g.DrawLine(borderPen, new Point(rect.X + box.Padding.Left + (int)(strSize.Width), rect.Y), new Point(rect.X + rect.Width, rect.Y));
            }
        }

        private void groupBox2_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void groupBox3_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void groupBox4_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void groupBox5_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void groupBox8_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void groupBox6_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void groupBox7_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void groupBox9_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void groupBox10_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void groupBox11_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        private void button12_Click(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count > 0)
            {

                excel.Application app = new excel.Application();
                app.Visible = true;
                Workbook kitap = app.Workbooks.Add(System.Reflection.Missing.Value);
                Worksheet sayfa = (Worksheet)kitap.Sheets[1];
                for (int i = 0; i < dataGridView3.Columns.Count; i++)
                {
                    excel.Range alan = (excel.Range)sayfa.Cells[1, 1];
                    alan.Cells[1, i + 1] = dataGridView3.Columns[i].HeaderText;
                }
                for (int i = 0; i < dataGridView3.Columns.Count; i++)
                {
                    for (int j = 0; j < dataGridView3.Rows.Count; j++)
                    {
                        excel.Range alan2 = (excel.Range)sayfa.Cells[j + 1, i + 1];
                        alan2.Cells[2, 1] = dataGridView3[i, j].Value;
                    }
                }

                kitap.Close();

                int id;
                // Find the Excel Process Id (ath the end, you kill him
                GetWindowThreadProcessId(app.Hwnd, out id);
                System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById(id);

                app.Quit();
                excelProcess.Kill();
                                
            }
        }

        
        private void groupBox12_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (textBox23.Text != "" && comboBox5.Text != "")
            {
                string sql = "insert into zrapor (tutar,kdv, tarih) values(@tutar, @kdv, @tarih)";
                SqlCommand komut = new SqlCommand(sql, baglanti);
                baglanti.Open();
                              
                
                komut.Parameters.AddWithValue("@tutar", double.Parse(textBox23.Text));
                komut.Parameters.AddWithValue("@kdv", double.Parse(textBox23.Text) * double.Parse(comboBox5.Text.ToString()));
                komut.Parameters.AddWithValue("@tarih", dateTimePicker17.Value);

                komut.ExecuteNonQuery();
                MessageBox.Show("Tutar Eklendi", "Z-Rapor");
                baglanti.Close();

            }
            else
            {
                MessageBox.Show("Tüm Bölümün Doldurulması Zorunludur", "Z-Rapor");
                baglanti.Close();
                return;
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
                DateTime a, b;
                DateTime.TryParse(dateTimePicker18.Text, out a);
                DateTime.TryParse(dateTimePicker19.Text, out b);
                string sql = "SELECT SUM(tutar) FROM zrapor where tarih between @tarih1 and @tarih2";
                SqlCommand komut = new SqlCommand(sql, baglanti);
                baglanti.Open();
                
                komut.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = a;
                komut.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = b;
                double j;
                if (komut.ExecuteScalar() != DBNull.Value)
                {
                    j = double.Parse(komut.ExecuteScalar().ToString());
                }
                else
                {
                    j = 0;
                }
                                
                textBox24.Text=String.Format("{0:0.00}", Math.Round(j, 4).ToString());
                baglanti.Close();
                       
        }

        private void groupBox13_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (textBox25.Text != "" && textBox11.Text != "" && dateTimePicker20.Text != "" && comboBox6.Text != "")
            {
                string sql = "insert into smasraf (hkalemi,htutar,tarih,kdv_mik) values(@hkalemi,@htutar,@tarih,@kdv_mik)";
                SqlCommand komut = new SqlCommand(sql, baglanti);
                baglanti.Open();
                komut.Parameters.AddWithValue("@hkalemi", textBox25.Text);
                komut.Parameters.AddWithValue("@htutar", double.Parse(textBox11.Text));
                komut.Parameters.AddWithValue("@tarih", dateTimePicker20.Value);
                komut.Parameters.AddWithValue("@kdv_mik", double.Parse(textBox11.Text) * double.Parse(comboBox6.Text.ToString()));
                komut.ExecuteNonQuery();
                baglanti.Close();

                masrafGetir3();
            }
            else
            {
                MessageBox.Show("Tüm alanlar doldurulmalıdır");
            }
        }

        private void tabPage3_Enter(object sender, EventArgs e)
        {
            masrafGetir3();
        }

        private void dataGridView4_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            //dataGridView4.Columns[0].Visible = false;
            dataGridView4.Columns[1].HeaderText = "Harcama Kalemi";
            dataGridView4.Columns[2].HeaderText = "Harcama Tutarı";
            dataGridView4.Columns[3].HeaderText = "Kdv Miktarı";
            dataGridView4.Columns[4].HeaderText = "Tarih";
            textBox25.Text = dataGridView4.CurrentRow.Cells[1].Value.ToString();
            textBox11.Text = dataGridView4.CurrentRow.Cells[2].Value.ToString();
            dateTimePicker20.Text = dataGridView4.CurrentRow.Cells[4].Value.ToString();
            
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (dataGridView4.GetCellCount(DataGridViewElementStates.Selected) > 0)
            {
                String sorgu = "DELETE FROM smasraf WHERE id=@id";
                SqlCommand komut = new SqlCommand(sorgu, baglanti);
                komut.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView4.CurrentRow.Cells[0].Value));
                baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();
                masrafGetir3();
            }
            else
            {
                MessageBox.Show("Silinecek Alan Seçmelisiniz");
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            if (dataGridView4.GetCellCount(DataGridViewElementStates.Selected) > 0)
            {
                string sorgu = "UPDATE smasraf SET hkalemi=@hkalemi, htutar=@htutar, tarih=@tarih WHERE id=@id";
                SqlCommand komut = new SqlCommand(sorgu, baglanti);
                komut.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView4.CurrentRow.Cells[0].Value));
                komut.Parameters.AddWithValue("@hkalemi", textBox25.Text);
                komut.Parameters.AddWithValue("@htutar", double.Parse(textBox11.Text));
                komut.Parameters.AddWithValue("@tarih", dateTimePicker4.Value);
                baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();
                masrafGetir3();
            }
            else
            {
                MessageBox.Show("Güncellenecek Alan Seçmelisiniz");
            }
        }

        private void groupBox15_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void groupBox14_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }
        void datagridHeaderText2()
        {
            string[] str = { "Harcama Kalemi", "Harcama Tutarı", "Kdv Miktarı", "Tarih" };
            for (int i = 0; i < dataGridView5.Columns.Count; i++)
            {
                dataGridView5.Columns[i].HeaderText = str[i];
            }
        }
        private void button23_Click(object sender, EventArgs e)
        {
            

            DateTime a, b;
            DateTime.TryParse(dateTimePicker23.Text, out a);
            DateTime.TryParse(dateTimePicker24.Text, out b);

            string sql = "SELECT hkalemi, htutar,kdv_mik, tarih FROM smasraf WHERE tarih BETWEEN @tarih1 and @tarih2";
            DataTable dt = new DataTable();
            SqlDataAdapter daa = new SqlDataAdapter(sql, baglanti);
            baglanti.Open();

            daa.SelectCommand.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = a;
            daa.SelectCommand.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = b;
            daa.Fill(dt);
            dataGridView5.DataSource = dt;
            baglanti.Close();
            dataGridView5.AllowUserToAddRows = false;
            double sum = 0;
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                sum = sum + double.Parse(dataGridView5.Rows[i].Cells[1].Value.ToString());

            }
            textBox28.Text = sum.ToString("0.00");
            //MessageBox.Show(dataGridView5.Rows[0].Cells[2].Value.ToString());
            double h = 0;
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                h = h + double.Parse(dataGridView5.Rows[i].Cells[2].Value.ToString());

            }
            textBox32.Text = h.ToString("0.00");
            datagridHeaderText2();

        }

        private void groupBox14_Enter(object sender, EventArgs e)
        {

        }

        private void button22_Click(object sender, EventArgs e)
        {
            DateTime a, b;
            DateTime.TryParse(dateTimePicker22.Text, out a);
            DateTime.TryParse(dateTimePicker21.Text, out b);

            string sql = "SELECT hkalemi, htutar, tarih FROM smasraf WHERE tarih BETWEEN @tarih1 and @tarih2 and hkalemi=@aranan";
            DataTable dt2 = new DataTable();
            SqlDataAdapter daa = new SqlDataAdapter(sql, baglanti);
            baglanti.Open();
            daa.SelectCommand.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = a;
            daa.SelectCommand.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = b;
            daa.SelectCommand.Parameters.AddWithValue("@aranan", textBox27.Text);
            daa.Fill(dt2);
            dataGridView5.DataSource = dt2;
            baglanti.Close();
            dataGridView5.AllowUserToAddRows = false;
            double sum = 0;
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                sum = sum + double.Parse(dataGridView5.Rows[i].Cells[1].Value.ToString());

            }
            textBox26.Text = sum.ToString("0.00");
            datagridHeaderText2();
        }

        private void groupBox11_Enter(object sender, EventArgs e)
        {

        }

        private void tabPage2_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void groupBox16_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Red, Color.Red);
        }

        private void button24_Click(object sender, EventArgs e)
        {
            string sql9 = "SELECT SUM(kdv) FROM zrapor where tarih between @tarih1 and @tarih2";
            baglanti.Open();
            DateTime h, i;
            DateTime.TryParse(dateTimePicker26.Text, out h);
            DateTime.TryParse(dateTimePicker25.Text, out i);

            SqlCommand komut1 = new SqlCommand(sql9, baglanti);
            komut1.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = h;
            komut1.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = i;
            double w;
            if (komut1.ExecuteScalar() != DBNull.Value)
            {
                w = double.Parse(komut1.ExecuteScalar().ToString());
                
            }
            else
            {
                w = 0;
                
            }
            textBox31.Text = w.ToString("0.00"); ;

            string sql10 = "SELECT SUM(kdv_mik) FROM smasraf where tarih between @tarih1 and @tarih2";
            DateTime m, n;
            DateTime.TryParse(dateTimePicker26.Text, out m);
            DateTime.TryParse(dateTimePicker25.Text, out n);
            SqlCommand komut2 = new SqlCommand(sql10, baglanti);
            komut2.Parameters.AddWithValue("@tarih1", SqlDbType.Date).Value = m;
            komut2.Parameters.AddWithValue("@tarih2", SqlDbType.Date).Value = n;
            double j;
            if (komut2.ExecuteScalar() != DBNull.Value)
            {
                j = double.Parse(komut2.ExecuteScalar().ToString());
            }
            else
            {
                j = 0;
            }
            

            string sql11 = "SELECT SUM(kdv_mik) FROM fatura where faturatarih between @tarih3 and @tarih4";
            DateTime o, u;
            DateTime.TryParse(dateTimePicker26.Text, out o);
            DateTime.TryParse(dateTimePicker25.Text, out u);
            SqlCommand komut3 = new SqlCommand(sql11, baglanti);
            komut3.Parameters.AddWithValue("@tarih3", SqlDbType.Date).Value = o;
            komut3.Parameters.AddWithValue("@tarih4", SqlDbType.Date).Value = u;
            
            double y;
            if (komut3.ExecuteScalar() != DBNull.Value) { 
            y = double.Parse(komut3.ExecuteScalar().ToString());
            }
            else
            {
                y = 0;
            }
            textBox29.Text = (j + y).ToString("0.00");
            baglanti.Close();
            
                   
            textBox30.Text = (w - (y+j)).ToString("0.00");


        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (dataGridView3.GetCellCount(DataGridViewElementStates.Selected) > 0)
            {
                try
                {
                    DateTime d;
                    string sorgu = "UPDATE fatura SET odendi=@odendi, otarihi=@otarihi WHERE id=@id";
                    SqlCommand komut = new SqlCommand(sorgu, baglanti);
                    komut.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView3.CurrentRow.Cells[0].Value));
                    komut.Parameters.AddWithValue("@odendi", true);
                    DateTime.TryParse(DateTime.Now.ToShortDateString(), out d);
                    komut.Parameters.AddWithValue("@otarihi", SqlDbType.Date).Value = d;

                    baglanti.Open();
                    komut.ExecuteNonQuery();
                    baglanti.Close();
                    masrafGetir2();
                    renklendir();
                    MessageBox.Show("Fatura Ödendi");

                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            else
            {
                MessageBox.Show("Fatura Seçiniz");
            }
        }

        private void groupBox16_Enter(object sender, EventArgs e)
        {

        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar !=8)
            {
                e.Handled = true;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void textBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void textBox23_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void textBox19_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }
    }
}
