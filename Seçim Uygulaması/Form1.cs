using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel; 
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Net;

/*metrogrid5'in amacı şeydi eğer lstkod ve lstisimdeki isimler
 giderse orada kalsın diyeydi sonuçta program kapandığı zaman 
listboxtan gider ama datagridden gitmez, bu amaçla metrogrid5'i koymuştum*/
/*eğer tek başına if olarak kalsaydı daha sonra alttan try catch komutuna yine devam edecektir
ama şimdi else koyduğun için eğer veri yoksa durucak ve try içindeki işlemleri yapmaya girişmeyecek.
eğer if şartı sağlamasaydı bu sefer direkt try yapısına inmiş olacaktı ama istemediğimiz şekilde yani veri
girişi yapılmış mı yapılmamış mı bunun kontrolünü yapmadan ilerleyecekti*/
/*int toplam = 0;
int[] sayilar = new int[metroGrid7.Rows.Count];
for (int i = 0; i < metroGrid7.Rows.Count; i++)
{
    sayilar[i] = Convert.ToInt32(metroGrid7.Rows[i].Cells[1].Value.ToString());
    toplam += sayilar[i];

    int value = Convert.ToInt32(metroGrid7.Rows[i].Cells[1].Value.ToString());
    int yüzde = (value / toplam) * 100;
    metroGrid7.Rows[i].Cells[0].Value = metroGrid7.Rows[i].Cells[0].Value.ToString() + " (%" + yüzde + ")";

doğru, satırın ismi değişiyor ya o yüzden yeni satır açtırtıyor
}*/
/*//microsoft office.14.object library, sonra nugetten microsoft office interop'u ekleyip yapacağız
 */
//bunu da ayarlayalım lan, eğer iki kere excel tekrar okutulmaya çalışılırsa birden fazla oy gönderimi demektir bu, 
//ama son gönderilen her zaman geçerli olacağı için istediği kadar 2 tane göndersin manasız olacaktır çünkü sadece son gönderilen geçerli olacaktır
//ama bu durum ilk gönderilenlerin geçersiz olmasına yol açacak lan!!! o zaman access bosalt'ı kapatmak lazım ya da access taratması yapmak 
//lazım(bu daha sağlıklı gibi, böylece istenmeyen kimseye birden fazla kod gönderilemeyecek)
//şifre değiştirme olayını eklememiz lazım her iki kısma da ya, ya da belirtmek lazım şifre değişirse işler olmaz filan diye
//buldum, oylama bitti ve veriler temizlendi olsun, bunun için access bosalt lazım son kısımda
//OHA MÜKEMMEL KODMUŞ LAN BU; (metroGrid5.DataSource as DataTable).DefaultView.RowFilter = string.Format("Name LIKE '{0}%' OR Name LIKE '% {0}%'", TextBox1.Text);
//metrogrid5'te arama yaptırıyoruz ya ondan dolayı sanki orada veriler silinmiş gibi gözüküyor ama değil arama sonucu boş çıktığı için orada kimse görünmüyor yoksa içindeki veriler hala duruyor yani
//doğrulama kodu çok güzel kod valla, takdir ettim çok mantıklı ve tutarlı yazılmış bir kod gerçekten*/
namespace Seçim_Uygulaması
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        DataTable table = new DataTable();
        DataTable table2 = new DataTable();
        DataTable table3 = new DataTable();
        DataTable table4 = new DataTable();
        DataTable table5 = new DataTable();
        DataTable table6 = new DataTable();
        DataTable table8 = new DataTable();
        DataTable table9 = new DataTable();
        DataTable dt = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }
        void tpl()
        {
            if (metroGrid1.Rows.Count == 0)
            {
                metroTile2.Enabled = false;
            }
            else
            {
                metroTile2.Enabled = true;
            }

            string[] cikartilacak2 = new string[listBox3.Items.Count];
            for (int g = 0; g < listBox3.Items.Count; g++)
            {
                cikartilacak2[g] = listBox3.Items[g].ToString();
            }

            listBox3.Items.Clear();

            string[] cikartilacak = cikartilacak2.Distinct().ToArray(); //buradaki distinctte bir sorun yok çünkü tek başına çalışıyor sadece

            for (int l = 0; l < cikartilacak.Length; l++)
            {
                listBox3.Items.Add(cikartilacak[l].ToString());
            }

            int n = listBox3.Items.Count;

            if (listBox3.Items.Count == 0)
            {
                listBox3.Items.Add("Hatalı mail adresi bulunamadı");
            }
            else if (listBox3.Items.Count == 1)
            {
                listBox3.Items.Add("-----------------------");
                listBox3.Items.Add("1 adet hatalı mail adresi çıkarıldı");
            }
            else
            {
                listBox3.Items.Add("-----------------------");
                listBox3.Items.Add(n.ToString() + " adet hatalı mail adresleri çıkarıldı");
            }

            metroGrid10.ClearSelection();
            int satir = metroGrid10.Rows.Count;
            int toplam = n + satir;
            metroLabel11.Text = toplam.ToString();
            metroLabel1.Text = satir.ToString();
            metroLabel11.Visible = true;
            metroLabel1.Visible = true;

            if (metroLabel1.Text == "0")
            {
                MessageBox.Show("Yüklemiş olduğunuz excelde herhangi bir veri bulunamadı ya da veriler içinden geçerli olan herhangi bir mail adresi bulunamadı. Lütfen excel dosyanızın içinde verilerin olduğundan ve var olan mail adreslerinin doğru olduğundan emin olunuz ve ardından tekrar deneyiniz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        void Tasarim_2()
        {
            metroGrid4.Columns[0].HeaderText = "isim";
            metroGrid4.Columns[1].HeaderText = "mail";
            metroGrid4.Columns[0].Width = 137;
            metroGrid4.Columns[1].Width = 180;
        }
        void tile3check()
        {
            OleDbCommand komut2;
            string vtyolu2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=tile3.accdb;Persist Security Info=True";
            OleDbConnection baglanti2 = new OleDbConnection(vtyolu2);
            baglanti2.Open();
            string sil = "delete from tile3";
            komut2 = new OleDbCommand(sil, baglanti2);
            komut2.ExecuteNonQuery();
            komut2.Dispose();
            baglanti2.Close();

            string[] kod_gonderilen = new string[lst_kodlar.Items.Count];
            string[] isim_gonderilen = new string[lst_kodlar.Items.Count];
            string[] kod_alinan = new string[metroGrid3.Rows.Count];
            string[] isim_alinan = new string[metroGrid3.Rows.Count];
            string[] oy1 = new string[metroGrid3.Rows.Count];
            for (int i = 0; i < lst_kodlar.Items.Count; i++)
            {
                kod_gonderilen[i] = lst_kodlar.Items[i].ToString();
                isim_gonderilen[i] = lst_isimler.Items[i].ToString();
            }
            for (int h = 0; h < metroGrid3.Rows.Count; h++)
            {
                kod_alinan[h] = metroGrid3.Rows[h].Cells[1].Value.ToString();
                string bosluksuz = kod_alinan[h].Trim();
                kod_alinan[h] = bosluksuz.ToString();

                isim_alinan[h] = metroGrid3.Rows[h].Cells[0].Value.ToString();
                oy1[h] = metroGrid3.Rows[h].Cells[2].Value.ToString();
            }
            //kod alınan > kod gonderilen aslında ya da maks eşit olurlar birbirlerine

            for (int k = 0; k < lst_kodlar.Items.Count; k++)
            {
                for (int l = 0; l < metroGrid3.Rows.Count; l++)
                {
                    if (kod_alinan[l].Contains(kod_gonderilen[k].ToString()) == true)
                    {
                        lst_onaylanan_kod.Items.Add(kod_gonderilen[k].ToString());
                        lst_onaylanan_isim.Items.Add(isim_gonderilen[k].ToString());
                        lst_onaylanan_oy.Items.Add(oy1[l].ToString()); //sorun buymuş demek, k yı l diye düzeltince sorun çözüldü
                    }
                }
            }

            OleDbCommand komut;
            string vtyolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=tile3.accdb;Persist Security Info=True";
            OleDbConnection baglanti = new OleDbConnection(vtyolu);

            for (int a = 0; a < lst_onaylanan_kod.Items.Count; a++)
            {
                string isim = lst_onaylanan_isim.Items[a].ToString();
                string kod = lst_onaylanan_kod.Items[a].ToString();
                string oy = lst_onaylanan_oy.Items[a].ToString();


                baglanti.Open();
                string ekle = "insert into tile3(isim,kod,oy) values (@isim,@kod,@oy)";
                komut = new OleDbCommand(ekle, baglanti);
                komut.Parameters.AddWithValue("@isim", isim.ToString());
                komut.Parameters.AddWithValue("@kod", kod.ToString());
                komut.Parameters.AddWithValue("@oy", oy.ToString());
                komut.ExecuteNonQuery();
                komut.Dispose();
                baglanti.Close();
            }

            OleDbConnection con;
            OleDbDataAdapter da;
            DataSet ds5;
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=tile3.accdb");
            da = new OleDbDataAdapter("SElect *from tile3", con);
            ds5 = new DataSet();
            con.Open();
            da.Fill(ds5, "tile3");
            mtg_aranacak.DataSource = ds5.Tables["tile3"];
            con.Close();

            string[] duzgunkod_hepsi = new string[lst_onaylanan_kod.Items.Count];
            for (int j = 0; j < lst_onaylanan_kod.Items.Count; j++)
            {
                duzgunkod_hepsi[j] = lst_onaylanan_kod.Items[j].ToString();
            }

            string[] tekrarsiz = duzgunkod_hepsi.Distinct().ToArray(); //buradaki distinctte bir sorun yok çünkü tek başına çalışıyor sadece


            if(table3.Columns.Count == 0)
            {
                table3.Columns.Add("Ad - Soyad", typeof(string));
                table3.Columns.Add("Doğrulama Kodu", typeof(string));
                table3.Columns.Add("Oy", typeof(string));
            }

            for (int b = 0; b < tekrarsiz.Length; b++)
            {
                DataView dv = ds5.Tables["tile3"].DefaultView;
                dv.RowFilter = "kod LIKE '" + tekrarsiz[b].ToString() + "%'";
                mtg_aranacak.DataSource = dv;

                table3.Rows.Add(mtg_aranacak.Rows[0].Cells[0].Value.ToString(), mtg_aranacak.Rows[0].Cells[1].Value.ToString(), mtg_aranacak.Rows[0].Cells[2].Value.ToString());
            }
            metroGrid9.DataSource = table3;
            metroGrid9.Columns[0].Width = 183;
            metroGrid9.Columns[1].Width = 148;
            metroGrid9.Columns[2].Width = 104;

            string[] cikar2 = new string[metroGrid3.Rows.Count];
            for (int t = 0; t < metroGrid3.Rows.Count; t++)
            {
                cikar2[t] = metroGrid3.Rows[t].Cells[1].Value.ToString();
            }
            string[] cikar = cikar2.Distinct().ToArray(); //buradaki distinctte bir sorun yok çünkü tek başına çalışıyor sadece

            for (int i = 0; i < metroGrid9.Rows.Count; i++)
            {
                listBox6.Items.Add(metroGrid9.Rows[i].Cells[1].Value.ToString());
            }

            for (int a = 0; a < cikar.Length; a++)
            {
                if (listBox6.Items.Contains(cikar[a]) == false)
                {
                    listBox4.Items.Add(cikar[a]);
                }
            }

            int n = listBox4.Items.Count;

            if (listBox4.Items.Count == 0)
            {
                listBox4.Items.Add("Hatalı mail adresi bulunamadı");
            }
            else if (listBox4.Items.Count == 1)
            {
                listBox4.Items.Add("-----------------------");
                listBox4.Items.Add("1 adet hatalı mail adresi çıkarıldı");
            }
            else
            {
                listBox4.Items.Add("-----------------------");
                listBox4.Items.Add(n.ToString() + " adet hatalı mail adresleri çıkarıldı");
            }

            int satir = metroGrid9.Rows.Count;
            metroLabel10.Text = satir.ToString();
            int toplam = n + satir;
            metroLabel6.Text = toplam.ToString();

            metroLabel6.Visible = false;
            metroLabel10.Visible = false;

            if (metroLabel10.Text == "0")
            {
                MessageBox.Show("Yüklemiş olduğunuz excelde herhangi bir veri bulunamadı ya da veriler içinden geçerli olan herhangi bir doğrulama kodu bulunamadı. Lütfen excel dosyanızın içinde verilerin olduğundan ve var olan doğrulama kodlarının doğru olduğundan emin olunuz ve ardından tekrar deneyiniz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                metroTile4.Enabled = false;
            }
            else
            {
                metroTile4.Enabled = true;
            }
        }
        void check1()
        {
            OleDbCommand komut;
            OleDbCommand komut2;
            string vtyolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=veri.accdb;Persist Security Info=True";
            OleDbConnection baglanti = new OleDbConnection(vtyolu);

            for (int a = 0; a < metroGrid4.Rows.Count; a++)
            {
                string isim = metroGrid4.Rows[a].Cells[0].Value.ToString();
                string e_mail = metroGrid4.Rows[a].Cells[1].Value.ToString();


                baglanti.Open();
                string ekle = "insert into veri(isim,mail) values (@isim,@mail)";
                komut = new OleDbCommand(ekle, baglanti);
                komut.Parameters.AddWithValue("@isim", isim.ToString());
                komut.Parameters.AddWithValue("@mail", e_mail.ToString());
                komut.ExecuteNonQuery();
                komut.Dispose();
                baglanti.Close();
                //tamam şu an veriler accesse eklendi artık, silinmesi gerekenler hala orada duruyor
            }

            for (int b = 0; b < listBox3.Items.Count; b++)
            {
                string h = listBox3.Items[b].ToString();

                baglanti.Open();
                string sil = "delete from veri where mail=@mail";
                komut2 = new OleDbCommand(sil, baglanti);
                komut2.Parameters.AddWithValue("@mail", h);
                komut2.ExecuteNonQuery();
                komut2.Dispose();
                baglanti.Close();
                //tamam şu an veriler silindi 
            }

            string pathconn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=veri.accdb;Persist Security Info=True";
            OleDbConnection conn = new OleDbConnection(pathconn);
            OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from veri", conn);
            DataTable dt2 = new DataTable();
            MyDataAdapter.Fill(dt2);
            metroGrid4.DataSource = dt2;
            Tasarim_2();

            baglanti.Open();
            string sil2 = "delete from veri";
            komut2 = new OleDbCommand(sil2, baglanti);
            komut2.ExecuteNonQuery();
            komut2.Dispose();
            baglanti.Close();
        }
        void check2()
        {
            /*datasource larda ekleme silme işlemini her zaman datatable üstünden yürütmek gerekiyormuş demek, enteresan ya
            orada aslında kodların çalışmama sebebini ben yanlış yorumladım alışkın olmadığım için çünkü o as datatable
            komutu sayesinde metrogridde eleman var olmasına rağmen aramadan kaynaklı olarak hafızada kalıyor ve görmüyor
            listboxları devreye sokunca anca düzeldi sonra kod*/

            for (int i = 0; i < metroGrid4.Rows.Count; i++)
            {
                listBox1.Items.Add(metroGrid4.Rows[i].Cells[0].Value.ToString());
                listBox2.Items.Add(metroGrid4.Rows[i].Cells[1].Value.ToString());
            }

            int x = metroGrid4.Rows.Count; 

            //tabi ya, arama yaptığı için tek bir tane satır kalıyor metrogridde sonra programda datagridde tek satır kaldığını düşünerek ve tek satır ilettiği 
            //için zaten işleyişi durduruyor; bundan dolayı bu rows.count'ı for içine yazmak yerine önce bir değerini almak lazım ardından değişkenin sahip olduğu 
            //değer kadar döndürmesini istemeyeliyiz döngüyü aramadan kaynaklı olarak datagrid'in satırının azalmasından bağımsız olarak
            for (int i = 0; i < x; i++)
            {
                //vay amq ulan çok saçma ama bu durum; abi yetersiz kod yazıyorsunuz sonra ben uğraşıyorum. Ne demek lan datagridi yenilememek
                //özellikle altta yenile demişim aq ben ne yapayım abi ya Allah aşkına daha
                //ulan her yerde hata olabilir diye test ettim aq en aklıma gelmeyen yerden çıktı hata iyi mi
                string a = listBox1.Items[i].ToString();
                string b = listBox2.Items[i].ToString();
                (metroGrid4.DataSource as DataTable).DefaultView.RowFilter = string.Format("mail LIKE '{0}%'", b);

                if (metroGrid4.Rows.Count == 1) //
                {
                    table5.Rows.Add(a, b);
                }

                //bu satır önemli bak dursun, demek datasource olanlara ekleme silme işlemlerini böyle yapıyormuşuz
                /*for (int i = 0; i < length; i++)
                {
                    metroGrid4.Rows.RemoveAt(i);
                    metroGrid4.Refresh();
                }*/

                table6.Rows.Clear();
                for (int g = 0; g < x; g++)
                {
                    DataRow dr = table6.NewRow();
                    dr["isim"] = listBox1.Items[g].ToString();
                    dr["mail"] = listBox2.Items[g].ToString();
                    table6.Rows.Add(dr);
                }
                metroGrid4.DataSource = table6;
            }
            metroGrid11.DataSource = table5;
            metroGrid11.Columns[0].Width = 137;
            metroGrid11.Columns[1].Width = 180;


            //treli name'i olan sütunlarda arama yaptırmadı nedense

            /*string[] mailkontrol = mail.Distinct().ToArray();
            string[] isimkontrol = isim.Distinct().ToArray();
            ulan distinctlerde böyle bir sorun yaşayacağım hiç aklıma gelmezdi lan ama aslında çok mantıklı bir hata lan
            aynı isimler ama farklı ve geçerli mail adresleri olduğu için adresler düzgünce ekleniyor ama isimler eksik kalıyor sonra
            sonra program hata vermeden çökmüş gibi oluyor
            tüm distinctleri kontrol et bi bakalım böyle bir sorun olma potansiyeli var mı acaba
            demek verileri o zaman tablelarla çözecek yine işte aynısı var mı yok mu şeklinde ve isimleri öyle ekleteceğiz, daha sağlıklı oluyor
            distinct güzel metot aslında ama daha bugsuz bir yol buldum sanırsam bu tablelar ile*/
            metroGrid1.ColumnCount = 2;

            metroGrid1.Columns[0].Name = "İsim ve Soyisim";
            metroGrid1.Columns[0].HeaderText = "İsim ve Soyisim";
            metroGrid1.Columns[0].Width = 137;

            metroGrid1.Columns[1].Name = "E-mail";
            metroGrid1.Columns[1].HeaderText = "E-mail";
            metroGrid1.Columns[1].Width = 180;

            for (int i = 0; i < metroGrid11.Rows.Count; i++)
            {
                int idx = metroGrid1.Rows.Add();
                metroGrid1.Rows[idx].Cells[0].Value = metroGrid11.Rows[i].Cells[0].Value.ToString();
                metroGrid1.Rows[idx].Cells[1].Value = metroGrid11.Rows[i].Cells[1].Value.ToString(); //value yazmadan neden cell'i bi tuhaf alıyor ki acaba ilginç gerçekten, nasıl bir amacı varsa artık c#'ı yazan kişiler için
            }
        }
        void check3()
        {
            //evet, mail adreslerini boşluksuz aldırmak lazım metrogrid1'e, çok fark ettiriyor o; ben öyle almıyormuydum bunları aslında ya
            try
            {
                for (int i = 0; i < metroGrid1.Rows.Count; i++)
                {
                    string a = metroGrid1.Rows[i].Cells[0].Value.ToString();
                    string b = metroGrid1.Rows[i].Cells[1].Value.ToString();
                    (metroGrid5.DataSource as DataTable).DefaultView.RowFilter = string.Format("mail LIKE '{0}%'", b);

                    if (metroGrid5.Rows.Count == 0)
                    {
                        table4.Rows.Add(a, b);
                    }
                }

                metroGrid10.DataSource = table4;
                metroGrid10.Columns[0].Width = 137;
                metroGrid10.Columns[1].Width = 180;
            }
            catch (Exception)
            {
                return;
            }
        }
        void accessekle()
        {
            OleDbCommand komut;
            string vtyolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=kod.accdb;Persist Security Info=True";
            OleDbConnection baglanti = new OleDbConnection(vtyolu);

            for (int a = 0; a < lst_kodlar.Items.Count; a++)
            {
                string isim = lst_isimler.Items[a].ToString();
                string mail = lst_mailler.Items[a].ToString();
                string kod = lst_kodlar.Items[a].ToString();

                baglanti.Open();
                string ekle = "insert into kod(isim,mail,kod) values (@isim,@mail,@kod)";
                komut = new OleDbCommand(ekle, baglanti);
                komut.Parameters.AddWithValue("@isim", isim.ToString());
                komut.Parameters.AddWithValue("@mail", mail.ToString());
                komut.Parameters.AddWithValue("@kod", kod.ToString());
                komut.ExecuteNonQuery();
                komut.Dispose();
                baglanti.Close();
            }
        }
        void griddoldur()
        {
            OleDbConnection con;
            OleDbDataAdapter da;
            DataSet ds;
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=kod.accdb");
            da = new OleDbDataAdapter("SElect *from kod", con);
            ds = new DataSet();
            con.Open();
            da.Fill(ds, "kod");
            metroGrid5.DataSource = ds.Tables["kod"];
            con.Close();
        }
        void koduret()
        {
            string GuvenlikKodu;
            GuvenlikKodu = "";
            int harf, bykharf, hangisi;
            Random Rharf = new Random();
            Random Rsayi = new Random();
            Random Rbykharf = new Random();
            Random Rhangisi = new Random();

            for (int b = 0; b < 6; b++)
            {
                int a = 0;
                hangisi = Rhangisi.Next(1, 3);
                if (hangisi == 1)
                {
                    GuvenlikKodu += Rsayi.Next(0, 10).ToString();
                }
                if (hangisi == 2)
                {
                    harf = Rharf.Next(1, 27);
                    for (char i = 'a'; i <= 'z'; i++)
                    {
                        a++;
                        if (a == harf)
                        {
                            bykharf = Rbykharf.Next(1, 3);
                            if (bykharf == 1)
                            {
                                GuvenlikKodu += i;
                            }
                            if (bykharf == 2)
                            {
                                GuvenlikKodu += i.ToString().ToUpper();
                            }
                        }
                    }
                }

            }

            txt_kod.Text = GuvenlikKodu;
            lst_kodlar.Items.Add(txt_kod.Text.ToString());
        }
        void dogrulamakodugonder()
        {
            try
            {
                SmtpClient sc = new SmtpClient
                {
                    Port = 587,
                    Timeout = 3600000,
                    Host = "smtp.gmail.com",
                    EnableSsl = true,
                    Credentials = new NetworkCredential(Settings.Default.mail,Settings.Default.sifre)
                };
                MailMessage mail = new MailMessage
                {
                    From = new MailAddress(Settings.Default.mail, "Hacettepe Üniversitesi Psikoloji Topluluğu (HÜPT)")
                };
                mail.To.Add(txt_email.Text.ToString());
                mail.Subject = "HÜPT - Oy Doğrulama Kodu"; mail.IsBodyHtml = true; mail.Body = "Merhaba sevgili " + txt_isim.Text.ToString() + "," + " <br/> <br/> " + "HÜPT seçimlerinde oy kullanabilmeniz için gereken doğrulama kodu: " + txt_kod.Text + " <br/> <br/> " + "Oy kullanacağınız için çok teşekkür ederiz.";
                sc.Send(mail);
                lst_isimler.Items.Add(txt_isim.Text.ToString());
                lst_mailler.Items.Add(txt_email.Text.ToString());
            }
            catch
            {
                return;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            table4.Columns.Add("İsim ve Soyisim", typeof(string));
            table4.Columns.Add("E-mail", typeof(string));

            table5.Columns.Add("isim", typeof(string));
            table5.Columns.Add("mail", typeof(string));

            table6.Columns.Add("isim", typeof(string));
            table6.Columns.Add("mail", typeof(string));
            //evet bu ikisini buraya koyunca düzeldi çünkü ben void3'e her tıkladığımda yeni sütun ekliyordu ve bu yüzden de
            //yeni satır eklemesini sağlatamadık ondan hata veriyordu program, artık hata vermiyor şu an

            griddoldur();
            if(metroGrid5.Rows.Count != 0)
            {
                for (int i = 0; i < metroGrid5.Rows.Count; i++)
                {
                    lst_isimler.Items.Add(metroGrid5.Rows[i].Cells[0].Value.ToString());
                    lst_mailler.Items.Add(metroGrid5.Rows[i].Cells[1].Value.ToString());
                    lst_kodlar.Items.Add(metroGrid5.Rows[i].Cells[2].Value.ToString());
                }
            }

            metroGrid9.Visible = false;
            metroLink1.Visible = false;
            listBox4.Visible = false;
            metroLabel6.Visible = false;
            metroLabel10.Visible = false;
            metroLabel5.Visible = false;
            metroLabel7.Visible = false;
        }
        private void metroTile1_Click(object sender, EventArgs e)
        {
            //ulan bende diyorum program neden mal gibi kalıyor,amq return yapmışım program ne yapsın bana hatayı gösteremiyor
            try
            {
                OpenFileDialog openfile1 = new OpenFileDialog
                {
                    Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
                    Title = "Veri Excel'ini seçiniz..."
                };
                if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.textBox1.Text = openfile1.FileName;
                }

                Excel.Application oXL = new Excel.Application(); //hmm demek nuget paketten bulmak gerekiyormuş seni ve sonrada öyle using Excel diyerek kullanmak gerekiyormuş
                if (textBox1.Text == string.Empty)
                {
                    return;
                }
                else
                {
                    listBox1.Items.Clear();
                    listBox2.Items.Clear();
                    metroGrid1.Rows.Clear();
                    table4.Rows.Clear();
                    table5.Rows.Clear();
                    table6.Rows.Clear();
                    /*bu niye aklıma gelmedi ki benim, o zaman buraya sıfırlama lazım tıpkı ötekiler gibi demek; 
                    evet tahmin ettiğim gibi düzeldi de gerçekten, sorun buradaymış demek ki*/

                    Excel.Workbook oWB = oXL.Workbooks.Open(textBox1.Text); // hata burada oluşuyor demek

                    List<string> liste = new List<string>();
                    foreach (Excel.Worksheet oSheet in oWB.Worksheets)
                    {
                        liste.Add(oSheet.Name);
                    }
                    oWB.Close();
                    oXL.Quit();
                    oWB = null;
                    oXL = null;
                    metroGrid2.DataSource = liste.Select(x => new { SayfaAdi = x }).ToList();
                    textBox2.Text = metroGrid2.Rows[0].Cells[0].Value.ToString();

                    OleDbCommand komut = new OleDbCommand();
                    string pathconn = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + textBox1.Text + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
                    OleDbConnection conn = new OleDbConnection(pathconn);
                    OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [" + textBox2.Text + "$]", conn);
                    DataTable dt3 = new DataTable();
                    MyDataAdapter.Fill(dt3);
                    metroGrid4.DataSource = dt3;
                    Tasarim_2();

                    Regex regex = new Regex(@"^([a-zA-Z0-9_\-\.]+)@hacettepe.edu.tr$");

                    listBox3.Items.Clear();
                    //şu mavi grid silinemediği için tekrarlanıyor

                    string[] email = new string[metroGrid4.Rows.Count];
                    for (int i = 0; i < metroGrid4.Rows.Count; i++)
                    {
                        email[i] = metroGrid4.Rows[i].Cells[1].Value.ToString();
                        string y = email[i].Trim(); //iyi oldu bu ya güzel işe yarıyor
                        email[i] = y.ToString();

                        metroGrid4.Rows[i].Cells[1].Value = email[i];

                        bool İsValidEmail = regex.IsMatch(email[i]);

                        if (!İsValidEmail)
                        {
                            listBox3.Items.Add(email[i].ToString());
                        }
                    }

                    metroGrid1.BorderStyle = System.Windows.Forms.BorderStyle.None;
                    metroGrid10.BorderStyle = System.Windows.Forms.BorderStyle.None;

                    if (metroLink3.Text == "Hatalı Mail Adreslerini Gizle")
                    {
                        listBox3.Visible = true;
                    }
                    else if (metroLink3.Text == "Hatalı Mail Adreslerini Göster")
                    {
                        listBox3.Visible = false;
                    }
                }

                check1(); //hatalı mail çıkarımı
                check2(); //ismi iki kere yazanların çıkarımı
                check3(); //önceden kod almış olanların çıkarımı
                tpl();

                textBox1.Clear();
                textBox2.Clear();
                lst_isimler.Items.Clear();
                lst_mailler.Items.Clear();
                lst_kodlar.Items.Clear();
                metroTile3.Enabled = false;
                griddoldur();
                metroGrid10.ClearSelection();
            }
            catch (Exception)
            {
                return;
            }
        }
        private void metroTile2_Click(object sender, EventArgs e)
        {
            DialogResult dialog;
            dialog = MessageBox.Show("Doğrulama kodlarını toplu bir şekilde e-mail göndermek istediğinize emin misiniz?", "Doğrulama Kodu Gönderimi Kontrolü", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialog == DialogResult.Yes)
            {
                metroLabel8.Visible = true;
                metroProgressBar1.Visible = true;
                metroLabel9.Visible = true;

                lst_isimler.Items.Clear();
                lst_mailler.Items.Clear();
                lst_kodlar.Items.Clear();

                for (int i = 0; i <= metroGrid10.Rows.Count; i++)
                {
                    if (i == metroGrid10.Rows.Count)
                    {
                        decimal yuzde = ((decimal)(i) / (decimal)metroGrid10.Rows.Count) * 100;
                        Application.DoEvents();
                        metroProgressBar1.Value = (int)yuzde;
                        yuzde = Math.Round(yuzde, 2);
                        metroLabel8.Text = "%" + yuzde.ToString();
                    }
                    else
                    {
                        txt_isim.Text = metroGrid10.Rows[i].Cells[0].Value.ToString();
                        txt_email.Text = metroGrid10.Rows[i].Cells[1].Value.ToString();

                        koduret();
                        dogrulamakodugonder();

                        decimal yuzde = ((decimal)(i + 1) / (decimal)metroGrid10.Rows.Count) * 100;
                        Application.DoEvents();
                        metroProgressBar1.Value = (int)yuzde;
                        yuzde = Math.Round(yuzde, 2);
                        metroLabel8.Text = "%" + yuzde.ToString();
                    }
                }

                metroLabel9.Visible = false;

                Form_Loading x = new Form_Loading();
                x.Show();

                accessekle();
                griddoldur();
                metroTile3.Enabled = true;
            }

            else
            {
                MessageBox.Show("Email gönderme işlemi başlatılmadı.", "Başlatılmayan İşlem", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void metroTile3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile1 = new OpenFileDialog
            {
                Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
                Title = "Veri Excel'ini seçiniz..."
            };
            if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox1.Text = openfile1.FileName;
            }

            Excel.Application oXL = new Excel.Application(); //hmm demek nuget paketten bulmak gerekiyormuş seni ve sonrada öyle using Excel diyerek kullanmak gerekiyormuş
            if (textBox1.Text == string.Empty)
            {
                return;
            }
            else
            {
                dt.Rows.Clear();
                dt.Columns.Clear();
                //tam tahmin ettiğim gibi, dv bir excel kaynağı olduğu için kaynağın sütunu belirli bir şey değil
                //sen buna bir kere sütun tanıttıktan sonra hep onu sütunu olarak tanıyor ve diğer sütunlar eklenmeye çalışıldığı zaman basıyor hatayı
                //üsttekinde neden bu durum yaşanmamış ki acaba enteresan, ondada mı yoksa sütunu filan sildim ya da o grid'in sütun değerleri belli ve sabit miydi
                //sanki sabit sütun değerleri vardı gibi geliyor bana sanırsam ondan dolayı bir sorun yaşanmadı üstte ama bunda sabit sütun değeri olmayınca böyle bir sıkıntı
                //yaşandı işte, yaşanacağı varmış demek yaşandı ve bitti :D
                //hakkaten böyleymiş lan, sabit sütun değeri ayırabilmek önemli bir şeymiş demek yoksa sonra bu şekilde sıkıntılar yaşayabiliyorsun
                //bu yüzden ya sabit sütun değerleri kullanacaksın ya da sütunları her seferinde silip yeniden yükleyeceksin
                //artık hangisi işine gelirse daha çok ona göre kullanırsın, bu bilgiyi öğrendiğim iyi oldu
                table3.Rows.Clear();
                table8.Rows.Clear();
                table9.Rows.Clear();

                Excel.Workbook oWB = oXL.Workbooks.Open(textBox1.Text); // hata burada oluşuyor demek

                List<string> liste = new List<string>();
                foreach (Excel.Worksheet oSheet in oWB.Worksheets)
                {
                    liste.Add(oSheet.Name);
                }
                oWB.Close();
                oXL.Quit();
                oWB = null;
                oXL = null;
                metroGrid2.DataSource = liste.Select(x => new { SayfaAdi = x }).ToList();
                textBox2.Text = metroGrid2.Rows[0].Cells[0].Value.ToString();

                OleDbCommand komut = new OleDbCommand();
                string pathconn = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + textBox1.Text + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
                OleDbConnection conn = new OleDbConnection(pathconn);
                OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [" + textBox2.Text + "$]", conn);
                MyDataAdapter.Fill(dt);
                metroGrid3.DataSource = dt;
            }
            if (metroLink1.Text == "Hatalı Doğrulama Kodlarını Gizle")
            {
                listBox4.Visible = true;
            }
            else if (metroLink1.Text == "Hatalı Doğrulama Kodlarını Göster")
            {
                listBox4.Visible = false;
            }

            metroGrid9.BorderStyle = System.Windows.Forms.BorderStyle.None;

            listBox4.Items.Clear();
            tile3check();
            textBox1.Clear();
            textBox2.Clear();
            lst_onaylanan_isim.Items.Clear();
            lst_onaylanan_kod.Items.Clear();
            lst_onaylanan_oy.Items.Clear();
            listBox6.Items.Clear();
            metroTile1.Enabled = false;
            metroTile2.Enabled = false;

            if(table8.Columns.Count == 0)
            {
                table8.Columns.Add("Sayfa Adı", typeof(string));
            }

            for (int i = 0; i < metroGrid2.Rows.Count; i++)
            {
                table8.Rows.Add(metroGrid2.Rows[i].Cells[0].Value.ToString());
            }

            metroGrid2.DataSource = table8;

            if(table9.Columns.Count == 0)
            {
                table9.Columns.Add("isim", typeof(string));
                table9.Columns.Add("mail", typeof(string));
                table9.Columns.Add("kod", typeof(string));
            }

            for (int i = 0; i < mtg_aranacak.Rows.Count; i++)
            {
                table9.Rows.Add(mtg_aranacak.Rows[i].Cells[0].Value.ToString(), mtg_aranacak.Rows[i].Cells[1].Value.ToString(), mtg_aranacak.Rows[i].Cells[2].Value.ToString());
            }

            mtg_aranacak.DataSource = table9;
            metroGrid9.ClearSelection();

            MessageBox.Show("Oylar Başarılı Bir Şekilde Sisteme Alınmıştır. Oy Sayım İşlemine Geçebiliriz", "Oylama Vakti", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void metroTile4_Click(object sender, EventArgs e)
        {
            metroGrid6.ColumnCount = 1;

            metroGrid6.Columns[0].Name = "Oylar";
            metroGrid6.Columns[0].HeaderText = "Oylar";
            metroGrid6.Columns[0].Width = 450;

            for (int i = 0; i < metroGrid9.Rows.Count; i++)
            {
                int idx = metroGrid6.Rows.Add();
                metroGrid6.Rows[idx].Cells[0].Value = "*";
            }

            metroGrid6.ClearSelection();

            metroGrid6.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            metroGrid6.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            if(table.Columns.Count == 0)
            {
                table.Columns.Add("Kişiler", typeof(string));
                table.Columns.Add("Oylari", typeof(string));
            }

            for (int i = 0; i < metroGrid9.Rows.Count; i++)
            {
                table.Rows.Add(metroGrid9.Rows[i].Cells[0].Value.ToString(), metroGrid9.Rows[i].Cells[2].Value.ToString());
            }

            metroGrid8.DataSource = table;

            if(table2.Columns.Count == 0)
            {
                table2.Columns.Add("Adaylar", typeof(string));
                table2.Columns.Add("Oy Sayıları", typeof(int));
            }
            metroGrid7.DataSource = table2;
            metroGrid7.Columns[0].Width = 156;
            metroGrid7.Columns[1].Width = 83;

            metroGrid7.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            metroGrid7.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;

            metroGrid7.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            metroGrid7.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            metroTile1.Enabled = false;
            metroTile2.Enabled = false;
            metroTile3.Enabled = false;
            metroTile4.Enabled = false;
            //buradan itibaren geri dönüş yok artık oy sayılacak yani abi lütfen, kaç oy olacağı da altta gözüküyor zaten

            metroLabel14.Text = metroGrid6.Rows.Count.ToString();
            metroLabel2.Visible = true;
            metroLabel14.Visible = true;

            metroGrid6.Visible = true;
            metroGrid6.ClearSelection();
        }
        private void metroGrid6_Click(object sender, EventArgs e)
        {
            try
            {
                int x = Convert.ToInt32(metroGrid6.CurrentRow.Index.ToString());
                metroGrid6.Rows[x].Cells[0].Value = metroGrid9.Rows[x].Cells[2].Value.ToString();
                if (lst_tıklatmama.Items.Contains(x.ToString()) == true)
                {
                    //belki unclickable yapamam seni ama unclickable olamadığına pişman edecek, onun o özelliğini fark türlerde deşfire edecek 
                    //bir şeyler yazabilirim sevgili cell'ciğim :))
                    //seni arkaplanda işaretlenmiş olarak işaretlerim ve böylece sana istendiği kadar "fiziksel" olarak tıklansın, 
                    //"zihinsel" olarak sana hiç bir şey tıklanamayacak hale getiririm :))
                }
                else
                {
                    DataView dv = table.DefaultView;
                    dv.RowFilter = "Oylari LIKE '" + metroGrid6.Rows[x].Cells[0].Value.ToString() + "%'";
                    metroGrid8.DataSource = dv;

                    listBox5.Items.Clear();
                    for (int i = 0; i < metroGrid7.Rows.Count; i++)
                    {
                        listBox5.Items.Add(metroGrid7.Rows[i].Cells[0].Value.ToString()); ;
                    }

                    if (listBox5.Items.Contains(metroGrid8.Rows[0].Cells[1].Value.ToString()) == false)
                    {
                        string y = "1";
                        int ceviri = Convert.ToInt32(y);
                        table2.Rows.Add(metroGrid6.Rows[x].Cells[0].Value.ToString(), ceviri);
                    }
                    else
                    {
                        for (int i = 0; i < metroGrid7.Rows.Count; i++)
                        {
                            if (metroGrid8.Rows[0].Cells[1].Value.ToString() == metroGrid7.Rows[i].Cells[0].Value.ToString())
                            {
                                int arttirilacak = Convert.ToInt32(metroGrid7.Rows[i].Cells[1].Value.ToString());
                                arttirilacak++;
                                metroGrid7.Rows[i].Cells[1].Value = arttirilacak;
                            }
                        }
                    }

                    metroGrid7.Sort(metroGrid7.Columns[1], ListSortDirection.Descending);
                    metroGrid7.ClearSelection();
                    lst_tıklatmama.Items.Add(x.ToString());

                    if (lst_tıklatmama.Items.Count == metroGrid6.Rows.Count)
                    {
                        MessageBox.Show("Oy sayım işlemi tamamlanmıştır. Kazanan aday/adaylarımızı tebrik ederiz.", "Oy Sayımı Bitişi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        btn_sifirla.Visible = true;
                    }

                }
                metroGrid7.ClearSelection();
            }
            catch
            {
                return;
            }
            //buldum, eğer yeni satırlar ekletirsek alta ------ lı yerlerin ardından, yüzdeli oy sonuçlarına yer veririz, ama ne derece gerekli olur ki bu?
            /*oha, boşluğa tıkladın mı en üsttekine oy gidiyor, çünkü null ya onlar satır "0" sayılıyor ve üstteki oyu açıyor, 
            yapacak bir şey yok buna en azından her satıra br kere tıklatmayı kontrol edebildiğim için sorun olmuyor*/
            /*ulan program haklı tabi amq, to string ile bitirdiğin değere nasıl aktarma yapacaksın, 
             zaten tostring diyerek değer bu demişsin sen, onun üstüne daha nasıl bir oynama yapmayı planlıyorsun acaba sevgili tahir?*/
        }
        private void metroLink1_Click(object sender, EventArgs e)
        {
            //Güzel sistem oldu
            //haha ulan listbox açıkken eklendiği zaman "hatalı kod bulunamadı" yazısını da eleman olarak kabul ediyor lan eğer açık kalıyorsa, kapalı kalıyorsa etmiyor adsgsd
            if (lst4_kutucuk.Items.Count == 0)
            {
                int butonabasissayisi = 0;
                butonabasissayisi++;
                lst4_kutucuk.Items.Add(butonabasissayisi.ToString());

                if (butonabasissayisi % 2 == 0)
                {
                    metroLink1.Text = "Hatalı Doğrulama Kodlarını Göster";
                    listBox4.Visible = false;
                }
                else
                {
                    metroLink1.Text = "Hatalı Doğrulama Kodlarını Gizle";
                    listBox4.Visible = true;
                    if (listBox4.Items.Count == 0)
                    {
                        listBox4.Items.Add("Hatalı doğrulama kodu bulunamadı");
                    }
                }
            }
            else
            {
                int yeni = Convert.ToInt32(lst4_kutucuk.Items.Count.ToString());
                yeni++;
                lst4_kutucuk.Items.Add(yeni.ToString());

                if(yeni %2 == 0)
                {
                    metroLink1.Text = "Hatalı Doğrulama Kodlarını Göster";
                    listBox4.Visible = false;
                }
                else
                {
                    metroLink1.Text = "Hatalı Doğrulama Kodlarını Gizle";
                    listBox4.Visible = true;
                    if (listBox4.Items.Count == 0)
                    {
                        listBox4.Items.Add("Hatalı doğrulama kodu bulunamadı");
                    }
                }
            }
        }
        private void metroLink3_Click(object sender, EventArgs e)
        {
            if (lst3_kutucuk.Items.Count == 0)
            {
                int butonabasissayisi = 0;
                butonabasissayisi++;
                lst3_kutucuk.Items.Add(butonabasissayisi.ToString());

                if (butonabasissayisi % 2 == 0)
                {
                    metroLink3.Text = "Hatalı Mail Adreslerini Göster";
                    listBox3.Visible = false;
                }
                else
                {
                    metroLink3.Text = "Hatalı Mail Adreslerini Gizle";
                    listBox3.Visible = true;
                    if (listBox3.Items.Count == 0)
                    {
                        listBox3.Items.Add("Hatalı mail adresleri bulunamadı");
                    }
                }
            }
            else
            {
                int yeni = Convert.ToInt32(lst3_kutucuk.Items.Count.ToString());
                yeni++;
                lst3_kutucuk.Items.Add(yeni.ToString());

                if (yeni % 2 == 0)
                {
                    metroLink3.Text = "Hatalı Mail Adreslerini Göster";
                    listBox3.Visible = false;
                }
                else
                {
                    metroLink3.Text = "Hatalı Mail Adreslerini Gizle";
                    listBox3.Visible = true;
                    if (listBox3.Items.Count == 0)
                    {
                        listBox3.Items.Add("Hatalı mail adresleri bulunamadı");
                    }
                }
            }
        }
        private void metroLink2_Click(object sender, EventArgs e)
        {
            frm_bilgi x = new frm_bilgi();
            x.Show();
        }
        private void btn_sifirla_Click(object sender, EventArgs e)
        {
            table.Rows.Clear();
            table2.Rows.Clear();
            table3.Rows.Clear();
            table4.Rows.Clear();
            table5.Rows.Clear();
            table6.Rows.Clear();
            table8.Rows.Clear();
            table9.Rows.Clear();
            dt.Rows.Clear();
            metroGrid1.Rows.Clear();
            metroGrid6.Rows.Clear();

            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox3.Items.Clear();
            listBox4.Items.Clear();
            listBox5.Items.Clear();
            listBox6.Items.Clear();
            listBox7.Items.Clear();
            lst_onaylanan_isim.Items.Clear();
            lst_onaylanan_kod.Items.Clear();
            lst_onaylanan_oy.Items.Clear();
            lst_onaylanmayan_isim.Items.Clear();
            lst_onaylanmayan_kod.Items.Clear();
            lst_onaylanmayan_oy.Items.Clear();
            lst_tıklatmama.Items.Clear();
            lst3_kutucuk.Items.Clear();
            lst4_kutucuk.Items.Clear();

            metroLabel2.Visible = false;
            metroLabel6.Text = "0";
            metroLabel6.Visible = false;
            metroLabel10.Text = "0";
            metroLabel10.Visible = false;
            metroLabel14.Text = "0";
            metroLabel14.Visible = false;

            metroTile1.Enabled = true;
            metroTile3.Enabled = true;
            btn_sifirla.Visible = false;
            metroGrid6.Visible = false;

            MessageBox.Show("Yeni Oy Sayımı için program hazır.","Bilgilendirme", MessageBoxButtons.OK,MessageBoxIcon.Information);
        }
        private void metroLink4_Click(object sender, EventArgs e)
        {
            sifre x = new sifre();
            x.Show();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            metroLink1.Visible = true;
            listBox4.Visible = true;
            metroLabel6.Visible = true;
            metroLabel10.Visible = true;
            metroLabel5.Visible = true;
            metroLabel7.Visible = true;
            metroGrid9.Visible = true;
            metroLabel19.Visible = false;
            metroButton1.Visible = false;
        }
    }
}
