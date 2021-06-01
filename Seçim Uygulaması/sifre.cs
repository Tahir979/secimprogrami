using System;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;

namespace Seçim_Uygulaması
{
    public partial class sifre : MetroFramework.Forms.MetroForm
    {
        public sifre()
        {
            InitializeComponent();
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

            txt_gonderilenkod.Text = GuvenlikKodu;
        }
        void kodgonder()
        {
            SmtpClient sc = new SmtpClient
            {
                Port = 587,
                Timeout = 3600000,
                Host = "smtp.gmail.com",
                EnableSsl = true,
                Credentials = new NetworkCredential("huptduyuru@gmail.com", "Hupt1994+")
            };
            MailMessage mail = new MailMessage
            {
                From = new MailAddress("huptduyuru@gmail.com", "Hacettepe Üniversitesi Psikoloji Topluluğu (HÜPT)")
            };
            mail.To.Add("mtekatli@hacettepe.edu.tr");
            mail.Subject = "HÜPT - Kod Gönderecek Hesapta Düzenleme için Doğrulama Kodu"; mail.IsBodyHtml = true; mail.Body = "Merhaba sevgili HÜPT YK Üyesi/Üyeleri," + " <br/> <br/> " + "Doğrulama kodu gönderecek olan hesapta düzenleme yapabilmeniz için gereken doğrulama kodu: " + txt_gonderilenkod.Text + " <br/> <br/> " + "Bu kodu lütfen boş olarak görmüş olduğunuz " + @" ""Kod"" " + " kutucuğuna yazarak " + @" ""Girdiğim Kodu Onaylıyorum"" " + " butonuna basarak kod gönderecek olan hesapta düzenleme işleminizi gerçekleştirebilirsiniz." + " <br/> <br/> " + "İyi Günler/Akşamlar Dilerim";
            sc.Send(mail);
        }
        private void sifre_Load(object sender, EventArgs e)
        {
            metroLabel9.Text = Settings.Default.mail.ToString();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dialog;
            dialog = MessageBox.Show("Doğrulama Kodu Almak İstediğinize Emin Misiniz?", "Doğrulama Kodu Gönderimi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialog == DialogResult.Yes)
            {
                koduret();
                kodgonder();

                metroLabel3.Enabled = true;
                txt_alinankod.Enabled = true;
                button1.Enabled = false;
                button2.Enabled = true;
                metroLink1.Visible = true;

                MessageBox.Show("İşlem Başarılı. Lütfen size gönderilen doğrulama kodunu aktif olmuş olan kutucuğa yazınız ve onaylayınız.", "Doğrulama Kodu Gönderimi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("İşlem Başlatılmadı.", "Doğrulama Kodu Gönderimi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialog;
            dialog = MessageBox.Show("İşlemi Başlatmak İstediğinize Emin Misiniz?", "Doğrulama Kodu Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialog == DialogResult.Yes)
            {
                string a_bosluksuz = txt_alinankod.Text.Trim();
                txt_alinankod.Text = a_bosluksuz;

                if (txt_alinankod.Text == txt_gonderilenkod.Text)
                {
                    MessageBox.Show("İşlem Başarılı. Hesapta düzenleme işlemlerinizi yapmaya başlayabilirsiniz.", "Doğrulama Kodu Onayı", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    button1.Enabled = false;
                    button2.Enabled = false;
                    button3.Enabled = true;
                    metroLabel1.Enabled = true;
                    metroLabel2.Enabled = true;
                    metroLabel3.Enabled = false;
                    txt_alinankod.Enabled = false;
                    txt_mail.Enabled = true;
                    txt_sifre.Enabled = true;
                    pictureBox2.Visible = true;
                    metroLink1.Visible = false;
                }
                else
                {
                    MessageBox.Show("Hatalı doğrulama kodu girildi.", "Doğrulama Kodu Onayı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("İşlem Başlatılmadı.", "Doğrulama Kodu Onayı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialog;
            dialog = MessageBox.Show("İşlemi Başlatmak İstediğinize Emin Misiniz?", "Adres Düzenleme", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialog == DialogResult.Yes)
            {
                Settings.Default.mail = txt_mail.Text;
                Settings.Default.sifre = txt_sifre.Text;
                Settings.Default.Save();

                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                metroLabel1.Enabled = false;
                metroLabel2.Enabled = false;
                metroLabel3.Enabled = false;
                txt_alinankod.Enabled = false;
                txt_mail.Enabled = false;
                txt_sifre.Enabled = false;
                txt_sifre.PasswordChar = '*';
                pictureBox2.Visible = true;
                pictureBox1.Visible = false;

                metroLabel9.Text = Settings.Default.mail.ToString();

                MessageBox.Show("Değişiklik başarıyla kaydedildi. Doğrulama kodları bundan sonra " + txt_mail.Text + " adresinden gönderilecektir.", "İşlem Tamamlandı!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("İşlem Başlatılmadı.", "Adres Düzenleme", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            pictureBox2.Visible = false;

            txt_sifre.PasswordChar = char.Parse("\0");
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
            pictureBox2.Visible = true;

            txt_sifre.PasswordChar = char.Parse("*");
        }
        private void metroLink1_Click(object sender, EventArgs e)
        {
            DialogResult dialog;
            dialog = MessageBox.Show("Yeni Doğrulama Kodu Almak İstediğinize Emin Misiniz?", "Yeni Doğrulama Kodu Gönderimi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialog == DialogResult.Yes)
            {
                txt_gonderilenkod.Clear();
                koduret();
                kodgonder();

                metroLabel3.Enabled = true;
                txt_alinankod.Enabled = true;
                button1.Enabled = false;
                button2.Enabled = true;
                metroLink1.Visible = true;

                MessageBox.Show("İşlem Başarılı. Eski kod artık geçerli değildir. Lütfen size gönderilen yeni doğrulama kodunu aktif olmuş olan kutucuğa yazınız ve onaylayınız.", "Doğrulama Kodu Gönderimi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("İşlem Başlatılmadı", "Doğrulama Kodu Gönderimi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
