using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Net.Mail;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string constr = "Data Source=192.168.20.11\\SQLEXPRESS;Initial Catalog=UniManagement; User Id =sa; Password=1234qwer;";
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                string query2 = "SELECT * FROM xUsers WHERE Email IS NOT NULL";
                string query3 = "SELECT ID_uzytkownika, Data, (SELECT Czy_roboczy FROM Kalendarz As k WHERE (DATEPART (DAY, k.Data) =@Day AND DATEPART (Month, k.Data) =@Month AND DATEPART (Year, k.Data) =@Year)) AS czy_rob FROM dPracePrzyProjektach AS p WHERE DATEPART (DAY, p.Data) =@Day AND DATEPART (Month, p.Data) =@Month AND DATEPART (Year, p.Data) =@Year";
                //WHERE DATEPART (DAY, Data) =@Day AND DATEPART (Month, Data) =@Month
                //string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;

                SqlConnection conn = new SqlConnection(constr);
                try
                {
                    conn.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show (ex.Message);
                }


                SqlDataAdapter da = new SqlDataAdapter(query2, constr);
                da.Fill(ds, "xUsers");

                using (SqlCommand cmd = new SqlCommand(query3))
                {
                    cmd.Connection = conn;
                    cmd.Parameters.AddWithValue("@Day", DateTime.Today.Day-1);
                    cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month);
                    cmd.Parameters.AddWithValue("@Year", DateTime.Today.Year);
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(ds, "dPracePrzyProjektach");
                    }
                }


                /////wysłanie maili z brakami w raporcie
                if (DateTime.Today.Day == 5)
                {
                    //string query4 = "SELECT Count(czy_roboczy) AS ilosc_roboczych FROM Kalendarz AS k WHERE DATEPART (yyyy, k.Data) = @Year AND DATEPART (m, k.Data) = @Month AND czy_roboczy='TRUE'";
                    string query4 = "SELECT Count(czy_roboczy)*8 AS m_godz FROM Kalendarz AS k WHERE DATEPART (yyyy, k.Data) = '2017' AND DATEPART (m, k.Data) = '6' AND czy_roboczy='TRUE'";
                    string query5 = "SELECT ID_uzytkownika, Sum(Datediff(ss, Godzina_od, Godzina_do) / 3600.0) AS m_rob FROM dPracePrzyProjektach WHERE Datepart(m, Data) = 5 AND Datepart(yyyy, Data) = 2017  Group BY ID_uzytkownika";
                    string query6 = "SELECT ID_uzytkownika, Data, Sum(Datediff(ss, Godzina_od, Godzina_do) / 3600.0) AS d_rob FROM dPracePrzyProjektach WHERE Datepart(m, Data) = 5 AND Datepart(yyyy, Data) = 2017 Group BY ID_uzytkownika, Data";
                    //string query6 = "SELECT dPracePrzyProjektach.ID_uzytkownika, Kalendarz.Data, Kalendarz.Dzien_tygodnia, Kalendarz.Czy_roboczy, Sum(Datediff(ss, dPracePrzyProjektach.Godzina_od, dPracePrzyProjektach.Godzina_do) / 3600.0) AS d_rob FROM Kalendarz LEFT OUTER JOIN dPracePrzyProjektach ON dPracePrzyProjektach.Data=Kalendarz.Data WHERE Datepart(m, Kalendarz.Data) = 5 AND Datepart(yyyy, Kalendarz.Data) = 2017 Group BY dPracePrzyProjektach.ID_uzytkownika, Kalendarz.Data, Kalendarz.Dzien_tygodnia, Kalendarz.Czy_roboczy ORDER BY Kalendarz.Data";
                    string query7 = "SELECT Data, Dzien_tygodnia, Czy_roboczy FROM Kalendarz WHERE Datepart(m, Data) = 5 AND Datepart(yyyy, Data) = 2017";

                    using (SqlCommand cmd = new SqlCommand(query5))
                    {
                        cmd.Connection = conn;
                        cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month);
                        cmd.Parameters.AddWithValue("@Year", DateTime.Today.Year);
                        using (SqlDataAdapter mda = new SqlDataAdapter(cmd))
                        {
                            mda.Fill(ds, "rob_miesiac");
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand(query4))
                    {
                        cmd.Connection = conn;
                        cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month);
                        cmd.Parameters.AddWithValue("@Year", DateTime.Today.Year);
                        using (SqlDataAdapter mda2 = new SqlDataAdapter(cmd))
                        {
                            mda2.Fill(ds, "miesiac");
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand(query6))
                    {
                        cmd.Connection = conn;
                        cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month);
                        cmd.Parameters.AddWithValue("@Year", DateTime.Today.Year);
                        using (SqlDataAdapter mda = new SqlDataAdapter(cmd))
                        {
                            mda.Fill(ds, "rob_dni");
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand(query7))
                    {
                        cmd.Connection = conn;
                        cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month);
                        cmd.Parameters.AddWithValue("@Year", DateTime.Today.Year);
                        using (SqlDataAdapter mda = new SqlDataAdapter(cmd))
                        {
                            mda.Fill(ds, "kalend");
                        }
                    }

                    foreach (DataRow row in ds.Tables["rob_miesiac"].Rows)
                    {
                        string tekst = "";
                        int m_rob = Convert.ToInt32(row["m_rob"]);
                        int m_godz = (int) ds.Tables["miesiac"].Rows[0][0];
                        if (m_rob > m_godz)
                        {
                            int wydruk = m_rob - m_godz;
                            tekst = "w miesiącu " + DateTime.Today.Month + " masz nadgodziny w liczbie " + wydruk + " godzin";
                            //MessageBox.Show(tekst);
                        }
                        if (m_rob < m_godz)
                        {
                            int wydruk = m_godz - m_rob;
                            tekst = "w miesiącu " + DateTime.Today.Month + " masz za mało przepracowanych godzin o " + wydruk + " godziny";
                            //MessageBox.Show(tekst);
                        }
                        if (m_rob == m_godz)
                        {
                            tekst = "w miesiącu " + DateTime.Today.Month + " masz przepracowane równe " + m_rob + " godziny";
                            //MessageBox.Show(tekst);
                        }

                        foreach (DataRow row2 in ds.Tables["xUsers"].Rows)
                        {
                            string messageBody = "<font>Poniżej zamieszczono wykaz z poprzedniego miesiąca: </font><br><br>";
                            string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                            string htmlTableEnd = "</table>";
                            string htmlHeaderRowStart = "<tr style =\"background-color:#6FA1D2; color:#ffffff;\">";
                            string htmlHeaderRowEnd = "</tr>";
                            string htmlTrStart = "<tr style =\"color:#555555;\">";
                            string htmlTrEnd = "</tr>";
                            string htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                            string htmlTdEnd = "</td>";

                            messageBody += htmlTableStart;
                            messageBody += htmlHeaderRowStart;
                            messageBody += htmlTdStart + "Data" + htmlTdEnd;
                            messageBody += htmlTdStart + "Dzień tygodnia" + htmlTdEnd;
                            messageBody += htmlTdStart + "Ilość przepracowanych godzin " + htmlTdEnd;
                            messageBody += htmlHeaderRowEnd;

                            foreach (DataRow Row2 in ds.Tables["kalend"].Rows)
                            {
                            if (Row["ID_uzytkownika"].ToString()==row2["Id"].ToString())
                                {
                                    foreach (DataRow Row in ds.Tables["rob_dni"].Rows)
                                    {
                                        if (Row["Data"].ToString() == Row2["Data"].ToString())
                                        {
                                            messageBody = messageBody + htmlTrStart;
                                            string data = Row["Data"].ToString();
                                            messageBody = messageBody + htmlTdStart + data + htmlTdEnd;
                                            messageBody = messageBody + htmlTdStart + Row["Dzien_tygodnia"] + htmlTdEnd;
                                            messageBody = messageBody + htmlTdStart + Row["d_rob"] + htmlTdEnd;
                                            messageBody = messageBody + htmlTrEnd;
                                        }
                                    }

                                }
                            }
                            messageBody = messageBody + htmlTableEnd;
                            if (row2["Id"].ToString() == row["ID_uzytkownika"].ToString())
                            {
                                string email = "s3rek92@gmail.com";
                                //string email = row["Email"].ToString();

                                MessageBox.Show("Trying to send email to: " + row2["Nazwisko"]);//zmieniec id na email

                                using (MailMessage mm = new MailMessage("Unimap.katowice@gmail.com", email))
                                {
                                    mm.Subject = "Miesięczna kontrola bazy";
                                    mm.IsBodyHtml = true;
                                    mm.Body = string.Format("<p>Witaj! "+row2["Imie"]+row2["Nazwisko"]+"</p>" + messageBody+ tekst);

                                    SmtpClient smtp = new SmtpClient();
                                    smtp.Host = "smtp.gmail.com";
                                    smtp.EnableSsl = true;
                                    System.Net.NetworkCredential credentials = new System.Net.NetworkCredential();
                                    credentials.UserName = "Unimap.katowice@gmail.com";
                                    credentials.Password = "MalBry2015";
                                    smtp.UseDefaultCredentials = true;
                                    smtp.Credentials = credentials;
                                    smtp.Port = 587;
                                    smtp.Send(mm);
                                    //MessageBox.Show("Email sent successfully to: " + email);
                                }
                            }
                        }
                    }
                }
                ///////////////////////////////////




                bool rob = (from DataRow dr in ds.Tables["dPracePrzyProjektach"].Rows
                          select (bool)dr["czy_rob"]).FirstOrDefault();
                if (rob)
                {

                    bool[] wierszeDoUsuniecia = new bool[ds.Tables["xUsers"].Rows.Count];

                    for (int j = 0; j < ds.Tables["xUsers"].Rows.Count; j++)
                    {
                        DataRow row = ds.Tables["xUsers"].Rows[j];
                        string ID = row["ID"].ToString();

                        for (int i = ds.Tables["dPracePrzyProjektach"].Rows.Count - 1; i >= 0; i--)
                        {
                            DataRow delIDS = ds.Tables["dPracePrzyProjektach"].Rows[i];
                            if (delIDS["ID_uzytkownika"].ToString() == ID)
                            {
                                try
                                {
                                    wierszeDoUsuniecia[j] = true;
                                }
                                catch { }
                            }
                        }

                    }

                    for (int i = ds.Tables["xUsers"].Rows.Count - 1; i >= 0; i--)
                    {
                        if (wierszeDoUsuniecia[i])
                        {
                            DataRow drrr = ds.Tables["xUsers"].Rows[i];
                            ds.Tables["xUsers"].Rows.Remove(drrr);
                        }
                    }

                    foreach (DataRow row in ds.Tables["xUsers"].Rows)
                    {


                        string email = "s3rek92@gmail.com";
                        //string email = row["Email"].ToString();

                        MessageBox.Show("Trying to send email to: " + email);

                        using (MailMessage mm = new MailMessage("Unimap.katowice@gmail.com", email))
                        {
                            mm.Subject = "test";
                            mm.Body = string.Format("<b>test</b>");

                            mm.IsBodyHtml = true;
                            SmtpClient smtp = new SmtpClient();
                            smtp.Host = "smtp.gmail.com";
                            smtp.EnableSsl = true;
                            System.Net.NetworkCredential credentials = new System.Net.NetworkCredential();
                            credentials.UserName = "Unimap.katowice@gmail.com";
                            credentials.Password = "MalBry2015";
                            smtp.UseDefaultCredentials = true;
                            smtp.Credentials = credentials;
                            smtp.Port = 587;
                            smtp.Send(mm);
                            MessageBox.Show("Email sent successfully to: " + email);
                        }
                    }
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
               // using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("SimpleService"))
               // {
                //    serviceController.Stop();
               // }
            }
        }
    }
}
