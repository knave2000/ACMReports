using System;
using System.Windows.Forms;
using Npgsql;
using System.Threading.Tasks;

namespace ACMReports
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();

            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            // Read settings from application configuration file
            textBox1.Text = Properties.Settings.Default["ACMServerIP"].ToString();
            textBox2.Text = Properties.Settings.Default["ACMServerPort"].ToString();
            textBox3.Text = Properties.Settings.Default["ACMUsername"].ToString();
            textBox4.Text = Properties.Settings.Default["ACMPassword"].ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Save settings in application configuration file
            Properties.Settings.Default["ACMServerIP"] = textBox1.Text;
            Properties.Settings.Default["ACMServerPort"] = textBox2.Text;
            Properties.Settings.Default["ACMUsername"] = textBox3.Text;
            Properties.Settings.Default["ACMPassword"] = textBox4.Text;
            Properties.Settings.Default.Save();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Test the connection to ACM transaction database
            try
            {
                // Read settings from application configuration file
                string ACMServerIP = Properties.Settings.Default["ACMServerIP"].ToString();
                string ACMServerPort = Properties.Settings.Default["ACMServerPort"].ToString();
                string ACMUsername = Properties.Settings.Default["ACMUsername"].ToString();
                string ACMPassword = Properties.Settings.Default["ACMPassword"].ToString();

                // PostgeSQL-style connection string
                // SSL encryption is required for ACM connection
                string connStr = String.Format("Server={0};Port={1};UserId={2};Password={3};" +
                        "Database=TransactionDB;SSL Mode=Require;Trust Server Certificate=true",
                        ACMServerIP, ACMServerPort, ACMUsername, ACMPassword);

                NpgsqlConnection conn = new NpgsqlConnection(connStr);
                conn.Open();
                conn.Close();
                MessageBox.Show("Connection to ACM successfully installed.");
            }
            catch
            {
                // if something went wrong, and you want to know why
                MessageBox.Show("Connection to ACM failed.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
