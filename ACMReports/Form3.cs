using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ACMReports
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();

            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            // Read settings from application configuration file
            textBox1.Text = Properties.Settings.Default["EntranceEvent"].ToString();
            textBox2.Text = Properties.Settings.Default["EntranceDoor"].ToString();
            textBox3.Text = Properties.Settings.Default["ExitEvent"].ToString();
            textBox4.Text = Properties.Settings.Default["ExitDoor"].ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Remove danger characters from SQL input parameters before save it
            char[] charsToRemove = {'@', ';', '\\', '\'', '/', '%', '|', '"', '[', ']', '(', ')', ',',
                                    '*', '?', '{', '}', '$', '#', '`', '~', '!', '^', '<', '>', '=',
                                    '\0', '\n', '\r', '\t'};

            foreach (Control gb in this.Controls)
            {
                if (gb is GroupBox)
                {
                    foreach (Control tb in gb.Controls)
                    {
                        if (tb is TextBox)
                        {
                            string str = ((TextBox)tb).Text;
                            foreach (char chr in charsToRemove)
                            {
                                str = str.Replace(System.Convert.ToString(chr), String.Empty);
                            }
                            tb.Text = str;
                        }
                    }
                }
            }

            // Save SQL Request
            Properties.Settings.Default["EntranceEvent"] = textBox1.Text;
            Properties.Settings.Default["EntranceDoor"] = textBox2.Text;
            Properties.Settings.Default["ExitEvent"] = textBox3.Text;
            Properties.Settings.Default["ExitDoor"] = textBox4.Text;
            Properties.Settings.Default.Save();
        }

    }
}
