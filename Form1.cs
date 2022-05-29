using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using itextsharp.pdfa;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Threading;

namespace CGPA_CALCULATOR
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            MaximizeBox = false;
            MessageBox.Show("Useful Instructions:" + Environment.NewLine + "1. Please Enter Registration Numbers Using the Formats Below:" + Environment.NewLine + "a. EE10254." + Environment.NewLine + "b. EE-10-254." + Environment.NewLine + "c. Please Don't Use EE/10/254." + Environment.NewLine + "2. Click the Export Button to Export CGPA Data to the CGPA Records Folder Located on the Desktop." + Environment.NewLine + "3. Insert New Course Code and Credit Unit and then Click on the Save Button to Save." + Environment.NewLine + "4. If you want to Delete Any Saved Course Code, Simply Select the Course Code Using the First Course Select Box and then Click on the Delete Course Code Button." + Environment.NewLine + " " + Environment.NewLine + "For More Infomation or Help, Contact Us On:" + Environment.NewLine + "1. Email Addresses: akifagoelectronics@yahoo.com, akifagoelect@yahoo.com" + Environment.NewLine + "2. Phone Numbers: +2347031011012, +2348022889554");
            string filepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            filepath = filepath + @"\CGPA RECORDS\";
            if (!Directory.Exists(filepath))
            {
                Directory.CreateDirectory(filepath);
            }
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            if (!Directory.Exists(filepath1))
            {
                Directory.CreateDirectory(filepath1);
                File.WriteAllText(Path.Combine(filepath1,"DEFAULT.txt"),"0");
            }
            string [] files = Directory.GetFiles(filepath1,"*.txt",SearchOption.AllDirectories);
            var namesonly = files.Select(f => Path.GetFileNameWithoutExtension(f)).ToArray();
            comboBox1.Items.AddRange(namesonly);
            comboBox2.Items.AddRange(namesonly);
            comboBox3.Items.AddRange(namesonly);
            comboBox4.Items.AddRange(namesonly);
            comboBox5.Items.AddRange(namesonly);
            comboBox6.Items.AddRange(namesonly);
            comboBox7.Items.AddRange(namesonly);
            comboBox8.Items.AddRange(namesonly);
            comboBox9.Items.AddRange(namesonly);
            comboBox10.Items.AddRange(namesonly);
            comboBox11.Items.AddRange(namesonly);
            comboBox12.Items.AddRange(namesonly);
            comboBox13.Items.AddRange(namesonly);
            comboBox14.Items.AddRange(namesonly);
            comboBox15.Items.AddRange(namesonly);
            comboBox16.Items.AddRange(namesonly);
            comboBox17.Items.AddRange(namesonly);
            comboBox18.Items.AddRange(namesonly);
            comboBox19.Items.AddRange(namesonly);
            comboBox20.Items.AddRange(namesonly);
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox1.Text, out result))
            {
                int a = Convert.ToInt32(textBox1.Text);
            if (int.TryParse(textBox2.Text, out result))
            {
                int b = Convert.ToInt32(textBox2.Text);
                    var c = a + b;
                    textBox3.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox124.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox124.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox124.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox124.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox124.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox124.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox124.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox124.Text = "ERROR";
                    }
                }
            }
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

            int result = 0;
            if (int.TryParse(textBox1.Text, out result))
            {
                int a = Convert.ToInt32(textBox1.Text);
                if (int.TryParse(textBox2.Text, out result))
                {
                    int b = Convert.ToInt32(textBox2.Text);
                    var c = a + b;
                    textBox3.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox124.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox124.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox124.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox124.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox124.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox124.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox124.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox124.Text = "ERROR";
                    }
                }
            }
        }
        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox9.Text, out result))
            {
                int a = Convert.ToInt32(textBox9.Text);
                if (int.TryParse(textBox10.Text, out result))
                {
                    int b = Convert.ToInt32(textBox10.Text);
                    var c = a + b;
                    textBox8.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox123.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox123.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox123.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox123.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox123.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox123.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox123.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox123.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox10.Text, out result))
            {
                int a = Convert.ToInt32(textBox10.Text);
                if (int.TryParse(textBox9.Text, out result))
                {
                    int b = Convert.ToInt32(textBox9.Text);
                    var c = a + b;
                    textBox8.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox123.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox123.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox123.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox123.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox123.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox123.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox123.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox123.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox14.Text, out result))
            {
                int a = Convert.ToInt32(textBox14.Text);
                if (int.TryParse(textBox15.Text, out result))
                {
                    int b = Convert.ToInt32(textBox15.Text);
                    var c = a + b;
                    textBox13.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox122.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox122.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox122.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox122.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox122.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox122.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox122.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox122.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox15.Text, out result))
            {
                int a = Convert.ToInt32(textBox15.Text);
                if (int.TryParse(textBox14.Text, out result))
                {
                    int b = Convert.ToInt32(textBox14.Text);
                    var c = a + b;
                    textBox13.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox122.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox122.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox122.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox122.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox122.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox122.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox122.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox122.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox19.Text, out result))
            {
                int a = Convert.ToInt32(textBox19.Text);
                if (int.TryParse(textBox20.Text, out result))
                {
                    int b = Convert.ToInt32(textBox20.Text);
                    var c = a + b;
                    textBox18.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox121.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox121.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox121.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox121.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox121.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox121.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox121.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox121.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox19.Text, out result))
            {
                int a = Convert.ToInt32(textBox19.Text);
                if (int.TryParse(textBox20.Text, out result))
                {
                    int b = Convert.ToInt32(textBox20.Text);
                    var c = a + b;
                    textBox18.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox121.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox121.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox121.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox121.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox121.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox121.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox121.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox121.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox25_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox24.Text, out result))
            {
                int a = Convert.ToInt32(textBox24.Text);
                if (int.TryParse(textBox25.Text, out result))
                {
                    int b = Convert.ToInt32(textBox25.Text);
                    var c = a + b;
                    textBox23.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox120.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox120.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox120.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox120.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox120.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox120.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox120.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox120.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox24.Text, out result))
            {
                int a = Convert.ToInt32(textBox24.Text);
                if (int.TryParse(textBox25.Text, out result))
                {
                    int b = Convert.ToInt32(textBox25.Text);
                    var c = a + b;
                    textBox23.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox120.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox120.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox120.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox120.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox120.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox120.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox120.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox120.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox29.Text, out result))
            {
                int a = Convert.ToInt32(textBox29.Text);
                if (int.TryParse(textBox30.Text, out result))
                {
                    int b = Convert.ToInt32(textBox30.Text);
                    var c = a + b;
                    textBox28.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox119.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox119.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox119.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox119.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox119.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox119.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox119.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox119.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox29_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox29.Text, out result))
            {
                int a = Convert.ToInt32(textBox29.Text);
                if (int.TryParse(textBox30.Text, out result))
                {
                    int b = Convert.ToInt32(textBox30.Text);
                    var c = a + b;
                    textBox28.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox119.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox119.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox119.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox119.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox119.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox119.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox119.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox119.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox35_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox34.Text, out result))
            {
                int a = Convert.ToInt32(textBox34.Text);
                if (int.TryParse(textBox35.Text, out result))
                {
                    int b = Convert.ToInt32(textBox35.Text);
                    var c = a + b;
                    textBox33.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox118.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox118.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox118.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox118.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox118.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox118.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox118.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox118.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox34_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox34.Text, out result))
            {
                int a = Convert.ToInt32(textBox34.Text);
                if (int.TryParse(textBox35.Text, out result))
                {
                    int b = Convert.ToInt32(textBox35.Text);
                    var c = a + b;
                    textBox33.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox118.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox118.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox118.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox118.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox118.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox118.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox118.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox118.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox40_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox39.Text, out result))
            {
                int a = Convert.ToInt32(textBox39.Text);
                if (int.TryParse(textBox40.Text, out result))
                {
                    int b = Convert.ToInt32(textBox40.Text);
                    var c = a + b;
                    textBox38.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox117.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox117.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox117.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox117.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox117.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox117.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox117.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox117.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox39_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox39.Text, out result))
            {
                int a = Convert.ToInt32(textBox39.Text);
                if (int.TryParse(textBox40.Text, out result))
                {
                    int b = Convert.ToInt32(textBox40.Text);
                    var c = a + b;
                    textBox38.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox117.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox117.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox117.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox117.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox117.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox117.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox117.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox117.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox45_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox44.Text, out result))
            {
                int a = Convert.ToInt32(textBox44.Text);
                if (int.TryParse(textBox45.Text, out result))
                {
                    int b = Convert.ToInt32(textBox45.Text);
                    var c = a + b;
                    textBox43.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox116.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox116.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox116.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox116.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox116.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox116.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox116.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox116.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox44_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox44.Text, out result))
            {
                int a = Convert.ToInt32(textBox44.Text);
                if (int.TryParse(textBox45.Text, out result))
                {
                    int b = Convert.ToInt32(textBox45.Text);
                    var c = a + b;
                    textBox43.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox116.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox116.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox116.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox116.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox116.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox116.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox116.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox116.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox50_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox49.Text, out result))
            {
                int a = Convert.ToInt32(textBox49.Text);
                if (int.TryParse(textBox50.Text, out result))
                {
                    int b = Convert.ToInt32(textBox50.Text);
                    var c = a + b;
                    textBox48.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox115.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox115.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox115.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox115.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox115.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox115.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox115.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox115.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox49_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox49.Text, out result))
            {
                int a = Convert.ToInt32(textBox49.Text);
                if (int.TryParse(textBox50.Text, out result))
                {
                    int b = Convert.ToInt32(textBox50.Text);
                    var c = a + b;
                    textBox48.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox115.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox115.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox115.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox115.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox115.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox115.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox115.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox115.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox55_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox54.Text, out result))
            {
                int a = Convert.ToInt32(textBox54.Text);
                if (int.TryParse(textBox55.Text, out result))
                {
                    int b = Convert.ToInt32(textBox55.Text);
                    var c = a + b;
                    textBox53.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox114.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox114.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox114.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox114.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox114.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox114.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox114.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox114.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox54_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox54.Text, out result))
            {
                int a = Convert.ToInt32(textBox54.Text);
                if (int.TryParse(textBox55.Text, out result))
                {
                    int b = Convert.ToInt32(textBox55.Text);
                    var c = a + b;
                    textBox53.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox114.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox114.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox114.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox114.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox114.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox114.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox114.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox114.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox60_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox59.Text, out result))
            {
                int a = Convert.ToInt32(textBox59.Text);
                if (int.TryParse(textBox60.Text, out result))
                {
                    int b = Convert.ToInt32(textBox60.Text);
                    var c = a + b;
                    textBox58.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox110.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox110.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox110.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox110.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox110.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox110.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox110.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox110.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox59_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox59.Text, out result))
            {
                int a = Convert.ToInt32(textBox59.Text);
                if (int.TryParse(textBox60.Text, out result))
                {
                    int b = Convert.ToInt32(textBox60.Text);
                    var c = a + b;
                    textBox58.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox110.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox110.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox110.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox110.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox110.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox110.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox110.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox110.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox65_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox64.Text, out result))
            {
                int a = Convert.ToInt32(textBox64.Text);
                if (int.TryParse(textBox65.Text, out result))
                {
                    int b = Convert.ToInt32(textBox65.Text);
                    var c = a + b;
                    textBox63.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox111.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox111.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox111.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox111.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox111.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox111.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox111.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox111.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox64_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox64.Text, out result))
            {
                int a = Convert.ToInt32(textBox64.Text);
                if (int.TryParse(textBox65.Text, out result))
                {
                    int b = Convert.ToInt32(textBox65.Text);
                    var c = a + b;
                    textBox63.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox111.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox111.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox111.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox111.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox111.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox111.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox111.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox111.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox70_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox69.Text, out result))
            {
                int a = Convert.ToInt32(textBox69.Text);
                if (int.TryParse(textBox70.Text, out result))
                {
                    int b = Convert.ToInt32(textBox70.Text);
                    var c = a + b;
                    textBox68.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox112.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox112.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox112.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox112.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox112.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox112.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox112.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox112.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox69_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox69.Text, out result))
            {
                int a = Convert.ToInt32(textBox69.Text);
                if (int.TryParse(textBox70.Text, out result))
                {
                    int b = Convert.ToInt32(textBox70.Text);
                    var c = a + b;
                    textBox68.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox112.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox112.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox112.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox112.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox112.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox112.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox112.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox112.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox75_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox74.Text, out result))
            {
                int a = Convert.ToInt32(textBox74.Text);
                if (int.TryParse(textBox75.Text, out result))
                {
                    int b = Convert.ToInt32(textBox75.Text);
                    var c = a + b;
                    textBox73.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox113.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox113.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox113.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox113.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox113.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox113.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox113.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox113.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox74_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox74.Text, out result))
            {
                int a = Convert.ToInt32(textBox74.Text);
                if (int.TryParse(textBox75.Text, out result))
                {
                    int b = Convert.ToInt32(textBox75.Text);
                    var c = a + b;
                    textBox73.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox113.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox113.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox113.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox113.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox113.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox113.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox113.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox113.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox80_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox79.Text, out result))
            {
                int a = Convert.ToInt32(textBox79.Text);
                if (int.TryParse(textBox80.Text, out result))
                {
                    int b = Convert.ToInt32(textBox80.Text);
                    var c = a + b;
                    textBox78.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox109.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox109.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox109.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox109.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox109.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox109.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox109.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox109.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox79_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox79.Text, out result))
            {
                int a = Convert.ToInt32(textBox79.Text);
                if (int.TryParse(textBox80.Text, out result))
                {
                    int b = Convert.ToInt32(textBox80.Text);
                    var c = a + b;
                    textBox78.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox109.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox109.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox109.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox109.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox109.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox109.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox109.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox109.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox85_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox84.Text, out result))
            {
                int a = Convert.ToInt32(textBox84.Text);
                if (int.TryParse(textBox85.Text, out result))
                {
                    int b = Convert.ToInt32(textBox85.Text);
                    var c = a + b;
                    textBox83.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox108.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox108.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox108.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox108.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox108.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox108.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox108.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox108.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox84_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox84.Text, out result))
            {
                int a = Convert.ToInt32(textBox84.Text);
                if (int.TryParse(textBox85.Text, out result))
                {
                    int b = Convert.ToInt32(textBox85.Text);
                    var c = a + b;
                    textBox83.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox108.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox108.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox108.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox108.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox108.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox108.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox108.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox108.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox90_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox89.Text, out result))
            {
                int a = Convert.ToInt32(textBox89.Text);
                if (int.TryParse(textBox90.Text, out result))
                {
                    int b = Convert.ToInt32(textBox90.Text);
                    var c = a + b;
                    textBox88.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox107.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox107.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox107.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox107.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox107.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox107.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox107.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox107.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox89_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox89.Text, out result))
            {
                int a = Convert.ToInt32(textBox89.Text);
                if (int.TryParse(textBox90.Text, out result))
                {
                    int b = Convert.ToInt32(textBox90.Text);
                    var c = a + b;
                    textBox88.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox107.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox107.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox107.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox107.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox107.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox107.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox107.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox107.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox99_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox98.Text, out result))
            {
                int a = Convert.ToInt32(textBox98.Text);
                if (int.TryParse(textBox99.Text, out result))
                {
                    int b = Convert.ToInt32(textBox99.Text);
                    var c = a + b;
                    textBox97.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox106.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox106.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox106.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox106.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox106.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox106.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox106.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox106.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox98_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox98.Text, out result))
            {
                int a = Convert.ToInt32(textBox98.Text);
                if (int.TryParse(textBox99.Text, out result))
                {
                    int b = Convert.ToInt32(textBox99.Text);
                    var c = a + b;
                    textBox97.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox106.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox106.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox106.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox106.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox106.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox106.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox106.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox106.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox104_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox103.Text, out result))
            {
                int a = Convert.ToInt32(textBox103.Text);
                if (int.TryParse(textBox104.Text, out result))
                {
                    int b = Convert.ToInt32(textBox104.Text);
                    var c = a + b;
                    textBox102.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox93.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox93.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox93.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox93.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox93.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox93.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox93.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox93.Text = "ERROR";
                    }
                }
            }
        }

        private void textBox103_TextChanged(object sender, EventArgs e)
        {
            int result = 0;
            if (int.TryParse(textBox103.Text, out result))
            {
                int a = Convert.ToInt32(textBox103.Text);
                if (int.TryParse(textBox104.Text, out result))
                {
                    int b = Convert.ToInt32(textBox104.Text);
                    var c = a + b;
                    textBox102.Text = c.ToString();
                    if (c < 40)
                    {
                        textBox93.Text = "F";
                    }
                    if (c >= 40)
                    {
                        textBox93.Text = "E";
                    }
                    if (c >= 45)
                    {
                        textBox93.Text = "D";
                    }
                    if (c >= 50)
                    {
                        textBox93.Text = "C";
                    }
                    if (c >= 60)
                    {
                        textBox93.Text = "B";
                    }
                    if (c >= 70)
                    {
                        textBox93.Text = "A";
                    }
                    if (c == 100)
                    {
                        textBox93.Text = "A";
                    }
                    if (c > 100)
                    {
                        textBox93.Text = "ERROR";
                    }
                }
            }
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            if (comboBox1.Text == "" || textBox3.Text == "")
            {
                MessageBox.Show("Please Select a Course Code From the List and Insert All Fields to Calculate");
            }
            else
            {
                if (comboBox2.Text == "" || textBox8.Text == "")
                {
                    textBox6.Text = "0";
                    textBox7.Text = "0";
                }
                if (comboBox3.Text == "" || textBox13.Text == "")
                {
                    textBox11.Text = "0";
                    textBox12.Text = "0";
                }
                if (comboBox4.Text == "" || textBox18.Text == "")
                {
                    textBox16.Text = "0";
                    textBox17.Text = "0";
                }
                if (comboBox5.Text == "" || textBox23.Text == "")
                {
                    textBox21.Text = "0";
                    textBox22.Text = "0";
                }
                if (comboBox6.Text == "" || textBox28.Text == "")
                {
                    textBox26.Text = "0";
                    textBox27.Text = "0";
                }
                if (comboBox7.Text == "" || textBox33.Text == "")
                {
                    textBox31.Text = "0";
                    textBox32.Text = "0";
                }
                if (comboBox8.Text == "" || textBox38.Text == "")
                {
                    textBox36.Text = "0";
                    textBox37.Text = "0";
                }
                if (comboBox9.Text == "" || textBox43.Text == "")
                {
                    textBox41.Text = "0";
                    textBox42.Text = "0";
                }
                if (comboBox10.Text == "" || textBox48.Text == "")
                {
                    textBox46.Text = "0";
                    textBox47.Text = "0";
                }
                if (comboBox11.Text == "" || textBox53.Text == "")
                {
                    textBox51.Text = "0";
                    textBox52.Text = "0";
                }
                if (comboBox12.Text == "" || textBox58.Text == "")
                {
                    textBox56.Text = "0";
                    textBox57.Text = "0";
                }
                if (comboBox13.Text == "" || textBox63.Text == "")
                {
                    textBox61.Text = "0";
                    textBox62.Text = "0";
                }
                if (comboBox14.Text == "" || textBox68.Text == "")
                {
                    textBox66.Text = "0";
                    textBox67.Text = "0";
                }
                if (comboBox15.Text == "" || textBox73.Text == "")
                {
                    textBox71.Text = "0";
                    textBox72.Text = "0";
                }
                if (comboBox16.Text == "" || textBox78.Text == "")
                {
                    textBox76.Text = "0";
                    textBox77.Text = "0";
                }
                if (comboBox17.Text == "" || textBox83.Text == "")
                {
                    textBox81.Text = "0";
                    textBox82.Text = "0";
                }
                if (comboBox18.Text == "" || textBox88.Text == "")
                {
                    textBox86.Text = "0";
                    textBox87.Text = "0";
                }
                if (comboBox19.Text == "" || textBox97.Text == "")
                {
                    textBox95.Text = "0";
                    textBox96.Text = "0";
                }
                if (comboBox20.Text == "" || textBox102.Text == "")
                {
                    textBox100.Text = "0";
                    textBox101.Text = "0";
                }
                double a = Convert.ToDouble(textBox5.Text);
                double b = Convert.ToDouble(textBox6.Text);
                double c = Convert.ToDouble(textBox11.Text);
                double d = Convert.ToDouble(textBox16.Text);
                double g = Convert.ToDouble(textBox21.Text);
                double h = Convert.ToDouble(textBox26.Text);
                double i = Convert.ToDouble(textBox31.Text);
                double j = Convert.ToDouble(textBox36.Text);
                double k = Convert.ToDouble(textBox41.Text);
                double l = Convert.ToDouble(textBox46.Text);
                double m = Convert.ToDouble(textBox51.Text);
                double n = Convert.ToDouble(textBox56.Text);
                double o = Convert.ToDouble(textBox61.Text);
                double p = Convert.ToDouble(textBox66.Text);
                double q = Convert.ToDouble(textBox71.Text);
                double r = Convert.ToDouble(textBox76.Text);
                double s = Convert.ToDouble(textBox81.Text);
                double t = Convert.ToDouble(textBox86.Text);
                double u = Convert.ToDouble(textBox95.Text);
                double v = Convert.ToDouble(textBox100.Text);

                double aa = Convert.ToDouble(textBox4.Text);
                double bb = Convert.ToDouble(textBox7.Text);
                double cc = Convert.ToDouble(textBox12.Text);
                double dd = Convert.ToDouble(textBox17.Text);
                double gg = Convert.ToDouble(textBox22.Text);
                double hh = Convert.ToDouble(textBox27.Text);
                double ii = Convert.ToDouble(textBox32.Text);
                double jj = Convert.ToDouble(textBox37.Text);
                double kk = Convert.ToDouble(textBox42.Text);
                double ll = Convert.ToDouble(textBox47.Text);
                double mm = Convert.ToDouble(textBox52.Text);
                double nn = Convert.ToDouble(textBox57.Text);
                double oo = Convert.ToDouble(textBox62.Text);
                double pp = Convert.ToDouble(textBox67.Text);
                double qq = Convert.ToDouble(textBox72.Text);
                double rr = Convert.ToDouble(textBox77.Text);
                double ss = Convert.ToDouble(textBox82.Text);
                double tt = Convert.ToDouble(textBox87.Text);
                double uu = Convert.ToDouble(textBox96.Text);
                double vv = Convert.ToDouble(textBox101.Text);

                double f = Convert.ToDouble(string.Format("{0:0.00}", ((a + b + c + d + g + h + i + j + k + l + m + n + o + p + q + r + s + t + u + v) / (aa + bb + cc + dd + gg + hh + ii + jj + kk + ll + mm + nn + oo + pp + qq + rr + ss + tt + uu + vv))));
                textBox105.Text = f.ToString();
                if (f == 0)
                {
                    textBox92.Text = "";
                }
                if (f > 0)
                {
                    textBox92.Text = "FAILED";
                }
                if (f >= 1.5)
                {
                    textBox92.Text = "THIRD CLASS";
                }
                if (f >= 2.5)
                {
                    textBox92.Text = "SECOND CLASS LOWER";
                }
                if (f >= 3.5)
                {
                    textBox92.Text = "SECOND CLASS UPPER";
                }
                if (f >= 4.5)
                {
                    textBox92.Text = "FIRST CLASS";
                }
                if (f == 5)
                {
                    textBox92.Text = "FIRST CLASS";
                }
                Thread.Sleep(50);
                //Clears all the Zeros
                if (comboBox1.Text == "" || textBox3.Text == "")
                {
                    textBox5.Text = " ";
                    textBox4.Text = " ";
                }
                if (comboBox2.Text == "" || textBox8.Text == "")
                {
                    textBox6.Text = " ";
                    textBox7.Text = " ";
                }
                if (comboBox3.Text == "" || textBox13.Text == "")
                {
                    textBox11.Text = " ";
                    textBox12.Text = " ";
                }
                if (comboBox4.Text == "" || textBox18.Text == "")
                {
                    textBox16.Text = " ";
                    textBox17.Text = " ";
                }
                if (comboBox5.Text == "" || textBox23.Text == "")
                {
                    textBox21.Text = " ";
                    textBox22.Text = " ";
                }
                if (comboBox6.Text == "" || textBox28.Text == "")
                {
                    textBox26.Text = " ";
                    textBox27.Text = " ";
                }
                if (comboBox7.Text == "" || textBox33.Text == "")
                {
                    textBox31.Text = " ";
                    textBox32.Text = " ";
                }
                if (comboBox8.Text == "" || textBox38.Text == "")
                {
                    textBox36.Text = " ";
                    textBox37.Text = " ";
                }
                if (comboBox9.Text == "" || textBox43.Text == "")
                {
                    textBox41.Text = " ";
                    textBox42.Text = " ";
                }
                if (comboBox10.Text == "" || textBox48.Text == "")
                {
                    textBox46.Text = " ";
                    textBox47.Text = " ";
                }
                if (comboBox11.Text == "" || textBox53.Text == "")
                {
                    textBox51.Text = " ";
                    textBox52.Text = " ";
                }
                if (comboBox12.Text == "" || textBox58.Text == "")
                {
                    textBox56.Text = " ";
                    textBox57.Text = " ";
                }
                if (comboBox13.Text == "" || textBox63.Text == "")
                {
                    textBox61.Text = " ";
                    textBox62.Text = " ";
                }
                if (comboBox14.Text == "" || textBox68.Text == "")
                {
                    textBox66.Text = " ";
                    textBox67.Text = " ";
                }
                if (comboBox15.Text == "" || textBox73.Text == "")
                {
                    textBox71.Text = " ";
                    textBox72.Text = " ";
                }
                if (comboBox16.Text == "" || textBox78.Text == "")
                {
                    textBox76.Text = " ";
                    textBox77.Text = " ";
                }
                if (comboBox17.Text == "" || textBox83.Text == "")
                {
                    textBox81.Text = " ";
                    textBox82.Text = " ";
                }
                if (comboBox18.Text == "" || textBox88.Text == "")
                {
                    textBox86.Text = " ";
                    textBox87.Text = " ";
                }
                if (comboBox19.Text == "" || textBox97.Text == "")
                {
                    textBox95.Text = " ";
                    textBox96.Text = " ";
                }
                if (comboBox20.Text == "" || textBox102.Text == "")
                {
                    textBox100.Text = " ";
                    textBox101.Text = " ";
                    progressBar1.Value = 100;
                }
            }
        }

        private void textBox124_TextChanged(object sender, EventArgs e)
        {
            if (textBox124.Text=="A")
            {
                int a = Convert.ToInt32(textBox4.Text);
                var b = 5 * a;
                textBox5.Text = b.ToString();
            }
            if (textBox124.Text == "B")
            {
                int a = Convert.ToInt32(textBox4.Text);
                var b = 4 * a;
                textBox5.Text = b.ToString();
            }
            if (textBox124.Text == "C")
            {
                int a = Convert.ToInt32(textBox4.Text);
                var b = 3 * a;
                textBox5.Text = b.ToString();
            }
            if (textBox124.Text == "D")
            {
                int a = Convert.ToInt32(textBox4.Text);
                var b = 2 * a;
                textBox5.Text = b.ToString();
            }
            if (textBox124.Text == "E")
            {
                int a = Convert.ToInt32(textBox4.Text);
                var b = 1 * a;
                textBox5.Text = b.ToString();
            }
            if (textBox124.Text == "F")
            {
                int a = Convert.ToInt32(textBox4.Text);
                var b = 0 * a;
                textBox5.Text = b.ToString();
            }
        }

        private void textBox123_TextChanged(object sender, EventArgs e)
        {
            if (textBox123.Text == "A")
            {
                int a = Convert.ToInt32(textBox7.Text);
                var b = 5 * a;
                textBox6.Text = b.ToString();
            }
            if (textBox123.Text == "B")
            {
                int a = Convert.ToInt32(textBox7.Text);
                var b = 4 * a;
                textBox6.Text = b.ToString();
            }
            if (textBox123.Text == "C")
            {
                int a = Convert.ToInt32(textBox7.Text);
                var b = 3 * a;
                textBox6.Text = b.ToString();
            }
            if (textBox123.Text == "D")
            {
                int a = Convert.ToInt32(textBox7.Text);
                var b = 2 * a;
                textBox6.Text = b.ToString();
            }
            if (textBox123.Text == "E")
            {
                int a = Convert.ToInt32(textBox7.Text);
                var b = 1 * a;
                textBox6.Text = b.ToString();
            }
            if (textBox123.Text == "F")
            {
                int a = Convert.ToInt32(textBox7.Text);
                var b = 0 * a;
                textBox6.Text = b.ToString();
            }
        }

        private void textBox122_TextChanged(object sender, EventArgs e)
        {
            if (textBox122.Text == "A")
            {
                int a = Convert.ToInt32(textBox12.Text);
                var b = 5 * a;
                textBox11.Text = b.ToString();
            }
            if (textBox122.Text == "B")
            {
                int a = Convert.ToInt32(textBox12.Text);
                var b = 4 * a;
                textBox11.Text = b.ToString();
            }
            if (textBox122.Text == "C")
            {
                int a = Convert.ToInt32(textBox12.Text);
                var b = 3 * a;
                textBox11.Text = b.ToString();
            }
            if (textBox122.Text == "D")
            {
                int a = Convert.ToInt32(textBox12.Text);
                var b = 2 * a;
                textBox11.Text = b.ToString();
            }
            if (textBox122.Text == "E")
            {
                int a = Convert.ToInt32(textBox12.Text);
                var b = 1 * a;
                textBox11.Text = b.ToString();
            }
            if (textBox122.Text == "F")
            {
                int a = Convert.ToInt32(textBox12.Text);
                var b = 0 * a;
                textBox11.Text = b.ToString();
            }
        }

        private void textBox121_TextChanged(object sender, EventArgs e)
        {
            if (textBox121.Text == "A")
            {
                int a = Convert.ToInt32(textBox17.Text);
                var b = 5 * a;
                textBox16.Text = b.ToString();
            }
            if (textBox121.Text == "B")
            {
                int a = Convert.ToInt32(textBox17.Text);
                var b = 4 * a;
                textBox16.Text = b.ToString();
            }
            if (textBox121.Text == "C")
            {
                int a = Convert.ToInt32(textBox17.Text);
                var b = 3 * a;
                textBox16.Text = b.ToString();
            }
            if (textBox121.Text == "D")
            {
                int a = Convert.ToInt32(textBox17.Text);
                var b = 2 * a;
                textBox16.Text = b.ToString();
            }
            if (textBox121.Text == "E")
            {
                int a = Convert.ToInt32(textBox17.Text);
                var b = 1 * a;
                textBox16.Text = b.ToString();
            }
            if (textBox121.Text == "F")
            {
                int a = Convert.ToInt32(textBox17.Text);
                var b = 0 * a;
                textBox16.Text = b.ToString();
            }
        }

        private void textBox120_TextChanged(object sender, EventArgs e)
        {
            if (textBox120.Text == "A")
            {
                int a = Convert.ToInt32(textBox22.Text);
                var b = 5 * a;
                textBox21.Text = b.ToString();
            }
            if (textBox120.Text == "B")
            {
                int a = Convert.ToInt32(textBox22.Text);
                var b = 4 * a;
                textBox21.Text = b.ToString();
            }
            if (textBox120.Text == "C")
            {
                int a = Convert.ToInt32(textBox22.Text);
                var b = 3 * a;
                textBox21.Text = b.ToString();
            }
            if (textBox120.Text == "D")
            {
                int a = Convert.ToInt32(textBox22.Text);
                var b = 2 * a;
                textBox21.Text = b.ToString();
            }
            if (textBox120.Text == "E")
            {
                int a = Convert.ToInt32(textBox22.Text);
                var b = 1 * a;
                textBox21.Text = b.ToString();
            }
            if (textBox120.Text == "F")
            {
                int a = Convert.ToInt32(textBox22.Text);
                var b = 0 * a;
                textBox21.Text = b.ToString();
            }
        }

        private void textBox119_TextChanged(object sender, EventArgs e)
        {
            if (textBox119.Text == "A")
            {
                int a = Convert.ToInt32(textBox27.Text);
                var b = 5 * a;
                textBox26.Text = b.ToString();
            }
            if (textBox119.Text == "B")
            {
                int a = Convert.ToInt32(textBox27.Text);
                var b = 4 * a;
                textBox26.Text = b.ToString();
            }
            if (textBox119.Text == "C")
            {
                int a = Convert.ToInt32(textBox27.Text);
                var b = 3 * a;
                textBox26.Text = b.ToString();
            }
            if (textBox119.Text == "D")
            {
                int a = Convert.ToInt32(textBox27.Text);
                var b = 2 * a;
                textBox26.Text = b.ToString();
            }
            if (textBox119.Text == "E")
            {
                int a = Convert.ToInt32(textBox27.Text);
                var b = 1 * a;
                textBox26.Text = b.ToString();
            }
            if (textBox119.Text == "F")
            {
                int a = Convert.ToInt32(textBox27.Text);
                var b = 0 * a;
                textBox26.Text = b.ToString();
            }
        }

        private void textBox118_TextChanged(object sender, EventArgs e)
        {
            if (textBox118.Text == "A")
            {
                int a = Convert.ToInt32(textBox32.Text);
                var b = 5 * a;
                textBox31.Text = b.ToString();
            }
            if (textBox118.Text == "B")
            {
                int a = Convert.ToInt32(textBox32.Text);
                var b = 4 * a;
                textBox31.Text = b.ToString();
            }
            if (textBox118.Text == "C")
            {
                int a = Convert.ToInt32(textBox32.Text);
                var b = 3 * a;
                textBox31.Text = b.ToString();
            }
            if (textBox118.Text == "D")
            {
                int a = Convert.ToInt32(textBox32.Text);
                var b = 2 * a;
                textBox31.Text = b.ToString();
            }
            if (textBox118.Text == "E")
            {
                int a = Convert.ToInt32(textBox32.Text);
                var b = 1 * a;
                textBox31.Text = b.ToString();
            }
            if (textBox118.Text == "F")
            {
                int a = Convert.ToInt32(textBox32.Text);
                var b = 0 * a;
                textBox31.Text = b.ToString();
            }
        }

        private void textBox117_TextChanged(object sender, EventArgs e)
        {
            if (textBox117.Text == "A")
            {
                int a = Convert.ToInt32(textBox37.Text);
                var b = 5 * a;
                textBox36.Text = b.ToString();
            }
            if (textBox117.Text == "B")
            {
                int a = Convert.ToInt32(textBox37.Text);
                var b = 4 * a;
                textBox36.Text = b.ToString();
            }
            if (textBox117.Text == "C")
            {
                int a = Convert.ToInt32(textBox37.Text);
                var b = 3 * a;
                textBox36.Text = b.ToString();
            }
            if (textBox117.Text == "D")
            {
                int a = Convert.ToInt32(textBox37.Text);
                var b = 2 * a;
                textBox36.Text = b.ToString();
            }
            if (textBox117.Text == "E")
            {
                int a = Convert.ToInt32(textBox37.Text);
                var b = 1 * a;
                textBox36.Text = b.ToString();
            }
            if (textBox117.Text == "F")
            {
                int a = Convert.ToInt32(textBox37.Text);
                var b = 0 * a;
                textBox36.Text = b.ToString();
            }
        }

        private void textBox116_TextChanged(object sender, EventArgs e)
        {
            if (textBox116.Text == "A")
            {
                int a = Convert.ToInt32(textBox42.Text);
                var b = 5 * a;
                textBox41.Text = b.ToString();
            }
            if (textBox116.Text == "B")
            {
                int a = Convert.ToInt32(textBox42.Text);
                var b = 4 * a;
                textBox41.Text = b.ToString();
            }
            if (textBox116.Text == "C")
            {
                int a = Convert.ToInt32(textBox42.Text);
                var b = 3 * a;
                textBox41.Text = b.ToString();
            }
            if (textBox116.Text == "D")
            {
                int a = Convert.ToInt32(textBox42.Text);
                var b = 2 * a;
                textBox41.Text = b.ToString();
            }
            if (textBox116.Text == "E")
            {
                int a = Convert.ToInt32(textBox42.Text);
                var b = 1 * a;
                textBox41.Text = b.ToString();
            }
            if (textBox116.Text == "F")
            {
                int a = Convert.ToInt32(textBox42.Text);
                var b = 0 * a;
                textBox41.Text = b.ToString();
            }
        }

        private void textBox115_TextChanged(object sender, EventArgs e)
        {
            if (textBox115.Text == "A")
            {
                int a = Convert.ToInt32(textBox47.Text);
                var b = 5 * a;
                textBox46.Text = b.ToString();
            }
            if (textBox115.Text == "B")
            {
                int a = Convert.ToInt32(textBox47.Text);
                var b = 4 * a;
                textBox46.Text = b.ToString();
            }
            if (textBox115.Text == "C")
            {
                int a = Convert.ToInt32(textBox47.Text);
                var b = 3 * a;
                textBox46.Text = b.ToString();
            }
            if (textBox115.Text == "D")
            {
                int a = Convert.ToInt32(textBox47.Text);
                var b = 2 * a;
                textBox46.Text = b.ToString();
            }
            if (textBox115.Text == "E")
            {
                int a = Convert.ToInt32(textBox47.Text);
                var b = 1 * a;
                textBox46.Text = b.ToString();
            }
            if (textBox115.Text == "F")
            {
                int a = Convert.ToInt32(textBox47.Text);
                var b = 0 * a;
                textBox46.Text = b.ToString();
            }
        }

        private void textBox114_TextChanged(object sender, EventArgs e)
        {
            if (textBox114.Text == "A")
            {
                int a = Convert.ToInt32(textBox52.Text);
                var b = 5 * a;
                textBox51.Text = b.ToString();
            }
            if (textBox114.Text == "B")
            {
                int a = Convert.ToInt32(textBox52.Text);
                var b = 4 * a;
                textBox51.Text = b.ToString();
            }
            if (textBox114.Text == "C")
            {
                int a = Convert.ToInt32(textBox52.Text);
                var b = 3 * a;
                textBox51.Text = b.ToString();
            }
            if (textBox114.Text == "D")
            {
                int a = Convert.ToInt32(textBox52.Text);
                var b = 2 * a;
                textBox51.Text = b.ToString();
            }
            if (textBox114.Text == "E")
            {
                int a = Convert.ToInt32(textBox52.Text);
                var b = 1 * a;
                textBox51.Text = b.ToString();
            }
            if (textBox114.Text == "F")
            {
                int a = Convert.ToInt32(textBox52.Text);
                var b = 0 * a;
                textBox51.Text = b.ToString();
            }
        }

        private void textBox110_TextChanged(object sender, EventArgs e)
        {
            if (textBox110.Text == "A")
            {
                int a = Convert.ToInt32(textBox57.Text);
                var b = 5 * a;
                textBox56.Text = b.ToString();
            }
            if (textBox110.Text == "B")
            {
                int a = Convert.ToInt32(textBox57.Text);
                var b = 4 * a;
                textBox56.Text = b.ToString();
            }
            if (textBox110.Text == "C")
            {
                int a = Convert.ToInt32(textBox57.Text);
                var b = 3 * a;
                textBox56.Text = b.ToString();
            }
            if (textBox110.Text == "D")
            {
                int a = Convert.ToInt32(textBox57.Text);
                var b = 2 * a;
                textBox56.Text = b.ToString();
            }
            if (textBox110.Text == "E")
            {
                int a = Convert.ToInt32(textBox57.Text);
                var b = 1 * a;
                textBox56.Text = b.ToString();
            }
            if (textBox110.Text == "F")
            {
                int a = Convert.ToInt32(textBox57.Text);
                var b = 0 * a;
                textBox56.Text = b.ToString();
            }
        }

        private void textBox111_TextChanged(object sender, EventArgs e)
        {
            if (textBox111.Text == "A")
            {
                int a = Convert.ToInt32(textBox62.Text);
                var b = 5 * a;
                textBox61.Text = b.ToString();
            }
            if (textBox111.Text == "B")
            {
                int a = Convert.ToInt32(textBox62.Text);
                var b = 4 * a;
                textBox61.Text = b.ToString();
            }
            if (textBox111.Text == "C")
            {
                int a = Convert.ToInt32(textBox62.Text);
                var b = 3 * a;
                textBox61.Text = b.ToString();
            }
            if (textBox111.Text == "D")
            {
                int a = Convert.ToInt32(textBox62.Text);
                var b = 2 * a;
                textBox61.Text = b.ToString();
            }
            if (textBox111.Text == "E")
            {
                int a = Convert.ToInt32(textBox62.Text);
                var b = 1 * a;
                textBox61.Text = b.ToString();
            }
            if (textBox111.Text == "F")
            {
                int a = Convert.ToInt32(textBox62.Text);
                var b = 0 * a;
                textBox61.Text = b.ToString();
            }
        }

        private void textBox112_TextChanged(object sender, EventArgs e)
        {
            if (textBox112.Text == "A")
            {
                int a = Convert.ToInt32(textBox67.Text);
                var b = 5 * a;
                textBox66.Text = b.ToString();
            }
            if (textBox112.Text == "B")
            {
                int a = Convert.ToInt32(textBox67.Text);
                var b = 4 * a;
                textBox66.Text = b.ToString();
            }
            if (textBox112.Text == "C")
            {
                int a = Convert.ToInt32(textBox67.Text);
                var b = 3 * a;
                textBox66.Text = b.ToString();
            }
            if (textBox112.Text == "D")
            {
                int a = Convert.ToInt32(textBox67.Text);
                var b = 2 * a;
                textBox66.Text = b.ToString();
            }
            if (textBox112.Text == "E")
            {
                int a = Convert.ToInt32(textBox67.Text);
                var b = 1 * a;
                textBox66.Text = b.ToString();
            }
            if (textBox112.Text == "F")
            {
                int a = Convert.ToInt32(textBox67.Text);
                var b = 0 * a;
                textBox66.Text = b.ToString();
            }
        }

        private void textBox113_TextChanged(object sender, EventArgs e)
        {
            if (textBox113.Text == "A")
            {
                int a = Convert.ToInt32(textBox72.Text);
                var b = 5 * a;
                textBox71.Text = b.ToString();
            }
            if (textBox113.Text == "B")
            {
                int a = Convert.ToInt32(textBox72.Text);
                var b = 4 * a;
                textBox71.Text = b.ToString();
            }
            if (textBox113.Text == "C")
            {
                int a = Convert.ToInt32(textBox72.Text);
                var b = 3 * a;
                textBox71.Text = b.ToString();
            }
            if (textBox113.Text == "D")
            {
                int a = Convert.ToInt32(textBox72.Text);
                var b = 2 * a;
                textBox71.Text = b.ToString();
            }
            if (textBox113.Text == "E")
            {
                int a = Convert.ToInt32(textBox72.Text);
                var b = 1 * a;
                textBox71.Text = b.ToString();
            }
            if (textBox113.Text == "F")
            {
                int a = Convert.ToInt32(textBox72.Text);
                var b = 0 * a;
                textBox71.Text = b.ToString();
            }
        }

        private void textBox109_TextChanged(object sender, EventArgs e)
        {
            if (textBox109.Text == "A")
            {
                int a = Convert.ToInt32(textBox77.Text);
                var b = 5 * a;
                textBox76.Text = b.ToString();
            }
            if (textBox109.Text == "B")
            {
                int a = Convert.ToInt32(textBox77.Text);
                var b = 4 * a;
                textBox76.Text = b.ToString();
            }
            if (textBox109.Text == "C")
            {
                int a = Convert.ToInt32(textBox77.Text);
                var b = 3 * a;
                textBox76.Text = b.ToString();
            }
            if (textBox109.Text == "D")
            {
                int a = Convert.ToInt32(textBox77.Text);
                var b = 2 * a;
                textBox76.Text = b.ToString();
            }
            if (textBox109.Text == "E")
            {
                int a = Convert.ToInt32(textBox77.Text);
                var b = 1 * a;
                textBox76.Text = b.ToString();
            }
            if (textBox109.Text == "F")
            {
                int a = Convert.ToInt32(textBox77.Text);
                var b = 0 * a;
                textBox76.Text = b.ToString();
            }
        }

        private void textBox108_TextChanged(object sender, EventArgs e)
        {
            if (textBox108.Text == "A")
            {
                int a = Convert.ToInt32(textBox82.Text);
                var b = 5 * a;
                textBox81.Text = b.ToString();
            }
            if (textBox108.Text == "B")
            {
                int a = Convert.ToInt32(textBox82.Text);
                var b = 4 * a;
                textBox81.Text = b.ToString();
            }
            if (textBox108.Text == "C")
            {
                int a = Convert.ToInt32(textBox82.Text);
                var b = 3 * a;
                textBox81.Text = b.ToString();
            }
            if (textBox108.Text == "D")
            {
                int a = Convert.ToInt32(textBox82.Text);
                var b = 2 * a;
                textBox81.Text = b.ToString();
            }
            if (textBox108.Text == "E")
            {
                int a = Convert.ToInt32(textBox82.Text);
                var b = 1 * a;
                textBox81.Text = b.ToString();
            }
            if (textBox108.Text == "F")
            {
                int a = Convert.ToInt32(textBox82.Text);
                var b = 0 * a;
                textBox81.Text = b.ToString();
            }
        }

        private void textBox107_TextChanged(object sender, EventArgs e)
        {
            if (textBox107.Text == "A")
            {
                int a = Convert.ToInt32(textBox87.Text);
                var b = 5 * a;
                textBox86.Text = b.ToString();
            }
            if (textBox107.Text == "B")
            {
                int a = Convert.ToInt32(textBox87.Text);
                var b = 4 * a;
                textBox86.Text = b.ToString();
            }
            if (textBox107.Text == "C")
            {
                int a = Convert.ToInt32(textBox87.Text);
                var b = 3 * a;
                textBox86.Text = b.ToString();
            }
            if (textBox107.Text == "D")
            {
                int a = Convert.ToInt32(textBox87.Text);
                var b = 2 * a;
                textBox86.Text = b.ToString();
            }
            if (textBox107.Text == "E")
            {
                int a = Convert.ToInt32(textBox87.Text);
                var b = 1 * a;
                textBox86.Text = b.ToString();
            }
            if (textBox107.Text == "F")
            {
                int a = Convert.ToInt32(textBox87.Text);
                var b = 0 * a;
                textBox86.Text = b.ToString();
            }
        }

        private void textBox106_TextChanged(object sender, EventArgs e)
        {
            if (textBox106.Text == "A")
            {
                int a = Convert.ToInt32(textBox96.Text);
                var b = 5 * a;
                textBox95.Text = b.ToString();
            }
            if (textBox106.Text == "B")
            {
                int a = Convert.ToInt32(textBox96.Text);
                var b = 4 * a;
                textBox95.Text = b.ToString();
            }
            if (textBox106.Text == "C")
            {
                int a = Convert.ToInt32(textBox96.Text);
                var b = 3 * a;
                textBox95.Text = b.ToString();
            }
            if (textBox106.Text == "D")
            {
                int a = Convert.ToInt32(textBox96.Text);
                var b = 2 * a;
                textBox95.Text = b.ToString();
            }
            if (textBox106.Text == "E")
            {
                int a = Convert.ToInt32(textBox96.Text);
                var b = 1 * a;
                textBox95.Text = b.ToString();
            }
            if (textBox106.Text == "F")
            {
                int a = Convert.ToInt32(textBox96.Text);
                var b = 0 * a;
                textBox95.Text = b.ToString();
            }
        }

        private void textBox93_TextChanged(object sender, EventArgs e)
        {
            if (textBox93.Text == "A")
            {
                int a = Convert.ToInt32(textBox101.Text);
                var b = 5 * a;
                textBox100.Text = b.ToString();
            }
            if (textBox93.Text == "B")
            {
                int a = Convert.ToInt32(textBox101.Text);
                var b = 4 * a;
                textBox100.Text = b.ToString();
            }
            if (textBox93.Text == "C")
            {
                int a = Convert.ToInt32(textBox101.Text);
                var b = 3 * a;
                textBox100.Text = b.ToString();
            }
            if (textBox93.Text == "D")
            {
                int a = Convert.ToInt32(textBox101.Text);
                var b = 2 * a;
                textBox100.Text = b.ToString();
            }
            if (textBox93.Text == "E")
            {
                int a = Convert.ToInt32(textBox101.Text);
                var b = 1 * a;
                textBox100.Text = b.ToString();
            }
            if (textBox93.Text == "F")
            {
                int a = Convert.ToInt32(textBox101.Text);
                var b = 0 * a;
                textBox100.Text = b.ToString();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox1.Text+".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox4.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox2.Text+".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox7.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox3.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox12.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox4.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox17.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox5.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox22.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox6.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox27.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox7.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox32.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox8.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox37.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox9.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox42.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox10.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox47.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox11.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox52.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox12.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox57.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox13.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox62.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox14.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox67.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox15.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox72.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox16.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox77.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox17.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox82.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox18.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox87.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox19.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox96.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void comboBox20_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            filepath1 = filepath1 + @"\CGPA FILES\";
            var course = filepath1 + comboBox20.Text + ".txt";
            string name = Path.Combine(course);
            System.IO.StreamReader objreader;
            objreader = new System.IO.StreamReader(name);
            textBox101.Text = objreader.ReadToEnd();
            objreader.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox91.Text=="")
            {
                MessageBox.Show("Please Insert New Course Code");
            }
            else
                if (textBox94.Text == "")
                {
                    MessageBox.Show("Please Insert New Course Unit");
                }
                else
                {
                    progressBar1.Value = 0;
                    string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    filepath1 = filepath1 + @"\CGPA FILES\";
                    if (!Directory.Exists(filepath1))
                    {
                        Directory.CreateDirectory(filepath1);
                    }
                    File.Delete(filepath1 + "DEFAULT.txt");
                    File.WriteAllText(Path.Combine(filepath1, textBox91.Text + ".txt"), textBox94.Text);
                    comboBox1.Items.Clear();
                    comboBox2.Items.Clear();
                    comboBox3.Items.Clear();
                    comboBox4.Items.Clear();
                    comboBox5.Items.Clear();
                    comboBox6.Items.Clear();
                    comboBox7.Items.Clear();
                    comboBox8.Items.Clear();
                    comboBox9.Items.Clear();
                    comboBox10.Items.Clear();
                    comboBox11.Items.Clear();
                    comboBox12.Items.Clear();
                    comboBox13.Items.Clear();
                    comboBox14.Items.Clear();
                    comboBox15.Items.Clear();
                    comboBox16.Items.Clear();
                    comboBox17.Items.Clear();
                    comboBox18.Items.Clear();
                    comboBox19.Items.Clear();
                    comboBox20.Items.Clear();
                    string[] files = Directory.GetFiles(filepath1, "*.txt", SearchOption.AllDirectories);
                    var namesonly = files.Select(f => Path.GetFileNameWithoutExtension(f)).ToArray();
                    comboBox1.Items.AddRange(namesonly);
                    comboBox2.Items.AddRange(namesonly);
                    comboBox3.Items.AddRange(namesonly);
                    comboBox4.Items.AddRange(namesonly);
                    comboBox5.Items.AddRange(namesonly);
                    comboBox6.Items.AddRange(namesonly);
                    comboBox7.Items.AddRange(namesonly);
                    comboBox8.Items.AddRange(namesonly);
                    comboBox9.Items.AddRange(namesonly);
                    comboBox10.Items.AddRange(namesonly);
                    comboBox11.Items.AddRange(namesonly);
                    comboBox12.Items.AddRange(namesonly);
                    comboBox13.Items.AddRange(namesonly);
                    comboBox14.Items.AddRange(namesonly);
                    comboBox15.Items.AddRange(namesonly);
                    comboBox16.Items.AddRange(namesonly);
                    comboBox17.Items.AddRange(namesonly);
                    comboBox18.Items.AddRange(namesonly);
                    comboBox19.Items.AddRange(namesonly);
                    comboBox20.Items.AddRange(namesonly);
                    textBox91.Text = "";
                    textBox94.Text = "";
                    progressBar1.Value = 100;
                }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(textBox127.Text=="")
            {
                MessageBox.Show("Please Insert Your Registration Number");
            }
            else
            if(textBox126.Text=="")
            {
                MessageBox.Show("Please Insert Your Semester");
            }
            else
            if (textBox125.Text == "")
            {
                MessageBox.Show("Please Insert Your Level");
            }
            else
            {
                progressBar1.Value = 0;
                var doc1 = new Document();
                string filepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                filepath = filepath + @"\CGPA RECORDS\" + textBox127.Text.ToString() + "/";
                if (!Directory.Exists(filepath))
                {
                    Directory.CreateDirectory(filepath);
                }
                string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                filepath1 = filepath1 + @"\CGPA FILES\" + textBox127.Text.ToString() + "/";
                PdfWriter.GetInstance(doc1, new FileStream(filepath + textBox126.Text + " " + textBox125.Text + ".pdf", FileMode.Create));
                doc1.Open();
                doc1.Add(new Paragraph(label14.Text + " " + textBox127.Text));
                doc1.Add(new Paragraph(label12.Text + " " + textBox125.Text));
                doc1.Add(new Paragraph(label13.Text + " " + textBox126.Text));
                doc1.Add(new Paragraph(label8.Text + " " + textBox105.Text));
                doc1.Add(new Paragraph(label7.Text + " " + textBox92.Text));
                doc1.Add(new Paragraph("                          "));
                doc1.Add(new Paragraph("                          "));
                PdfPTable table = new PdfPTable(7);
                table.TotalWidth = 520f;
                table.LockedWidth = true;
                table.HorizontalAlignment = 0;
                table.AddCell(label1.Text);
                table.AddCell(label2.Text);
                table.AddCell(label3.Text);
                table.AddCell(label5.Text);
                table.AddCell(label11.Text);
                table.AddCell(label4.Text);
                table.AddCell(label6.Text);

                if (comboBox1.Text == "" || textBox3.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox1.Text);
                table.AddCell(textBox1.Text);
                table.AddCell(textBox2.Text);
                table.AddCell(textBox3.Text);
                table.AddCell(textBox124.Text);
                table.AddCell(textBox4.Text);
                table.AddCell(textBox5.Text);

                if (comboBox2.Text == "" || textBox8.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox2.Text);
                table.AddCell(textBox10.Text);
                table.AddCell(textBox9.Text);
                table.AddCell(textBox8.Text);
                table.AddCell(textBox123.Text);
                table.AddCell(textBox7.Text);
                table.AddCell(textBox6.Text);

                if (comboBox3.Text == "" || textBox13.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox3.Text);
                table.AddCell(textBox15.Text);
                table.AddCell(textBox14.Text);
                table.AddCell(textBox13.Text);
                table.AddCell(textBox122.Text);
                table.AddCell(textBox12.Text);
                table.AddCell(textBox11.Text);

                if (comboBox4.Text == "" || textBox18.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox4.Text);
                table.AddCell(textBox20.Text);
                table.AddCell(textBox19.Text);
                table.AddCell(textBox18.Text);
                table.AddCell(textBox121.Text);
                table.AddCell(textBox17.Text);
                table.AddCell(textBox16.Text);

                if (comboBox5.Text == "" || textBox23.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox5.Text);
                table.AddCell(textBox25.Text);
                table.AddCell(textBox24.Text);
                table.AddCell(textBox23.Text);
                table.AddCell(textBox120.Text);
                table.AddCell(textBox22.Text);
                table.AddCell(textBox21.Text);

                if (comboBox6.Text == "" || textBox28.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox6.Text);
                table.AddCell(textBox30.Text);
                table.AddCell(textBox29.Text);
                table.AddCell(textBox28.Text);
                table.AddCell(textBox119.Text);
                table.AddCell(textBox27.Text);
                table.AddCell(textBox26.Text);

                if (comboBox7.Text == "" || textBox33.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox7.Text);
                table.AddCell(textBox35.Text);
                table.AddCell(textBox34.Text);
                table.AddCell(textBox33.Text);
                table.AddCell(textBox118.Text);
                table.AddCell(textBox32.Text);
                table.AddCell(textBox31.Text);

                if (comboBox8.Text == "" || textBox38.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox8.Text);
                table.AddCell(textBox40.Text);
                table.AddCell(textBox39.Text);
                table.AddCell(textBox38.Text);
                table.AddCell(textBox117.Text);
                table.AddCell(textBox37.Text);
                table.AddCell(textBox36.Text);

                if (comboBox9.Text == "" || textBox43.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox9.Text);
                table.AddCell(textBox45.Text);
                table.AddCell(textBox44.Text);
                table.AddCell(textBox43.Text);
                table.AddCell(textBox116.Text);
                table.AddCell(textBox42.Text);
                table.AddCell(textBox41.Text);

                if (comboBox10.Text == "" || textBox48.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox10.Text);
                table.AddCell(textBox50.Text);
                table.AddCell(textBox49.Text);
                table.AddCell(textBox48.Text);
                table.AddCell(textBox115.Text);
                table.AddCell(textBox47.Text);
                table.AddCell(textBox46.Text);

                if (comboBox11.Text == "" || textBox53.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox11.Text);
                table.AddCell(textBox55.Text);
                table.AddCell(textBox54.Text);
                table.AddCell(textBox53.Text);
                table.AddCell(textBox114.Text);
                table.AddCell(textBox52.Text);
                table.AddCell(textBox51.Text);

                if (comboBox12.Text == "" || textBox58.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox12.Text);
                table.AddCell(textBox60.Text);
                table.AddCell(textBox59.Text);
                table.AddCell(textBox58.Text);
                table.AddCell(textBox110.Text);
                table.AddCell(textBox57.Text);
                table.AddCell(textBox56.Text);

                if (comboBox13.Text == "" || textBox63.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox13.Text);
                table.AddCell(textBox65.Text);
                table.AddCell(textBox64.Text);
                table.AddCell(textBox63.Text);
                table.AddCell(textBox111.Text);
                table.AddCell(textBox62.Text);
                table.AddCell(textBox61.Text);

                if (comboBox14.Text == "" || textBox68.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox14.Text);
                table.AddCell(textBox70.Text);
                table.AddCell(textBox69.Text);
                table.AddCell(textBox68.Text);
                table.AddCell(textBox112.Text);
                table.AddCell(textBox67.Text);
                table.AddCell(textBox66.Text);

                if (comboBox15.Text == "" || textBox73.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox15.Text);
                table.AddCell(textBox75.Text);
                table.AddCell(textBox74.Text);
                table.AddCell(textBox73.Text);
                table.AddCell(textBox113.Text);
                table.AddCell(textBox72.Text);
                table.AddCell(textBox71.Text);

                if (comboBox16.Text == "" || textBox78.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox16.Text);
                table.AddCell(textBox80.Text);
                table.AddCell(textBox79.Text);
                table.AddCell(textBox78.Text);
                table.AddCell(textBox109.Text);
                table.AddCell(textBox77.Text);
                table.AddCell(textBox76.Text);

                if (comboBox17.Text == "" || textBox83.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox17.Text);
                table.AddCell(textBox85.Text);
                table.AddCell(textBox84.Text);
                table.AddCell(textBox83.Text);
                table.AddCell(textBox108.Text);
                table.AddCell(textBox82.Text);
                table.AddCell(textBox81.Text);

                if (comboBox18.Text == "" || textBox88.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox18.Text);
                table.AddCell(textBox90.Text);
                table.AddCell(textBox89.Text);
                table.AddCell(textBox88.Text);
                table.AddCell(textBox107.Text);
                table.AddCell(textBox87.Text);
                table.AddCell(textBox86.Text);

                if (comboBox19.Text == "" || textBox97.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox19.Text);
                table.AddCell(textBox99.Text);
                table.AddCell(textBox98.Text);
                table.AddCell(textBox97.Text);
                table.AddCell(textBox106.Text);
                table.AddCell(textBox96.Text);
                table.AddCell(textBox95.Text);

                if (comboBox20.Text == "" || textBox102.Text == "")
                {
                    try
                    {
                        doc1.Add(table);
                        doc1.Close();
                    }
                    catch (DocumentException)
                    {

                    }
                }
                table.AddCell(comboBox20.Text);
                table.AddCell(textBox104.Text);
                table.AddCell(textBox103.Text);
                table.AddCell(textBox102.Text);
                table.AddCell(textBox93.Text);
                table.AddCell(textBox101.Text);
                table.AddCell(textBox100.Text);
                try
                {
                    doc1.Add(table);
                    doc1.Close();
                }
                catch (DocumentException)
                {

                }
                progressBar1.Value = 100;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "")
            {
                MessageBox.Show("Please Select Course Code Using the First Course Code Select List to Delete");
            }
            else
            {
                progressBar1.Value = 0;
                string filepath1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                filepath1 = filepath1 + @"\CGPA FILES\";
                File.Delete(filepath1 + comboBox1.Text + ".txt");
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                comboBox3.Items.Clear();
                comboBox4.Items.Clear();
                comboBox5.Items.Clear();
                comboBox6.Items.Clear();
                comboBox7.Items.Clear();
                comboBox8.Items.Clear();
                comboBox9.Items.Clear();
                comboBox10.Items.Clear();
                comboBox11.Items.Clear();
                comboBox12.Items.Clear();
                comboBox13.Items.Clear();
                comboBox14.Items.Clear();
                comboBox15.Items.Clear();
                comboBox16.Items.Clear();
                comboBox17.Items.Clear();
                comboBox18.Items.Clear();
                comboBox19.Items.Clear();
                comboBox20.Items.Clear();
                string[] files = Directory.GetFiles(filepath1, "*.txt", SearchOption.AllDirectories);
                var namesonly = files.Select(f => Path.GetFileNameWithoutExtension(f)).ToArray();
                comboBox1.Items.AddRange(namesonly);
                comboBox2.Items.AddRange(namesonly);
                comboBox3.Items.AddRange(namesonly);
                comboBox4.Items.AddRange(namesonly);
                comboBox5.Items.AddRange(namesonly);
                comboBox6.Items.AddRange(namesonly);
                comboBox7.Items.AddRange(namesonly);
                comboBox8.Items.AddRange(namesonly);
                comboBox9.Items.AddRange(namesonly);
                comboBox10.Items.AddRange(namesonly);
                comboBox11.Items.AddRange(namesonly);
                comboBox12.Items.AddRange(namesonly);
                comboBox13.Items.AddRange(namesonly);
                comboBox14.Items.AddRange(namesonly);
                comboBox15.Items.AddRange(namesonly);
                comboBox16.Items.AddRange(namesonly);
                comboBox17.Items.AddRange(namesonly);
                comboBox18.Items.AddRange(namesonly);
                comboBox19.Items.AddRange(namesonly);
                comboBox20.Items.AddRange(namesonly);
                progressBar1.Value = 100;
            }
        }
    }
}
