using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace sparsh_v2._0
{
    public partial class Form1 : Form
    {
        Image file;
        public Form1()
        {
            InitializeComponent();
            getAvailablePorts();
        }
        void getAvailablePorts()
        {
            String[] ports = SerialPort.GetPortNames();
            comboBox1.Items.AddRange(ports);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.Text == "")
                {
                    MessageBox.Show("please select port number");
                }
                else
                {
                    serialPort1.PortName = comboBox1.Text;

                    serialPort1.Open();
                    progressBar1.Value = 100;
                    button3.Enabled = true;
                    button4.Enabled = true;
                    richTextBox1.Enabled = true;
                    button1.Enabled = false;
                    button2.Enabled = true;

                }
            }
            catch (UnauthorizedAccessException)
            {
                richTextBox1.Text = "Unauthorised Access";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            serialPort1.Close();
            progressBar1.Value = 0;
            button3.Enabled = false;
            button4.Enabled = false;
            richTextBox1.Enabled = false;
            button1.Enabled = true;
            button2.Enabled = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            serialPort1.WriteLine(richTextBox1.Text);
            richTextBox1.Text = "";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                richTextBox1.Text = serialPort1.ReadExisting();
            }
            catch (TimeoutException)
            {
                richTextBox1.Text = "TimeOut Exception";
                
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            Stream myStream;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if ((myStream = openFileDialog1.OpenFile()) != null)
                {
                    string strfilename = openFileDialog1.FileName;
                    string filetext = File.ReadAllText(strfilename);
                    richTextBox1.Text = filetext;
                }
            }
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using (Stream s = File.Open(saveFileDialog1.FileName, FileMode.CreateNew))
                using (StreamWriter sw = new StreamWriter(s))
                {
                    sw.Write(richTextBox1.Text);
                }
            }
            
        }

        private void button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog f = new OpenFileDialog();
            if (f.ShowDialog() == DialogResult.OK)
            {
                file = Image.FromFile(f.FileName);
                pictureBox1.Image = file;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            SaveFileDialog f = new SaveFileDialog();
            if (f.ShowDialog() == DialogResult.OK)
            {
                file.Save(f.FileName);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
           
        }
            
        

        private void button8_Click(object sender, EventArgs e)
        {
            
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
