using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConsoleApp6
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog2.ShowDialog();
        }

        public void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            string name = openFileDialog1.SafeFileName;
            textBox1.Text=name;

        }

        public void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            string name = openFileDialog2.SafeFileName;
            textBox2.Text= name;
        }

        private void fileSystemWatcher1_Changed(object sender, System.IO.FileSystemEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Entrypoint runapp = new Entrypoint();
            runapp.Start(textBox1.Text,textBox2.Text);
            listBox1_SelectedIndexChanged(null,null);
            
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox1.DataSource = Entrypoint.changelist;
            File.Delete(@"c:\xls\Changes.txt");
            using (File.Create(@"c:\xls\Changes.txt")) ;
            using (TextWriter tw = new StreamWriter(@"c:\xls\Changes.txt"))
            {
                foreach (String s in Entrypoint.changelist)
                    tw.WriteLine(s);
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
