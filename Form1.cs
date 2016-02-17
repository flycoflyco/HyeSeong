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


namespace ClassLibrary1
{

    public partial class Form1 : Form
    {

        public event EventHandler Click2;
        public event EventHandler Click3;
        public event EventHandler Click4;

        Class1 cl;
        testExcel cl2;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) // 간단한 선그리기
        {
            cl2 = new testExcel();
            
        }


//        public event EventHandler Click;
        private void button2_Click(object sender, EventArgs e)
        {
            if (Click2 != null)
            {
                Click2(this, EventArgs.Empty);

            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(Click3 != null)
            {
                Click3(this, EventArgs.Empty);
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (Click4 != null)
            {
                Click4(this, EventArgs.Empty);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
          
        }

        private void update_Click(object sender, EventArgs e)
        {
            cl = new Class1();
            cl.testpara(lv1);
            
        }



    }
}
