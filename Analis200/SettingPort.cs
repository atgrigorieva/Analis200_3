using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Threading;
namespace Analis200
{
    public partial class SettingPort : Form
    {
        //SettingPort _form2 = new SettingPort(this);
        //string name;
        Analis _Analis;

        public SettingPort(Analis parent)
        {

            InitializeComponent();
            this._Analis = parent;
            //CO();
            SW();
            InitializeTimer();
        }
        private int counter;
        private void InitializeTimer()
        {
            // Run this procedure in an appropriate event.
            counter = 0;
            timer1.Interval = 600;
            timer1.Enabled = true;
            // Hook up timer's tick event handler.
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
        }
        private void timer1_Tick(object sender, System.EventArgs e)
        {
            
                GE();
            
        }
       /* public void CO()
        {
            
            string b = "";
            int byteRecieved = _Analis.newPort.ReadBufferSize;
            Thread.Sleep(100);
            byte[] buffer = new byte[byteRecieved];
            _Analis.newPort.Read(buffer, 0, byteRecieved);
            for (int i = 0; i <= 50; i++)
            {
                b = b + Convert.ToChar(buffer[i]);
            }
            var arr = b.Split("\r".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            string D1 = arr[1];
            label1.Text = D1;
            string D2 = arr[2];
            label2.Text = D2;
            string D3 = arr[3];
            label3.Text = D3;
            string D4 = arr[4];
            label4.Text = D4;
            string D5 = arr[5];
            label5.Text = D5;
            string D6 = arr[6];
            label6.Text = D6;
            string D7 = arr[7];
            label7.Text = D7;
        }*/
        public void SW()
        {
            _Analis.newPort.Write("SW 300\r");

            int byteRecieved1 = _Analis.newPort.ReadBufferSize;
            Thread.Sleep(100);
            byte[] buffer1 = new byte[byteRecieved1];
            _Analis.newPort.Read(buffer1, 0, byteRecieved1);

        }
        public void GE()
        {
            _Analis.newPort.Write("GE 1\r");

            int byteRecieved4 = _Analis.newPort.ReadBufferSize;
            Thread.Sleep(100);
            byte[] buffer4 = new byte[byteRecieved4];
            _Analis.newPort.Read(buffer4, 0, byteRecieved4);
            string GE1 = "";
            for (int i = 0; i <= 50; i++)
            {
                GE1 = GE1 + Convert.ToChar(buffer4[i]);
            }
            var arr4 = GE1.Split("\r".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            GE1 = arr4[1];
            textBox3.Text = GE1;

        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

      
        public void SettingPort_Load(object sender, EventArgs e)
        {
            //char[] OpenPribor = { Convert.ToChar('C'), Convert.ToChar('O'), Convert.ToChar('\r') };
            

            
            

        }
       

    } 

}
