using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SerialPort
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string stroka = "";
        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            this.Invoke(new EventHandler(DoUpdate));
        }

        private void DoUpdate(object s, EventArgs e)
        {
            stroka = stroka + serialPort1.ReadExisting();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            serialPort1.PortName = "COM3"; //Указываем наш порт - в данном случае COM1.
            serialPort1.BaudRate = 9600; //указываем скорость.
            serialPort1.DataBits = 8;
            serialPort1.Open(); //Открываем порт.     


            if (serialPort1.IsOpen)
                MessageBox.Show("Порт успешно открыт");
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            serialPort1.Close(); //Закрываем порт.
            
        }

           
    }
    
}

