using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.IO.Ports;
using System.Threading.Tasks;
namespace Analis200
{
    public partial class SWForm : Form
    {
        Analis _Analis;
        public string GW1_2;
        public byte[] GWbuffer;
        public SWForm(Analis parent)
        {
            InitializeComponent();
            this._Analis = parent;
            
        }
        
        public void SWForm_Load(object sender, EventArgs e)
        {
            
        }
        public void SW()
        {
            string SWText1 = SWText.Text;
            _Analis.newPort.Write("SW " + SWText.Text + "\r");
            Thread.Sleep(60000);
            int byteRecieved1 = _Analis.newPort.ReadBufferSize;
            Thread.Sleep(1500);
           byte[] buffer1 = new byte[byteRecieved1];
            _Analis.newPort.Read(buffer1, 0, byteRecieved1);
            _Analis.GWNew.Text = SWText.Text;
            
            // _Analis.GW();
        }

        
        
        private void SWButton_Click_1(object sender, EventArgs e)
        {

            SW();
          // _Analis.GW();
            Close();
           
        }

        private void SWForm_FormClosed(object sender, FormClosedEventArgs e)
        {
           // _Analis.GW();
        }
    }
}
