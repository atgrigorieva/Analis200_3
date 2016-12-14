using System;
using System.Windows.Forms;
using System.Threading;
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
            string indata = _Analis.newPort.ReadExisting();

            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = _Analis.newPort.ReadExisting();
                }
            }
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
