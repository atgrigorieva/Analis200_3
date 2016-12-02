using System;
using System.Drawing;
using System.Windows.Forms;
namespace Analis200
{
    public partial class Analiz : Form
    {
        Analis _Analis;
        public int oldValue = 3;
        public Analiz(Analis parent)
        {
            InitializeComponent();
            this._Analis = parent;
        }

        public void Analiz_Load(object sender, EventArgs e)
        {
            for (int i=1; i<=20; i++)
            {
                WLNO1.Items.Add(i);
            }
           
            var height = 118;
            var labelx = 18;
            
            for (int i = 0; i <= 9; i++)
            {
                var label = new Label();
                label.Name = "label" + i++.ToString();
                label.Text = "WL. " + i-- + " (нм)";
                label.Location = new Point(labelx, height);
                height += label.Height;
                this.Controls.Add(label);
            }
            var height1 = 115;
            var textBoxx = 120;
            
            for (int i = 0; i <= 9; i++) {
                _Analis.textBox[i] = new TextBox();
                _Analis.textBox[i].Name = "textbox" + i++.ToString();
                i--;
                _Analis.textBox[i].Text = Convert.ToString(_Analis.textBox[i].Name);
                _Analis.textBox[i].Width = 121;
                _Analis.textBox[i].Height = 20;
                _Analis.textBox[i].Location = new Point(textBoxx, height1);
                height1 += _Analis.textBox[i].Height+3;
                _Analis.textBox[i].Enabled = false;
                this.Controls.Add(_Analis.textBox[i]);

            }
            
            var labelx1 = 400;
            var height2 = 118;
            
            for (int i = 10; i <= 19; i++)
            {
                var label = new Label();
                label.Name = "label" + i++.ToString();
                label.Text = "WL. " + i-- + " (нм)";
                label.Location = new Point(labelx1, height2);
                height2 += label.Height;
                this.Controls.Add(label);
            }
            var height3 = 115;
            var textBoxx1 = 502;
            for (int i = 10; i <= 19; i++)
            {
                _Analis.textBox[i] = new TextBox();
                _Analis.textBox[i].Name = "textbox" + i.ToString();
                _Analis.textBox[i].Text = Convert.ToString(_Analis.textBox[i].Name);
                _Analis.textBox[i].Width = 121;
                _Analis.textBox[i].Height = 20;
                _Analis.textBox[i].Location = new Point(textBoxx1, height3);
                height3 += _Analis.textBox[i].Height+3;
                _Analis.textBox[i].Enabled = false;
                this.Controls.Add(_Analis.textBox[i]);
                

            }
            WLNO1.SelectedIndex = 2;
            for (int i = Convert.ToInt32(WLNO1.SelectedIndex) - 1; i >= 0; i--)
            {
                this._Analis.textBox[i].Enabled = true;

            }
        }

        private void WLNO1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            _Analis.countWL = Convert.ToInt32(WLNO1.SelectedItem);
            
            if (this.oldValue > _Analis.countWL)
            {

                for (int i1 = 0; i1 <= 19; i1++)
                {
                    _Analis.textBox[i1].Enabled = false;
                }

                for (int i = _Analis.countWL - 1; i >= 0; i--)
                {
                    _Analis.textBox[i].Enabled = true;

                }
            }
            else
            {
                for (int i = _Analis.countWL - 1; i >= 1; i--)
                {
                    _Analis.textBox[i].Enabled = true;

                }
            }
            oldValue = _Analis.countWL;
        
    }

        public void WL()
        {
            _Analis.WLREMOVE();
            _Analis.WLADD();


        }

        private void button1_Click(object sender, EventArgs e)
        {
            WL();
           
            Close();
            
        }
    }
}
