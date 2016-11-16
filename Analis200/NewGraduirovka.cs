using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
namespace Analis200
{
    public partial class NewGraduirovka : Form
    {
        Analis _Analis;
        public int oldValue = 3;
        public NewGraduirovka(Analis parent)
        {
            InitializeComponent();
            this._Analis = parent;
            Ed.SelectedIndex = 9;
            radioButton1.Checked = true;
            radioButton4.Checked = true;
            radioButton6.Checked = true;
            this.radioButton1.CheckedChanged += new EventHandler(radioButton1_CheckedChanged);
            this.radioButton2.CheckedChanged += new EventHandler(radioButton2_CheckedChanged);
            this.radioButton3.CheckedChanged += new EventHandler(radioButton3_CheckedChanged);
            this.radioButton4.CheckedChanged += new EventHandler(radioButton4_CheckedChanged);
            this.radioButton5.CheckedChanged += new EventHandler(radioButton5_CheckedChanged);
            this.radioButton6.CheckedChanged += new EventHandler(radioButton6_CheckedChanged);
            this.radioButton7.CheckedChanged += new EventHandler(radioButton7_CheckedChanged);
            if (_Analis.ComPodkl == true)
            {
                WL_grad.Text = _Analis.DLWL;
            }
        }

        private void NewGraduirovka_Load(object sender, EventArgs e)
        {
           
            _Analis.SposobZadan = "По СО";
            var height = 22;
            var labelx = 6;
            for (int i = 0; i <= 9; i++)
            {
                var label = new Label();
                label.Name = "CO" + i++.ToString();
                label.Text = "CO " + i-- + " =";
                label.Width = 40;
                label.Location = new Point(labelx, height);
                height += label.Height;
                groupBox6.Controls.Add(label);
            }
            var height1 = 19;
            var textBoxx = 52;
            
           
            for (int i = 0; i <= 9; i++)
            {
                _Analis.textBoxCO[i] = new TextBox();
                _Analis.textBoxCO[i].Name = "COtext" + i++.ToString();
                i--;
                _Analis.textBoxCO[i].Text = Convert.ToString("0,00");
                _Analis.textBoxCO[i].Width = 100;
                _Analis.textBoxCO[i].Height = 20;
                _Analis.textBoxCO[i].Location = new Point(textBoxx, height1);
                height1 += _Analis.textBoxCO[i].Height + 3;
                _Analis.textBoxCO[i].Enabled = false;
                groupBox6.Controls.Add(_Analis.textBoxCO[i]);
            }
            var height2 = 22;
            var labelx1 = 198;
            for (int i = 10; i <= 19; i++)
            {
                var label = new Label();
                label.Name = "CO" + i++.ToString();
                label.Text = "CO " + i-- + " =";
                label.Width = 40;
                label.Location = new Point(labelx1, height2);
                height2 += label.Height;
                this.Controls.Add(label);
                groupBox6.Controls.Add(label);
            }
            var height3 = 19;
            var textBoxx3 = 244;
            for (int i = 10; i <= 19; i++)
            {
                _Analis.textBoxCO[i] = new TextBox();
                _Analis.textBoxCO[i].Name = "COtext" + i++.ToString();
                i--;
                _Analis.textBoxCO[i].Text = Convert.ToString("0,00");
                _Analis.textBoxCO[i].Width = 100;
                _Analis.textBoxCO[i].Height = 20;
                _Analis.textBoxCO[i].Location = new Point(textBoxx3, height3);
                height3 += _Analis.textBoxCO[i].Height + 3;
                _Analis.textBoxCO[i].Enabled = false;
                groupBox6.Controls.Add(_Analis.textBoxCO[i]);
            }
            
            numericUpDown4.Value = 3;
            for (int i = Convert.ToInt32(numericUpDown4.Value) - 1; i >= 0; i--)
            {
                this._Analis.textBoxCO[i].Enabled = true;

            }
            k0Text.Text = string.Format("{0:0.0000}", 0);
            k1Text.Text = string.Format("{0:0.0000}", 0);
            k2Text.Text = string.Format("{0:0.0000}", 0);


           // MessageBox.Show(_Analis.WidthCuvette2);

            Veshestvo.Text = _Analis.Veshestvo1;
            WL_grad.Text = _Analis.wavelength1;
           int index = Opt_dlin_cuvet.FindString(_Analis.WidthCuvette);          
           
         //  MessageBox.Show(index.ToString());
            Opt_dlin_cuvet.SelectedIndex = index;
            Down.Text = _Analis.BottomLine;
            Up.Text = _Analis.TopLine;
            ND.Text = _Analis.ND;
            Description.Text = _Analis.Description;
            dateTimePicker1.Text = _Analis.DateTime;
            Ispolnitel.Text = _Analis.Ispolnitel;
            numericUpDown3.Value = Convert.ToInt32(_Analis.CountSeriya2);
            numericUpDown4.Value = Convert.ToInt32(_Analis.CountInSeriya2);
         //   numericUpDown1.Value = Convert.ToInt32(_Analis.DateTime1_1_1);
          //  textBox4.Text = _Analis.Pogreshnost2;
            for (int j = 0; j < numericUpDown4.Value; j++)
            {

                if (_Analis.Stolbec != null)
                {
                    _Analis.textBoxCO[j].Text = _Analis.Stolbec[j, 1];
                }
                
            }
            }

        public void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            //DateTime time1 = new DateTime((dateTimePicker1.Text));
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

       

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            k0Text.Enabled = false;
            k1Text.Enabled = false;
            k2Text.Enabled = false;
            numericUpDown3.Enabled = true;
            numericUpDown4.Enabled = true;
            k0Text.Text = string.Format("{0:0.0000}", 0);
            k1Text.Text = string.Format("{0:0.0000}", 0);
            k2Text.Text = string.Format("{0:0.0000}", 0);
            _Analis.SposobZadan = "По СО";
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            _Analis.SposobZadan = "Ввод коэффициентов";
            numericUpDown3.Enabled = false;
            numericUpDown4.Enabled = false;
            for (int i1 = 0; i1 <= 19; i1++)
            {
                _Analis.textBoxCO[i1].Enabled = false;
            }

            if (radioButton1.Checked == true && radioButton2.Checked == false && radioButton3.Checked == false)
            {
                k0Text.Enabled = false;
                k1Text.Enabled = true;
                k2Text.Enabled = false;
                k0Text.Text = string.Format("{0:0.0000}", 0);
                k2Text.Text = string.Format("{0:0.0000}", 0);
                _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
            }
            else
            {
                if (radioButton2.Checked == true && radioButton1.Checked == false && radioButton3.Checked == false)
                {
                    k0Text.Enabled = true;
                    k1Text.Enabled = true;
                    k2Text.Enabled = false;

                    k2Text.Text = string.Format("{0:0.0000}", 0);
                    _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                    _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                    _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                }
                else
                {
                    k0Text.Enabled = true;
                    k1Text.Enabled = true;
                    k2Text.Enabled = true;

                    _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                    _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                    _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                }
            }

        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            _Analis.radioButton1.Checked = true;
            if (radioButton7.Checked == true)
            {
                if (radioButton1.Checked == true && radioButton2.Checked == false && radioButton3.Checked == false)
                {
                    k0Text.Enabled = false;
                    k1Text.Enabled = true;
                    k2Text.Enabled = false;
                    k0Text.Text = string.Format("{0:0.0000}", 0);
                    k2Text.Text = string.Format("{0:0.0000}", 0);
                    _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                    _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                    _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                }
                else
                {
                    if (radioButton2.Checked == true && radioButton1.Checked == false && radioButton3.Checked == false)
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = false;

                        k2Text.Text = string.Format("{0:0.0000}", 0);
                        _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                        _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                        _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                    }
                    else
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = true;

                        _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                        _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                        _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                    }
                }
            }

        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            _Analis.radioButton2.Checked = true;
            if (radioButton7.Checked == true)
            {
                if (radioButton1.Checked == true && radioButton2.Checked == false && radioButton3.Checked == false)
                {
                    k0Text.Enabled = false;
                    k1Text.Enabled = true;
                    k2Text.Enabled = false;
                    k0Text.Text = string.Format("{0:0.0000}", 0);
                    k2Text.Text = string.Format("{0:0.0000}", 0);
                    _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                    _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                    _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                }
                else
                {
                    if (radioButton2.Checked == true && radioButton1.Checked == false && radioButton3.Checked == false)
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = false;

                        k2Text.Text = string.Format("{0:0.0000}", 0);
                        _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                        _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                        _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                    }
                    else
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = true;

                        _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                        _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                        _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                    }
                }
            }
  
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            _Analis.radioButton3.Checked = true;
            if (radioButton7.Checked == true)
            {
                if (radioButton1.Checked == true && radioButton2.Checked == false && radioButton3.Checked == false)
                {
                    k0Text.Enabled = false;
                    k1Text.Enabled = true;
                    k2Text.Enabled = false;
                    k0Text.Text = string.Format("{0:0.0000}", 0);
                    k2Text.Text = string.Format("{0:0.0000}", 0);
                    _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                    _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                    _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                }
                else
                {
                    if (radioButton2.Checked == true && radioButton1.Checked == false && radioButton3.Checked == false)
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = false;

                        k2Text.Text = string.Format("{0:0.0000}", 0);
                        _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                        _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                        _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                    }
                    else
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = true;

                        _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                        _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                        _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                    }
                }
            }

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            WL();
            _Analis.textBox3.Text = textBox4.Text;
            if (radioButton7.Checked == true)
            {
                _Analis.groupBox2.Visible = false;
                if (radioButton1.Checked == true && radioButton2.Checked == false && radioButton3.Checked == false)
                {
                    k0Text.Enabled = false;
                    k1Text.Enabled = true;
                    k2Text.Enabled = false;
                    k0Text.Text = string.Format("{0:0.0000}", 0);
                    k2Text.Text = string.Format("{0:0.0000}", 0);
                    _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                    _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                    _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                }
                else
                {
                    if (radioButton2.Checked == true && radioButton1.Checked == false && radioButton3.Checked == false)
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = false;

                        k2Text.Text = string.Format("{0:0.0000}", 0);
                        _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                        _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                        _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                    }
                    else
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = true;

                        _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                        _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                        _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                    }
                }
            }
            else
            {
                if (radioButton4.Checked == true)
                {
                    _Analis.radioButton4.Checked = true;
                }
                else
                {
                    _Analis.radioButton5.Checked = true;
                }
            }
            _Analis.edconctr = Ed.Text;
            if (_Analis.ComPodkl == true)
            {
                SW();
                _Analis.SAGE(ref _Analis.countSA, ref _Analis.GE5_1_0);
            }
            Close();
        }
        public void SW()
        {
            string SWText1 = WL_grad.Text;
            _Analis.newPort.Write("SW " + WL_grad.Text + "\r");
            Thread.Sleep(20000);
            int byteRecieved1 = _Analis.newPort.ReadBufferSize;
            Thread.Sleep(1500);
            byte[] buffer1 = new byte[byteRecieved1];
            _Analis.newPort.Read(buffer1, 0, byteRecieved1);
            _Analis.GWNew.Text = WL_grad.Text;

            // _Analis.GW();
        }
        public void WL()
        {
            _Analis.WLREMOVE1();
            _Analis.WLREMOVESTR1();
            _Analis.NoCaIzm = Convert.ToInt32(numericUpDown3.Value);
            _Analis.NoCaSer = Convert.ToInt32(numericUpDown4.Value);
            _Analis.WL_grad1 = WL_grad.Text;
            _Analis.WLADD1();
            _Analis.WLADDSTR1();


        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            //_Analis.radioButton4.Checked = true;
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
           // _Analis.radioButton5.Checked = true;
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            _Analis.NoCoIzmer = Convert.ToInt32(numericUpDown4.Value);


            if (this.oldValue > numericUpDown4.Value)
            {
                                
                        for (int i1 = 0; i1 <=19; i1++)
                        {
                        _Analis.textBoxCO[i1].Enabled = false;
                        }

                for (int i = _Analis.NoCoIzmer-1; i >= 0; i--)
                {
                    _Analis.textBoxCO[i].Enabled = true;

                }
            }
            else
            {
                for (int i = _Analis.NoCoIzmer-1; i >= 1; i--)
                {
                    _Analis.textBoxCO[i].Enabled = true;

                }
            }
            oldValue = _Analis.NoCoIzmer;
        }

        private void WL_grad_TextChanged(object sender, EventArgs e)
        {

        }

        private void Ed_SelectedIndexChanged(object sender, EventArgs e)
        {
            _Analis.edconctr = Ed.Text;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

          
            _Analis.Days = Convert.ToInt32(numericUpDown1.Value);
            _Analis.numericUpDown1.Value = numericUpDown1.Value;
        }

        private void Opt_dlin_cuvet_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }
        }
    }
}
