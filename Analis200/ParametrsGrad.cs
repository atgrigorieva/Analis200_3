﻿using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using SWF = System.Windows.Forms;
namespace Analis200
{
    public partial class ParametrsGrad : Form
    {
        Analis _Analis;
        public int oldValue = 3;
        public ParametrsGrad(Analis parent)
        {
            InitializeComponent();
            this._Analis = parent;
            Ed.SelectedIndex = 9;

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

            _Analis.chart1.Series[0].Points.Clear();
            _Analis.chart1.Series[1].Points.Clear();
        }

        private void ParametrsGrad_Load(object sender, EventArgs e)
        {
            _Analis.SposobZadan = "По СО";
            var height = 22;
            var labelx = 6;
            numericUpDown4.Value = oldValue;
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
                _Analis.textBoxCO[i].KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
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
                _Analis.textBoxCO[i].KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
            }

            numericUpDown4.Value = 3;
            for (int i = Convert.ToInt32(numericUpDown4.Value) - 1; i >= 0; i--)
            {
                this._Analis.textBoxCO[i].Enabled = true;

            }
            radioButton1.Checked = true;
            radioButton4.Checked = true;
            radioButton6.Checked = true;
            k0Text.Text = string.Format("{0:0.0000}", 0);
            k1Text.Text = string.Format("{0:0.0000}", 0);
            k2Text.Text = string.Format("{0:0.0000}", 0);

            _Analis.Veshestvo1 = Veshestvo.Text;
          //  _Analis.wavelength1 = WL_grad.Text;
            _Analis.WidthCuvette = Opt_dlin_cuvet.Text;
            _Analis.textBox3.Text = Opt_dlin_cuvet.Text;
            _Analis.textBox1.Text = Description.Text;
            _Analis.BottomLine = Down.Text;
            _Analis.TopLine = Up.Text;
            _Analis.ND = ND.Text;
            _Analis.Description = Description.Text;
            _Analis.DateTime = dateTimePicker1.Text;
            _Analis.Ispolnitel = Ispolnitel.Text;
            _Analis.CountSeriya2 = Convert.ToString(numericUpDown3.Value);
            _Analis.CountInSeriya2 = Convert.ToString(numericUpDown4.Value);
            //       _Analis.NoCaIzm = Convert.ToInt32(numericUpDown4.Value);
            //      _Analis.NoCaSer = Convert.ToInt32(numericUpDown3.Value);

            if (k0Text.Text == "не число" || k1Text.Text == "не число" || k2Text.Text == "не число")
            {
                k0Text.Text = string.Format("{0:0.0000}", 0);
                k2Text.Text = string.Format("{0:0.0000}", 0);
                k1Text.Text = string.Format("{0:0.0000}", 0);
            }
            _Analis.textBox4.Text = k0Text.Text;
            _Analis.textBox5.Text =  k1Text.Text;
            _Analis.textBox6.Text = k2Text.Text;
            //  _Analis.GWNew.Text
            if (_Analis.ComPodkl == true)
            {
                WL_grad.Text = _Analis.DLWL;
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
            _Analis.Zavisimoct = "C(A)";
        }

        public void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            k0Text.Enabled = false;
            k1Text.Enabled = false;
            k2Text.Enabled = false;
            _Analis.Table1.Visible = true;
            for (int i1 = 0; i1 < numericUpDown4.Value; i1++)
            {
                _Analis.textBoxCO[i1].Enabled = true;
            }
            numericUpDown3.Enabled = true;
            numericUpDown4.Enabled = true;
            k0Text.Text = string.Format("{0:0.0000}", 0);
            k1Text.Text = string.Format("{0:0.0000}", 0);
            k2Text.Text = string.Format("{0:0.0000}", 0);
      
            _Analis.SposobZadan = "По СО";
        }

        public void radioButton7_CheckedChanged(object sender, EventArgs e)
        {

            _Analis.SposobZadan = "Ввод коэффициентов";
            numericUpDown3.Enabled = false;
            numericUpDown4.Enabled = false;
            _Analis.Table1.Visible = false;
            for (int i1 = 0; i1 <= numericUpDown4.Value; i1++)
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
                double k0 = Convert.ToDouble(k0Text.Text);
                double k1 = Convert.ToDouble(k1Text.Text);
                double k2 = Convert.ToDouble(k2Text.Text);
                _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C";
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
                    double k0 = Convert.ToDouble(k0Text.Text);
                    double k1 = Convert.ToDouble(k1Text.Text);
                    double k2 = Convert.ToDouble(k2Text.Text);
                    _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C" + k0.ToString(" 0.0000 ;- 0.0000 ");
                }
                else
                {
                    k0Text.Enabled = true;
                    k1Text.Enabled = true;
                    k2Text.Enabled = true;

                    _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                    _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                    _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                    double k0 = Convert.ToDouble(k0Text.Text);
                    double k1 = Convert.ToDouble(k1Text.Text);
                    double k2 = Convert.ToDouble(k2Text.Text);
                    _Analis.label14.Text = "A(C) = " + k0.ToString(" 0.0000 ;- 0.0000 ") + k1.ToString("0.0000 ;- 0.0000 ") + "*C" + k2.ToString("0.0000 ;- 0.0000 ") + "*C^2";
                }
            }

        }
        public void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            _Analis.aproksim = "Линейная через 0";
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
                    double k0 = Convert.ToDouble(k0Text.Text);
                    double k1 = Convert.ToDouble(k1Text.Text);
                    double k2 = Convert.ToDouble(k2Text.Text);
                    _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C";
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
                        double k0 = Convert.ToDouble(k0Text.Text);
                        double k1 = Convert.ToDouble(k1Text.Text);
                        double k2 = Convert.ToDouble(k2Text.Text);
                        _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C" + k0.ToString(" 0.0000 ;- 0.0000 ");
                    }
                    else
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = true;

                        _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                        _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                        _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                        double k0 = Convert.ToDouble(k0Text.Text);
                        double k1 = Convert.ToDouble(k1Text.Text);
                        double k2 = Convert.ToDouble(k2Text.Text);
                        _Analis.label14.Text = "A(C) = " + k0.ToString(" 0.0000 ;- 0.0000 ") + k1.ToString("0.0000 ;- 0.0000 ") + "*C" + k2.ToString("0.0000 ;- 0.0000 ") + "*C^2";
                    }
                }
            }

        }
        public void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            _Analis.aproksim = "Линейная";
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
                    double k0 = Convert.ToDouble(k0Text.Text);
                    double k1 = Convert.ToDouble(k1Text.Text);
                    double k2 = Convert.ToDouble(k2Text.Text);
                    _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C";
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
                        double k0 = Convert.ToDouble(k0Text.Text);
                        double k1 = Convert.ToDouble(k1Text.Text);
                        double k2 = Convert.ToDouble(k2Text.Text);
                        _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C" + k0.ToString(" 0.0000 ;- 0.0000 ");
                    }
                    else
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = true;

                        _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                        _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                        _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                        double k0 = Convert.ToDouble(k0Text.Text);
                        double k1 = Convert.ToDouble(k1Text.Text);
                        double k2 = Convert.ToDouble(k2Text.Text);
                        _Analis.label14.Text = "A(C) = " + k0.ToString(" 0.0000 ;- 0.0000 ") + k1.ToString("0.0000 ;- 0.0000 ") + "*C" + k2.ToString("0.0000 ;- 0.0000 ") + "*C^2";
                    }
                }
            }
        }

        public void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            _Analis.aproksim = "Квадратичная";
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
                    double k0 = Convert.ToDouble(k0Text.Text);
                    double k1 = Convert.ToDouble(k1Text.Text);
                    double k2 = Convert.ToDouble(k2Text.Text);
                    _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C";
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
                        double k0 = Convert.ToDouble(k0Text.Text);
                        double k1 = Convert.ToDouble(k1Text.Text);
                        double k2 = Convert.ToDouble(k2Text.Text);
                        _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C" + k0.ToString(" 0.0000 ;- 0.0000 ");
                    }
                    else
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = true;

                        _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                        _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                        _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                        double k0 = Convert.ToDouble(k0Text.Text);
                        double k1 = Convert.ToDouble(k1Text.Text);
                        double k2 = Convert.ToDouble(k2Text.Text);
                        _Analis.label14.Text = "A(C) = " + k0.ToString(" 0.0000 ;- 0.0000 ") + k1.ToString("0.0000 ;- 0.0000 ") + "*C" + k2.ToString("0.0000 ;- 0.0000 ") + "*C^2";
                    }
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            WL();
            _Analis.textBox1.Text = Description.Text;
            _Analis.textBox2.Text = Opt_dlin_cuvet.Text;
            _Analis.textBox3.Text = textBox4.Text;
            _Analis.tabPage4.Parent = null;
            if (radioButton4.Checked == true)
            {
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
                        double k0 = Convert.ToDouble(k0Text.Text);
                        double k1 = Convert.ToDouble(k1Text.Text);
                        double k2 = Convert.ToDouble(k2Text.Text);
                        _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C";

                        for (double i = 0; i <= 3; i++)
                        {
                            double x2 = i;
                            double y2 = i*k1;
                            _Analis.chart1.Series[1].Points.AddXY(x2, y2);
                            //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                            _Analis.chart1.Series[0].Enabled = false;
                            _Analis.chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + Ed.Text;
                            _Analis.chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                            _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                            //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                            _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                        }
                       // _Analis.chart1
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
                            double k0 = Convert.ToDouble(k0Text.Text);
                            double k1 = Convert.ToDouble(k1Text.Text);
                            double k2 = Convert.ToDouble(k2Text.Text);
                            _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C" + k0.ToString(" + 0.0000 ;- 0.0000 ");
                            for (double i = 0; i <= 3; i++)
                            {
                                double x2 = i;
                                double y2 = i * k1 + k0;
                                _Analis.chart1.Series[1].Points.AddXY(x2, y2);
                                _Analis.chart1.Series[0].Enabled = false;
                                //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                                _Analis.chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + Ed.Text;
                                _Analis.chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                                _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                                //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                                _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                            }
                        }
                        else
                        {
                            k0Text.Enabled = true;
                            k1Text.Enabled = true;
                            k2Text.Enabled = true;

                            _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                            _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                            _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                            double k0 = Convert.ToDouble(k0Text.Text);
                            double k1 = Convert.ToDouble(k1Text.Text);
                            double k2 = Convert.ToDouble(k2Text.Text);
                            _Analis.label14.Text = "A(C) = " + k0.ToString(" 0.0000 ;- 0.0000 ") + k1.ToString(" + 0.0000 ;- 0.0000 ") + "*C" + k2.ToString(" + 0.0000 ;- 0.0000 ") + "*C^2";
                            for (double i = 0; i <= 3; i++)
                            {
                                double x2 = i;
                                double y2 = i * k1 + k0 + i*k2*k2;
                                _Analis.chart1.Series[0].Enabled = false;
                                _Analis.chart1.Series[1].Points.AddXY(x2, y2);
                                //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                                _Analis.chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + Ed.Text;
                                _Analis.chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                                _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                                //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                                _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                            }
                        }
                    }
                    if (radioButton6.Checked == true)
                    {
                        _Analis.SposobZadan = "По СО";
                    }
                    else
                    {
                        _Analis.SposobZadan = "Ввод коэффициентов";
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
                    if (radioButton4.Checked == true)
                    {
                        _Analis.Zavisimoct = "A(C)";
                    }
                    else
                    {
                        _Analis.Zavisimoct = "C(A)";
                    }
                    if (radioButton1.Checked == true)
                    {
                        _Analis.aproksim = "Линейная через 0";
                    }
                    else
                    {
                        if (radioButton2.Checked == true)
                        {
                            _Analis.aproksim = "Линейная";
                        }
                        else
                        {
                            _Analis.aproksim = "Квадратичная";
                        }
                    }
                }
            }
            else
            {
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
                        double k0 = Convert.ToDouble(k0Text.Text);
                        double k1 = Convert.ToDouble(k1Text.Text);
                        double k2 = Convert.ToDouble(k2Text.Text);
                        _Analis.label14.Text = "C(A) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*A";
                        for (double i = 0; i <= 3; i++)
                        {
                            double x2 = i;
                            double y2 = i * k1;
                            _Analis.chart1.Series[1].Points.AddXY(y2, x2);
                            //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                            _Analis.chart1.Series[0].Enabled = false;
                            _Analis.chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + Ed.Text;
                            _Analis.chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                            _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                            //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                            _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                        }
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
                            double k0 = Convert.ToDouble(k0Text.Text);
                            double k1 = Convert.ToDouble(k1Text.Text);
                            double k2 = Convert.ToDouble(k2Text.Text);
                            _Analis.label14.Text = "C(A) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*A" + k0.ToString(" + 0.0000 ;- 0.0000 ");
                            for (double i = 0; i <= 3; i++)
                            {
                                double x2 = i;
                                double y2 = i * k1 + k0;
                                _Analis.chart1.Series[1].Points.AddXY(y2, x2);
                                _Analis.chart1.Series[0].Enabled = false;
                                //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                                _Analis.chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + Ed.Text;
                                _Analis.chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                                _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                                //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                                _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                            }
                        }
                        else
                        {
                            k0Text.Enabled = true;
                            k1Text.Enabled = true;
                            k2Text.Enabled = true;

                            _Analis.textBox4.Text = string.Format("{0:0.0000}", k0Text.Text);
                            _Analis.textBox5.Text = string.Format("{0:0.0000}", k1Text.Text);
                            _Analis.textBox6.Text = string.Format("{0:0.0000}", k2Text.Text);
                            double k0 = Convert.ToDouble(k0Text.Text);
                            double k1 = Convert.ToDouble(k1Text.Text);
                            double k2 = Convert.ToDouble(k2Text.Text);
                            _Analis.label14.Text = "C(A) = " + k0.ToString(" 0.0000 ;- 0.0000 ") + k1.ToString(" + 0.0000 ;- 0.0000 ") + "*A" + k2.ToString(" + 0.0000 ;- 0.0000 ") + "*A^2";
                            for (double i = 0; i <= 3; i++)
                            {
                                double x2 = i;
                                double y2 = i * k1 + k0 + i * k2 * k2;
                                _Analis.chart1.Series[0].Enabled = false;
                                _Analis.chart1.Series[1].Points.AddXY(y2, x2);
                                //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                                _Analis.chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + Ed.Text;
                                _Analis.chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                                _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                                //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                                _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                            }
                        }
                    }
                    if (radioButton6.Checked == true)
                    {
                        _Analis.SposobZadan = "По СО";
                    }
                    else
                    {
                        _Analis.SposobZadan = "Ввод коэффициентов";
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
                    if (radioButton4.Checked == true)
                    {
                        _Analis.Zavisimoct = "A(C)";
                    }
                    else
                    {
                        _Analis.Zavisimoct = "C(A)";
                    }
                    if (radioButton1.Checked == true)
                    {
                        _Analis.aproksim = "Линейная через 0";
                    }
                    else
                    {
                        if (radioButton2.Checked == true)
                        {
                            _Analis.aproksim = "Линейная";
                        }
                        else
                        {
                            _Analis.aproksim = "Квадратичная";
                        }
                    }
                }
            }
            _Analis.edconctr = Ed.Text;
            if (_Analis.ComPodkl == true)
            {
                SW();
                _Analis.SAGE(ref _Analis.countSA, ref _Analis.GE5_1_0);
            }
            _Analis.параметрыToolStripMenuItem.Enabled = true;
            _Analis.button10.Enabled = true;
            Close();
        }
        public void SW()
        {
            LogoForm();
            string SWText1 = WL_grad.Text;
            _Analis.newPort.Write("SW " + WL_grad.Text + "\r");
            Thread.Sleep(20000);
            int byteRecieved1 = _Analis.newPort.ReadBufferSize;
            Thread.Sleep(1500);
            byte[] buffer1 = new byte[byteRecieved1];
            _Analis.newPort.Read(buffer1, 0, byteRecieved1);
            _Analis.GWNew.Text = WL_grad.Text;
            SWF.Application.OpenForms["LogoFrm"].Close();
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
            _Analis.Zavisimoct = "A(C)";
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            // _Analis.radioButton5.Checked = true;
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            _Analis.NoCoIzmer = Convert.ToInt32(numericUpDown4.Value);
            if (numericUpDown4.Value == 1)
            {
                radioButton2.Enabled = false;
                radioButton3.Enabled = false;
            }
            if (numericUpDown4.Value == 2)
            {
                radioButton2.Enabled = true;
                radioButton3.Enabled = false;
            }
            if (numericUpDown4.Value >= 3)
            {
                radioButton2.Enabled = true;
                radioButton3.Enabled = true;
            }



            if (this.oldValue > numericUpDown4.Value)
            {

                for (int i1 = 0; i1 <= 19; i1++)
                {
                    _Analis.textBoxCO[i1].Enabled = false;
                }

                for (int i = _Analis.NoCoIzmer - 1; i >= 0; i--)
                {
                    _Analis.textBoxCO[i].Enabled = true;

                }
            }
            else
            {
                for (int i = _Analis.NoCoIzmer - 1; i >= 1; i--)
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
                MessageBox.Show("Только цифры!");
            }
        }
        void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("Только цифры!");
            }
        }
        public void LogoForm()
        {
            Form LogoForm = new Form();
            // LogoForm.BackColor = System.Drawing.Color.White;
            LogoForm.BackgroundImage = System.Drawing.Image.FromFile("Yasnovka_DLWALVE.png");
            LogoForm.AutoScaleMode = AutoScaleMode.Font;
            LogoForm.Size = new Size(430, 107);
            LogoForm.Text = "Установка длины волны...";
            LogoForm.MinimizeBox = false;
            LogoForm.MaximizeBox = false;
            LogoForm.AutoSize = false;
            LogoForm.Name = "LogoFrm";
            LogoForm.ShowInTaskbar = false;
            LogoForm.StartPosition = FormStartPosition.CenterScreen;
            LogoForm.ControlBox = false;
            LogoForm.FormBorderStyle = FormBorderStyle.None;
            /*PictureBox PicBox = new PictureBox();
            PicBox.Size = new Size(307, 179);
            PicBox.Location = new System.Drawing.Point(12, 12);
            PicBox.ImageLocation = "D:\\Analis-samo\\Analis200\\Analis200\bin\x64\\Release\\Calibrovka.png";
            PicBox.SizeMode = PictureBoxSizeMode.Zoom;
            LogoForm.Controls.Add(PicBox);*/
            LogoForm.Show();
        }

        private void Down_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("Только цифры!");
            }
        }

        private void Up_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("Только цифры!");
            }
        }

        private void k0Text_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("Только цифры!");
            }
        }

        private void k1Text_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("Только цифры!");
            }
        }

        private void k2Text_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("Только цифры!");
            }
        }
    }
}
