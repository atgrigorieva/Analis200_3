using System;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Xml;
using SWF = System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing.Printing;
using System.Text.RegularExpressions;

namespace Analis200
{

    public partial class Analis : Form
    {
        public SerialPort newPort;
        public SWF.TextBox[] textBox = new SWF.TextBox[20];
        public SWF.TextBox[] textBoxCO = new SWF.TextBox[20];
        public string WL_grad1;
        public int NoCaIzm;
        public int NoCaIzm1;
        public int NoCaSer;
        public int NoCaSer1;
        public int countWL = 3;
        public byte[] buffer4;
        public string GW1_2;
        public string GE5_1_0 = "";
        public byte[] GWbuffer;
        public int NoCoIzmer;
        public int NoCoSeria;
        public string edconctr;
        public int StolbecCol = 0;
        public string SposobZadan;
        public int Days;
        public bool ComPodkl = false;
        public string DLWL;
        public string versionPribor;
        public Analis()
        {

            InitializeComponent();
            this.StartPosition = FormStartPosition.WindowsDefaultBounds;
            //Microsoft.Win32.SystemEvents.DisplaySettingsChanged += new EventHandler(SystemEvents_DisplaySettingsChanged);
            //this.FormBorderStyle = FormBorderStyle.None;
            // this.WindowState = FormWindowState.Maximized;
            if (ComPodkl == false)
            {
                this.подключитьToolStripMenuItem.Enabled = true;
                button2.Enabled = true;
                button13.Enabled = false;
                button12.Enabled = false;
                button14.Enabled = false;
                this.настройкаПортаToolStripMenuItem.Enabled = false;
                this.информацияToolStripMenuItem.Enabled = false;
                this.калибровкаToolStripMenuItem.Enabled = false;
                this.темновойТокToolStripMenuItem.Enabled = false;
                this.измеритьToolStripMenuItem.Enabled = false;
                this.одноволновоеИзмерениеToolStripMenuItem.Enabled = false;
                this.одноволновоеИзмерениеToolStripMenuItem.Enabled = false;
                this.калибровкаДляОдноволновогоАнализаToolStripMenuItem.Enabled = false;
            }
            //   groupBox3.Enabled = false;
            // groupBox2.Enabled = false;
            tabPage4.Parent = null;
            radioButton1.Enabled = false;
            radioButton2.Enabled = false;
            radioButton3.Enabled = false;
            radioButton4.Enabled = false;
            radioButton5.Enabled = false;

        }


        public void Analis_Load(object sender, EventArgs e)
        {
            edconctr = "%";
            SposobZadan = "По СО";

            dateTimePicker1.Text = DateTime;

            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;

            chart1.Series[0].Points.Add();

            chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            Zavisimoct = "A(C)";
            aproksim = "Линейная";
            Table1.AllowUserToResizeRows = false;


        }
        public int countSA;
        bool Izmerenie1 = true;
        public void LogoForm()
        {
            Form LogoForm = new Form();
            // LogoForm.BackColor = System.Drawing.Color.White;
            LogoForm.BackgroundImage = System.Drawing.Image.FromFile("Calibrovka.png");
            LogoForm.AutoScaleMode = AutoScaleMode.Font;
            LogoForm.Size = new Size(430, 107);
            LogoForm.Text = "Калибровка...";
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
        public void SAGE(ref int countSA, ref string GE5_1_0)
        {
            bool message1 = true;
            if (versionPribor.Contains("2"))
            { countSA = 8; }
            else
            {
                countSA = 4;
            }
            // while (message1 == true)
            //{
            //picturebox1.Show("Калибровка");
            LogoForm();
            // System.IO.FileStream fs = new System.IO.FileStream(@"Calibrovka.png", System.IO.FileMode.Open);
            //  System.Drawing.Image img = System.Drawing.Image.FromStream(fs);
            //  fs.Close();
            //  pictureBox.Image = img;
            newPort.Write("SA " + countSA + "\r");
            // int SAAnalisByteRecieved1 = newPort.ReadBufferSize;
            // Thread.Sleep(100);
            //byte[] SAAnalisBuffer1 = new byte[SAAnalisByteRecieved1];
            // newPort.Read(SAAnalisBuffer1, 0, SAAnalisByteRecieved1);
            string indata = newPort.ReadExisting();
            int indata_zero = 0;
            string indata_0;
            bool indata_bool = true;
            while (indata_bool == true)
            {

                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = newPort.ReadExisting();

                }
            }

            newPort.Write("GE 1\r");

            string GE5_1 = "";
            int GEbyteRecieved4_1 = newPort.ReadBufferSize;
            byte[] GEbuffer4_1 = new byte[GEbyteRecieved4_1];
            newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

            indata = newPort.ReadExisting();
            indata_zero = 0;
            indata_0 = "";
            indata_bool = true;
            while (indata_bool == true)
            {

                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {

                    indata = newPort.ReadExisting();
                    indata_0 += indata;
                }
            }
            Regex regex = new Regex(@"\W");
            GE5_1 = regex.Replace(indata_0, "");
            GE5_1_0 = regex.Replace(indata_0, "");
            GEText.Text = GE5_1_0;
            double GAText1 = (Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1_0)) * 100;

            GAText.Text = string.Format("{0:0.00}", GAText1);

            double OptPlot = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1));

            double OptPlot1 = OptPlot - Math.Truncate(OptPlot);
            OptichPlot.Text = string.Format("{0:0.0000}", OptPlot1);
            while (Convert.ToInt32(GE5_1) > 30000 && countSA > 1)
            {
                countSA--;
                newPort.Write("SA " + countSA + "\r");
                int SAAnalisByteRecieved1_1_1 = newPort.ReadBufferSize;
                // Thread.Sleep(100);
                indata = newPort.ReadExisting();
                indata_zero = 0;
                indata_0 = "";
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();
                    }
                }
                //   byte[] SAAnalisBuffer1_1_1 = new byte[SAAnalisByteRecieved1_1_1];
                //  newPort.Read(SAAnalisBuffer1_1_1, 0, SAAnalisByteRecieved1_1_1);

                newPort.Write("GE 1\r");
                GEbyteRecieved4_1 = newPort.ReadBufferSize;
                GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

                indata = newPort.ReadExisting();
                indata_zero = 0;
                indata_0 = "";
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {

                        indata = newPort.ReadExisting();
                        indata_0 += indata;
                    }
                }
                regex = new Regex(@"\W");
                GE5_1 = regex.Replace(indata_0, "");

                GE5_1_0 = regex.Replace(indata_0, "");
                GEText.Text = GE5_1_0;

                GAText1 = (Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1_0)) * 100;

                GAText.Text = string.Format("{0:0.00}", GAText1);
                //double GAText2 = (Convert.ToDouble(GE3) / Convert.ToDouble(GE5)) * 0.01;
                OptPlot = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1));
                // double OptPlot1 = (OptPlot - Math.Truncate(OptPlot))*500;
                // OptPlot1 = (OptPlot1 - Math.Truncate(OptPlot1))*10;
                //OptPlot1 = Math.Truncate(OptPlot1)/5000;
                OptPlot1 = OptPlot - Math.Truncate(OptPlot);
                OptichPlot.Text = string.Format("{0:0.0000}", OptPlot1);
            }
            SWF.Application.OpenForms["LogoFrm"].Close();
            //  message1 = false;
            // }
            if (Izmerenie1 == false)
            {
                InitializeTimer();
            }
        }

        private void справкаToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            // Process.Start();
        }
        public string Veshestvo1;
        public string wavelength1;
        public string WidthCuvette;
        public string BottomLine;
        public string TopLine;
        public string ND;
        public string Description;
        public string DateTime;
        public string Ispolnitel;
        public string CountSeriya;
        public string CountInSeriya;
        private void параметрыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                //    this.textBox4.Text = string.Format("{0:0.0000}", 0);
                //   this.textBox5.Text = string.Format("{0:0.0000}", 0);
                //   this.textBox6.Text = string.Format("{0:0.0000}", 0);
                //    Izmerenie1 = true;
                while (true)
                {
                    int i = Table1.Columns.Count - 1;//С какого столбца начать
                    if (Table1.Columns.Count == 3 + NoCaIzm)
                        break;
                    Table1.Columns.RemoveAt(i);
                }


                NewGraduirovka _NewGraduirovka = new NewGraduirovka(this);

                _NewGraduirovka.button1.Click += (NewGraduirovka, eSlave) =>
                {
                    Veshestvo1 = _NewGraduirovka.Veshestvo.Text;
                    wavelength1 = _NewGraduirovka.WL_grad.Text;
                    WidthCuvette = _NewGraduirovka.Opt_dlin_cuvet.Text;
                    BottomLine = _NewGraduirovka.Down.Text;
                    TopLine = _NewGraduirovka.Up.Text;
                    ND = _NewGraduirovka.ND.Text;
                    Description = _NewGraduirovka.Description.Text;
                    DateTime = _NewGraduirovka.dateTimePicker1.Text;
                    Ispolnitel = _NewGraduirovka.Ispolnitel.Text;
                    textBox3.Text = _NewGraduirovka.textBox4.Text;
                    CountSeriya = _NewGraduirovka.numericUpDown3.Text;
                    CountInSeriya = _NewGraduirovka.numericUpDown4.Text;
                    edconctr = _NewGraduirovka.Ed.Text;
                    Days = Convert.ToInt32(_NewGraduirovka.numericUpDown1.Value);
                    label6.Text = dateTimePicker1.Value.AddDays(Days).ToString("dd.MM.yyyy");
                    if (_NewGraduirovka.radioButton7.Checked == true)
                    {
                        this.textBox4.Text = string.Format("{0:0.0000}", _NewGraduirovka.k0Text.Text);
                        this.textBox5.Text = string.Format("{0:0.0000}", _NewGraduirovka.k1Text.Text);
                        this.textBox6.Text = string.Format("{0:0.0000}", _NewGraduirovka.k2Text.Text);


                    }
                    else
                    {
                        if (_NewGraduirovka.radioButton6.Checked == true)
                        {
                            //  this.textBox4.Text = string.Format("{0:0.0000}", 0);
                            //   this.textBox5.Text = string.Format("{0:0.0000}", 0);
                            ///   this.textBox6.Text = string.Format("{0:0.0000}", 0);
                        }
                    }

                };
                _NewGraduirovka.ShowDialog();
                dateTimePicker1.Text = DateTime;
                tabControl2.SelectTab(tabPage3);
                chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                //       Table1.Columns.Add("X*X", "Concetr*Concetr");
                //Table1.Columns.Add("X*Y", "Asred*Concetr");
                //CalibrovkaGrad();
                // MessageBox.Show(Convert.ToString(Days));

            }
            else
            {

                Parametr1 _Parametr1 = new Parametr1(this);
                _Parametr1.ShowDialog();
            }
        }

        public void WLREMOVE1()
        {
            while (true)
            {
                int i = Table1.Columns.Count - 1;//С какого столбца начать
                if (Table1.Columns[i].Name == "Asred")
                    break;
                Table1.Columns.RemoveAt(i);
            }

        }
        public void WLADD1()
        {

            for (int i = 1; i <= NoCaIzm; i++)
            {

                DataGridViewTextBoxColumn firstColumn1 =
                new DataGridViewTextBoxColumn();
                firstColumn1.HeaderText = "A; Сер" + i;
                firstColumn1.Name = "A;Ser (" + i;
                firstColumn1.ValueType = Type.GetType("System.Double");

                Table1.Columns.Add(firstColumn1);
                //  firstColumn1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
                //   firstColumn1.EditingControlShowing

            }

            for (int i = 1; i <= NoCaIzm; i++)
            {
                Table1.Columns["A;Ser (" + i].Width = 50;
            }
            Concetr.HeaderText = "Конц " + edconctr;
        }
        public void WLADDSTR1()
        {

            for (int i = 1; i <= NoCaSer; i++)
            {
                Table1.Rows.Add(i, textBoxCO[i - 1].Text);


            }
            // Table1.Rows.Add();
        }
        public void WLREMOVESTR1()
        {
            Table1.Rows.Clear();

        }

        public void WLREMOVE2()
        {
            while (true)
            {
                int i = Table2.Columns.Count - 1;//С какого столбца начать
                if (Table2.Columns[i].Name == "Obrazec")
                    break;
                Table2.Columns.RemoveAt(i);
            }

        }
        public void WLADD2()
        {
            if (NoCaIzm1 >= 2)
            {
                for (int i = 1; i <= NoCaIzm1; i++)
                {

                    DataGridViewTextBoxColumn firstColumn2 =
                    new DataGridViewTextBoxColumn();
                    firstColumn2.HeaderText = "A; Сер." + i;
                    firstColumn2.Name = "A;Ser" + i;
                    firstColumn2.ValueType = Type.GetType("System.Double");
                    Table2.Columns.Add(firstColumn2);
                    DataGridViewTextBoxColumn firstColumn3 =
                    new DataGridViewTextBoxColumn();
                    firstColumn3.HeaderText = "C, " + edconctr + "; Сер." + i;
                    firstColumn3.Name = "C,edconctr;Ser." + i;
                    firstColumn3.ValueType = Type.GetType("System.Double");
                    Table2.Columns.Add(firstColumn3);
                    // Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A; Сер" + i].KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
                    firstColumn3.ReadOnly = true;
                }
            }
            else
            {

                DataGridViewTextBoxColumn firstColumn2_1 =
                        new DataGridViewTextBoxColumn();
                firstColumn2_1.HeaderText = "A; Сер." + 1;
                firstColumn2_1.Name = "A;Ser" + 1;
                firstColumn2_1.ValueType = Type.GetType("System.Double");
                Table2.Columns.Add(firstColumn2_1);
                DataGridViewTextBoxColumn firstColumn3_1 =
                new DataGridViewTextBoxColumn();
                firstColumn3_1.HeaderText = "C, " + edconctr + "; Сер." + 1;
                firstColumn3_1.Name = "C,edconctr;Ser." + 1;
                firstColumn3_1.ValueType = Type.GetType("System.Double");
                Table2.Columns.Add(firstColumn3_1);
                firstColumn3_1.ReadOnly = true;
            }
            DataGridViewTextBoxColumn firstColumn4 =
            new DataGridViewTextBoxColumn();
            firstColumn4.HeaderText = "Cср, " + edconctr;
            firstColumn4.Name = "Ccr";
            firstColumn4.ValueType = Type.GetType("System.Double");
            Table2.Columns.Add(firstColumn4);
            firstColumn4.ReadOnly = true;
            DataGridViewTextBoxColumn firstColumn5 =
            new DataGridViewTextBoxColumn();
            firstColumn5.HeaderText = "d, %";
            firstColumn5.Name = "d%";
            firstColumn5.ValueType = Type.GetType("System.Double");
            firstColumn5.ReadOnly = true;
            Table2.Columns.Add(firstColumn5);

        }
        public void WLADDSTR2()
        {
            count = 0;
            if (NoCaSer1 > 1)
            {
                for (int i = 1; i <= NoCaSer1; i++)
                {
                    Table2.Rows.Add(i);
                    Table2.Rows[count].Cells["Column1"].Value = count + 1;
                    count++;
                }
            }
            else
            {
                Table2.Rows.Add(1);
                Table2.Rows[count].Cells["Column1"].Value = count + 1;
                count++;
            }
            for (int i = 0; i < Table2.RowCount - 1; i++)
            {

                if (Table2.Rows[i].Cells["Column1"].Value == null)
                {
                    Table2.Rows.RemoveAt(i);
                    i--;
                }
            }
        }
        public void WLREMOVESTR2()
        {
            Table2.Rows.Clear();

        }
        public string portsName = "";
        public bool ComPort = false;
        public void подключитьToolStripMenuItem_Click(object sender, EventArgs e)
        {

            SettingPort _SettingPort = new SettingPort(this);
            if (nonPort == true)
            {
                _SettingPort.ShowDialog();
            }
            else
            {
                _SettingPort.Dispose();
            }
            //_SettingPort.Close();
            if (nonPort == true && Izmerenie1 != false)
            {
                newPort = new SerialPort();

                try
                {
                    // настройки порта (Communication interface)
                    newPort.PortName = portsName;
                    newPort.BaudRate = 19200;
                    newPort.DataBits = 8;
                    newPort.Parity = System.IO.Ports.Parity.None;
                    newPort.StopBits = System.IO.Ports.StopBits.One;
                    // Установка таймаутов чтения/записи (read/write timeouts)
                    newPort.ReadTimeout = 20000;
                    newPort.WriteTimeout = 20000;
                    //    newPort.DataReceived += new SerialDataReceivedEventHandler(newPort_DataReceived);
                    newPort.RtsEnable = false;
                    newPort.DtrEnable = true;
                    newPort.Open();
                    // MessageBox.Show("ПОРТ ОТКРЫТ " + newPort.PortName);


                    newPort.DiscardInBuffer();
                    newPort.DiscardOutBuffer();
                }
                catch (Exception)
                {
                    MessageBox.Show("Порт не был выбран!");
                    return;

                }

                //char[] OpenPribor = { Convert.ToChar('C'), Convert.ToChar('O'), Convert.ToChar('\r') };
                //newPort.Write(OpenPribor, 0, OpenPribor.Length);


                Analis__Activated();
                CO();
                // GW();
                RD();
                //GW();
                //SA();
                Izmerenie1 = false;
                ComPodkl = true;
                SAGE(ref countSA, ref GE5_1_0);
                this.подключитьToolStripMenuItem.Enabled = false;
                this.настройкаПортаToolStripMenuItem.Enabled = true;
                this.информацияToolStripMenuItem.Enabled = true;
                this.калибровкаToolStripMenuItem.Enabled = true;
                this.темновойТокToolStripMenuItem.Enabled = true;
                this.измеритьToolStripMenuItem.Enabled = true;
                this.одноволновоеИзмерениеToolStripMenuItem.Enabled = true;
                this.одноволновоеИзмерениеToolStripMenuItem.Enabled = true;
                this.калибровкаДляОдноволновогоАнализаToolStripMenuItem.Enabled = true;
                button2.Enabled = false;
                button13.Enabled = true;
                button12.Enabled = true;
                button14.Enabled = true;
                ComPort = true;
                button1.Enabled = true;

                wavelength1 = GWNew.Text;
            }
        }

        string SW1 = "";
        public void SW()
        {
            newPort.Write("SW 550\r");


            string indata = newPort.ReadExisting();

            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = newPort.ReadExisting();
                }
            }



            GWNew.Text = "550";

        }
        public void SW2()
        {
            LogoForm2();
            string SWText1 = wavelength1;
            newPort.Write("SW " + wavelength1 + "\r");
            //  Thread.Sleep(20000);
            string indata = newPort.ReadExisting();

            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = newPort.ReadExisting();
                }
            }
            GWNew.Text = wavelength1;
            SWF.Application.OpenForms["LogoFrm2"].Close();
            // _Analis.GW();
        }
        public void InitializeTimer()
        {
            // Run this procedure in an appropriate event.
            if (Izmerenie1 == false)
            {
                timer1.Interval = 6000;
                timer1.Enabled = true;
                // Hook up timer's tick event handler.
                this.timer1.Tick += new System.EventHandler(this.timer1_Tick);

            }

            //break;

        }
        public void LogoForm2()
        {
            Form LogoForm2 = new Form();
            // LogoForm.BackColor = System.Drawing.Color.White;
            LogoForm2.BackgroundImage = System.Drawing.Image.FromFile("Yasnovka_DLWALVE.png");
            LogoForm2.AutoScaleMode = AutoScaleMode.Font;
            LogoForm2.Size = new Size(430, 107);
            LogoForm2.Text = "Установка длины волны...";
            LogoForm2.MinimizeBox = false;
            LogoForm2.MaximizeBox = false;
            LogoForm2.AutoSize = false;
            LogoForm2.Name = "LogoFrm2";
            LogoForm2.ShowInTaskbar = false;
            LogoForm2.StartPosition = FormStartPosition.CenterScreen;
            LogoForm2.ControlBox = false;
            LogoForm2.FormBorderStyle = FormBorderStyle.None;
            /*PictureBox PicBox = new PictureBox();
            PicBox.Size = new Size(307, 179);
            PicBox.Location = new System.Drawing.Point(12, 12);
            PicBox.ImageLocation = "D:\\Analis-samo\\Analis200\\Analis200\bin\x64\\Release\\Calibrovka.png";
            PicBox.SizeMode = PictureBoxSizeMode.Zoom;
            LogoForm.Controls.Add(PicBox);*/
            LogoForm2.Show();
        }
        public void timer1_Tick(object sender, System.EventArgs e)
        {
            GE4();
        }
        public void Analis__Activated()
        {
            newPort.Write("CO\r");

        }
        public void длинаволныToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SWForm _SWForm = new SWForm(this);
            _SWForm.ShowDialog();
            //GW();
            //Thread.Sleep(20000);
            SAGE(ref countSA, ref GE5_1_0);
            wavelength1 = GWNew.Text;
        }
        public void CO()
        {

            string b = "";
            int byteRecieved = newPort.ReadBufferSize;
            Thread.Sleep(500);
            byte[] buffer = new byte[byteRecieved];
            newPort.Read(buffer, 0, byteRecieved);

            string GW1 = "";

            for (int i = 0; i <= 50; i++)
            {
                GW1 = GW1 + Convert.ToChar(buffer[i]);
            }
            var GWarr = GW1.Split("\r".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            // GW1 = Convert.ToString(GWbuffer[0]);

            GW1_2 = GWarr[2];
            GWNew.Text = GW1_2;
            versionPribor = GWarr[1];
        }



        public void GE4()
        {
            if (Izmerenie1 == false)
            {
                newPort.Write("GE 1\r");
                string GE5 = "";
                int GEbyteRecieved4_1 = newPort.ReadBufferSize;
                byte[] GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

                string indata = newPort.ReadExisting();

                string indata_0 = "";
                bool indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {

                        indata = newPort.ReadExisting();
                        indata_0 += indata;
                    }
                }

                Regex regex = new Regex(@"\W");
                GE5 = regex.Replace(indata_0, "");


                GEText.Text = GE5;

                double GAText1 = (Convert.ToDouble(GE5) / Convert.ToDouble(GE5_1_0)) * 100;

                GAText.Text = string.Format("{0:0.00}", GAText1);
                //double GAText2 = (Convert.ToDouble(GE3) / Convert.ToDouble(GE5)) * 0.01;
                double OptPlot = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5));
                // double OptPlot1 = (OptPlot - Math.Truncate(OptPlot))*500;
                // OptPlot1 = (OptPlot1 - Math.Truncate(OptPlot1))*10;
                //OptPlot1 = Math.Truncate(OptPlot1)/5000;
                double OptPlot1 = OptPlot - Math.Truncate(OptPlot);
                OptichPlot.Text = string.Format("{0:0.0000}", OptPlot1);
            }
        }

        public void RD()
        {
            newPort.Write("RD\r");

            Thread.Sleep(500);
            //  byte[] buffer1 = new byte[byteRecieved1];
            string indata = newPort.ReadExisting();

            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = newPort.ReadExisting();
                }
            }


        }


        private void Analis_FormClosed_1(object sender, FormClosedEventArgs e)
        {
            //  newPort.Write("QU\r");
        }

        private void Analis_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            if (ComPort == true)
            {
                char[] ClosePribor = { Convert.ToChar('Q'), Convert.ToChar('U'), Convert.ToChar('\r') };
                newPort.Write("QU\r");
                Thread.Sleep(500);
                //  byte[] buffer1 = new byte[byteRecieved1];
                string indata = newPort.ReadExisting();

                bool indata_bool = true;
                while (indata_bool == true)
                {
                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();
                    }
                }

                newPort.Close();
            }
            else
            {
                SWF.Application.Exit();
            }
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }



        private void волновойАнализToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Analiz _Analiz = new Analiz(this);
            _Analiz.ShowDialog();
            // tabControl1.SelectTab(tabPage1);
        }

        public void WLADD()
        {
            MessageBox.Show(Convert.ToString(countWL));
            for (int i = 0; i <= countWL - 1; i++)
            {

                DataGridViewTextBoxColumn firstColumn =
                new DataGridViewTextBoxColumn();
                firstColumn.HeaderText = "WL (" + textBox[i].Text + ")";
                firstColumn.Name = "WL (" + textBox[i].Text + ")";
                //Table.Columns.Add(firstColumn);

            }
        }

        public void WLREMOVE()
        {
            while (true)
            {
                // int i = Table.Columns.Count - 1;//С какого столбца начать
                // if (Table.Columns[i].Name == "NameCol")
                //     break;
                // Table.Columns.RemoveAt(i);
            }

        }

        int count = 0;

        public string[] Stroka;
        // public int[] mass1;

        public string GE5_1_1 = "";
        public string GE5_1_0_1 = "";
        public int resulcCountSA1;


        public int[] mass0;
        public int[] massSA;
        public void калибровкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Calibrovka();
            Calibrovka(ref mass0, ref massSA);

            Izmerenie1 = true;


            // MessageBox.Show(Convert.ToString(mass0[0]));
            // MessageBox.Show(Convert.ToString(massSA[0]));
            //  MessageBox.Show(Convert.ToString(mass0[1]));
            //MessageBox.Show(Convert.ToString(massSA[1]));
        }


        public void Calibrovka(ref int[] mass0, ref int[] massSA)
        {

            mass0 = new int[countWL];
            massSA = new int[countWL];
            for (int i = 0; i <= countWL - 1; i++)
            {
                int countSA1;
                if (versionPribor.Contains("2"))
                { countSA1 = 8; }
                else
                {
                    countSA1 = 4;
                }
                string SWAnalis = textBox[i].Text;
                newPort.Write("SW " + SWAnalis + "\r");
                string indata = newPort.ReadExisting();

                bool indata_bool = true;
                while (indata_bool == true)
                {
                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();
                    }
                }

                newPort.Write("RD\r");
                indata = newPort.ReadExisting();

                indata_bool = true;
                while (indata_bool == true)
                {
                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();
                    }
                }


                newPort.Write("SA " + countSA1 + "\r");
                indata = newPort.ReadExisting();
                string indata_0;
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();

                    }
                }

                newPort.Write("GE 1\r");

                string GE5Izmer = "";
                int GEbyteRecieved4_1 = newPort.ReadBufferSize;
                byte[] GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

                indata = newPort.ReadExisting();

                indata_0 = "";
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {

                        indata = newPort.ReadExisting();
                        indata_0 += indata;
                    }
                }
                Regex regex = new Regex(@"\W");
                GE5_1_1 = regex.Replace(indata_0, "");


                //Thread.Sleep(1500);

                while (Convert.ToInt32(GE5_1_1) > 30000 && countSA1 > 1)
                {
                    countSA1--;
                    newPort.Write("SA " + countSA1 + "\r");
                    indata = newPort.ReadExisting();
                    indata_0 = "";
                    indata_bool = true;
                    while (indata_bool == true)
                    {

                        if (indata.Contains(">"))
                        {

                            indata_bool = false;

                        }

                        else {
                            indata = newPort.ReadExisting();

                        }
                    }

                    newPort.Write("GE 1\r");

                    GE5Izmer = "";
                    GEbyteRecieved4_1 = newPort.ReadBufferSize;
                    GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                    newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

                    indata = newPort.ReadExisting();

                    indata_0 = "";
                    indata_bool = true;
                    while (indata_bool == true)
                    {

                        if (indata.Contains(">"))
                        {

                            indata_bool = false;

                        }

                        else {

                            indata = newPort.ReadExisting();
                            indata_0 += indata;
                        }
                    }
                    regex = new Regex(@"\W");
                    GE5_1_1 = regex.Replace(indata_0, "");


                    GE5_1_0_1 = regex.Replace(indata_0, "");
                    resulcCountSA1 = countSA1;

                }
                //Thread.Sleep(2500);
                mass0[i] = Convert.ToInt32(GE5_1_0_1);
                massSA[i] = Convert.ToInt32(resulcCountSA1);
            }
        }


        public void измеритьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Thread.Sleep(100);
            Izmerenie();
        }

        public void Izmerenie()
        {
            MessageBox.Show(Convert.ToString(massSA[0]));
            MessageBox.Show(Convert.ToString(mass0[0]));
            string[] mass1 = new string[countWL];
            for (int i = 0; i <= countWL - 1; i++)
            {

                string SWAnalis = textBox[i].Text;
                newPort.Write("SW " + SWAnalis + "\r");

                Thread.Sleep(500);

                string indata = newPort.ReadExisting();

                bool indata_bool = true;
                while (indata_bool == true)
                {
                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();
                    }
                }

                newPort.Write("SA " + massSA[i] + "\r");
                indata = newPort.ReadExisting();
                string indata_0;
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();

                    }
                }

                newPort.Write("GE 1\r");

                string GE15_1 = "";
                int GEbyteRecieved4_1 = newPort.ReadBufferSize;
                byte[] GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

                indata = newPort.ReadExisting();

                indata_0 = "";
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {

                        indata = newPort.ReadExisting();
                        indata_0 += indata;
                    }
                }
                Regex regex = new Regex(@"\W");
                GE15_1 = regex.Replace(indata_0, "");


                double GAText1_1 = (Convert.ToDouble(GE15_1) / Convert.ToDouble(mass0[i])) * 100;

                double OptPlot1 = Math.Log10(1 / GAText1_1);

                double OptPlot1_1 = OptPlot1 - Math.Truncate(OptPlot1);

                mass1[i] = string.Format("{0:0.0000}", OptPlot1_1);



            }
            // Table.Rows.Add();
            //  Table.Rows[count].Cells["No"].Value = count;
            //  Table.Rows[count].Cells["NameCol"].Value = "Измерение" + count;
            for (int i = 0; i <= countWL - 1; i++)
            {
                for (int j = 0; j <= countWL - 1; j++)
                {
                    //Table.Rows.Add(count, "Измерение" + count);

                    //          Table.Rows[count].Cells["WL (" + textBox[i].Text + ")"].Value = mass1[j];
                }

            }
            count++;
            // _Analis.GW();
        }
        private void экспортToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                ExportToExcel();
            }
            else
            {
                ExportToExcel2();
            }
        }
        public void ExportToExcel()
        {
            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as Excel File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
                //Excel.Application exApp = new Excel.Application();
                exApp.Application.Workbooks.Add(Type.Missing);

                exApp.Columns.ColumnWidth = 20;
                for (int i = 1; i < this.Table1.Columns.Count + 1; i++)
                {
                    exApp.Cells[1, i] = this.Table1.Columns[i - 1].HeaderText;
                }
                Thread.Sleep(500);
                for (int i = 0; i < this.Table1.Rows.Count; i++)
                {
                    Thread.Sleep(2000);
                    for (int j = 0; j < this.Table1.Columns.Count; j++)
                    {
                        exApp.Cells[i + 2, j + 1] = this.Table1.Rows[i].Cells[j].Value;
                    }
                }

                exApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                exApp.ActiveWorkbook.Saved = true;
                exApp.Visible = true;
                //  exApp.Quit();

            }
        }


        public void ExportToExcel2()
        {

            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as Excel File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
                //Excel.Application exApp = new Excel.Application();
                exApp.Application.Workbooks.Add(Type.Missing);

                exApp.Columns.ColumnWidth = 20;
                for (int i = 1; i < this.Table2.Columns.Count + 1; i++)
                {
                    exApp.Cells[1, i] = this.Table2.Columns[i - 1].HeaderText;
                }
                Thread.Sleep(500);
                for (int i = 0; i < this.Table2.Rows.Count; i++)
                {
                    Thread.Sleep(2000);
                    for (int j = 0; j < this.Table2.Columns.Count; j++)
                    {
                        exApp.Cells[i + 2, j + 1] = this.Table2.Rows[i].Cells[j].Value;
                    }
                }

                exApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                exApp.ActiveWorkbook.Saved = true;
                // exApp.Quit();
                exApp.Visible = true;
            }
        }
        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                ExportToPDF1();
            }
            else
            {
                ExportToPDF();
            }
        }
        public void ExportToPDF()
        {
            string head = @"Протокол выполнения измерений";

            //Creating iTextSharp Table from the DataTable data
            PdfPTable pdfTable2 = new PdfPTable(Table2.ColumnCount);

            pdfTable2.DefaultCell.Padding = 5;
            pdfTable2.WidthPercentage = 100;

            pdfTable2.HorizontalAlignment = Element.ALIGN_LEFT;
            pdfTable2.DefaultCell.BorderWidth = 1;
            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\georgia.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 12f, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font fontBold = new iTextSharp.text.Font(baseFont, 18f, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font fontBold1 = new iTextSharp.text.Font(baseFont, 10f, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font font1 = new iTextSharp.text.Font(baseFont, 5f, iTextSharp.text.Font.BOLD);
            // iTextSharp.text.Font fontLeft = new iTextSharp.text.Font(baseFont, 9f, iTextSharp.text.Font.NORMAL);

            //Adding Header row
            foreach (DataGridViewColumn column in Table2.Columns)
            {

                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText, fontBold1));
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                //cell.BackgroundColor = new iTextSharp.text.Color(0, 168, 107);
                cell.Padding = 1;
                cell.PaddingBottom = 5;
                pdfTable2.AddCell(cell);

            }

            //Adding DataRow
            foreach (DataGridViewRow row in Table2.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    pdfTable2.AddCell(new Phrase(Convert.ToString(cell.Value), font));


                }
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Pdf File |*.pdf";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Document doc = new Document(PageSize.A4, 50f, 20f, 20f, 0f);
                PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));

                doc.Open();
                //iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance("Image.jpeg");

                Paragraph welcomeParagraph = new Paragraph("Протокол выполнения измерений\n", fontBold);
                welcomeParagraph.Alignment = iTextSharp.text.Element.ALIGN_CENTER;

                Paragraph FileName2 = new Paragraph("Имя файла: " + filepath2, font);
                Paragraph Description2 = new Paragraph("Описание: " + textBox8.Text, font);


                Paragraph DateTime2 = new Paragraph("Дата: " + dateTimePicker2.Value.ToString("dd.MM.yyyy"), font);
                Paragraph WaveLength2 = new Paragraph("Длина волны: " + wavelength1, font);
                Paragraph Pogresh2 = new Paragraph("Погрешность методики: " + textBox7.Text, font);
                Paragraph Opt_dlin_cuvet2 = new Paragraph("Оптическая длина кюветы: " + Opt_dlin_cuvet.Text, font);
                Paragraph F1 = new Paragraph("F1 = " + F1Text.Text, font);
                Paragraph F2 = new Paragraph("F2 = " + F2Text.Text, font);

                Paragraph Graduirovka2 = new Paragraph("Градуировка: ", font);
                Paragraph FileName1 = new Paragraph("Имя файла: " + filepath, font);
                Paragraph Description1 = new Paragraph("Описание: " + Description, font);
                Paragraph Date1 = new Paragraph("Дата: " + DateTime, font);
                Paragraph Date2 = new Paragraph("Действительна до: " + dateTimePicker1.Value.AddDays(Days).ToString("dd.MM.yyyy"), font);
                Paragraph Pogresh1 = new Paragraph("Погрешность методики: " + textBox3.Text, font);
                Paragraph GradYrav = new Paragraph("Градуировочное уравнение: " + label14.Text, font);
                Paragraph ND2 = new Paragraph("НД: " + ND, font);

                Paragraph DateIzmer2 = new Paragraph("Данные измерений: ", font);


                Paragraph Spectrofotometr2 = new Paragraph("Спектфотометр: ", font);
                Paragraph Model2 = new Paragraph("Модель: __________________________", font);
                Paragraph Date3 = new Paragraph("Поверка действительна до: __________________________", font);
                Paragraph SerNom2 = new Paragraph("Серийный номер: __________________________", font);
                Paragraph InventarNo2 = new Paragraph("Инветарный номер: __________________________", font);

                Paragraph Vipolnil = new Paragraph("Измерения выполнил(а): ____________________________________________________", font);
                Paragraph welcomeParagraph1 = new Paragraph("\n", fontBold);



                PdfPTable table = new PdfPTable(9);
                PdfPCell cell = new PdfPCell(DateTime2);
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                table.AddCell(cell);

                /*  cell = new PdfPCell();
                  cell.BorderWidth = 0;
                  table.AddCell(cell);*/

                cell = new PdfPCell(WaveLength2);
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                table.AddCell(cell);

                cell = new PdfPCell(Pogresh2);
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                table.AddCell(cell);

                cell = new PdfPCell(Opt_dlin_cuvet2);
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                table.AddCell(cell);

                cell = new PdfPCell(F1);
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                table.AddCell(cell);

                cell = new PdfPCell(F2);
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                table.AddCell(cell);



                PdfPTable table1 = new PdfPTable(6);
                PdfPCell cell1 = new PdfPCell(FileName1);
                cell1.BorderWidth = 0;
                cell1.Colspan = 6;
                table1.AddCell(cell1);

                /*  cell = new PdfPCell();
                  cell.BorderWidth = 0;
                  table.AddCell(cell);*/

                cell1 = new PdfPCell(Description1);
                cell1.BorderWidth = 0;
                cell1.Colspan = 6;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(Date1);
                cell1.BorderWidth = 0;
                cell1.Colspan = 2;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(Date2);
                cell1.BorderWidth = 0;
                cell1.Colspan = 2;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(Pogresh1);
                cell1.BorderWidth = 0;
                cell1.Colspan = 2;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(GradYrav);
                cell1.BorderWidth = 0;
                cell1.Colspan = 6;
                table1.AddCell(cell1);

                PdfPTable table2 = new PdfPTable(1);
                PdfPCell cell2 = new PdfPCell(table1);
                cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                cell2.BorderWidth = 0;
                //cell2.Colspan = 1;
                table2.AddCell(cell2);

                PdfPTable table3 = new PdfPTable(1);
                PdfPCell cell3 = new PdfPCell(table);
                cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                cell3.BorderWidth = 0;
                table3.AddCell(cell3);

                doc.Add(welcomeParagraph);
                doc.Add(welcomeParagraph1);
                doc.Add(FileName2);
                doc.Add(welcomeParagraph1);
                doc.Add(Description2);
                doc.Add(welcomeParagraph1);
                doc.Add(table3);
                doc.Add(welcomeParagraph1);
                doc.Add(Graduirovka2);
                doc.Add(welcomeParagraph1);
                doc.Add(table2);
                doc.Add(welcomeParagraph1);
                doc.Add(ND2);
                // doc.Add(pdfTable);
                doc.Add(welcomeParagraph1);
                doc.Add(DateIzmer2);
                doc.Add(welcomeParagraph1);
                doc.Add(pdfTable2);
                doc.Add(welcomeParagraph1);
                doc.Add(Vipolnil);
                // doc.Add(Chart_Image);


                doc.Close();
                // sfd.Visible = true;
            }

        }

        public double Asred1;
        int NCels = 2;
        public void одноволновоеИзмерениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // MessageBox.Show(NoCaIzm.ToString());
            double sum = 0.0;
            int startIndexCell = 2;
            int endIndexCell = startIndexCell + NoCaIzm;
            int rowIndex = Table1.CurrentRow.Index;

            bool doNotWrite = false;
            if (tabControl2.SelectedIndex == 0)
            {
                string SWAnalis = WL_grad1;
                string GE5Izmer = "";
                string GE5_1_1 = "";
                /*newPort.Write("SW " + SWAnalis + "\r");
                int SWAnalisByteRecieved1 = newPort.ReadBufferSize;
                Thread.Sleep(100);
                byte[] SWAnalisBuffer1 = new byte[SWAnalisByteRecieved1];
                newPort.Read(SWAnalisBuffer1, 0, SWAnalisByteRecieved1);*/

                newPort.Write("SA " + countSA + "\r");
                string indata = newPort.ReadExisting();
                string indata_0;
                bool indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();

                    }
                }

                newPort.Write("GE 1\r");

                GE5Izmer = "";
                int GEbyteRecieved4_1 = newPort.ReadBufferSize;
                byte[] GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

                indata = newPort.ReadExisting();

                indata_0 = "";
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {

                        indata = newPort.ReadExisting();
                        indata_0 += indata;
                    }
                }
                Regex regex = new Regex(@"\W");
                GE5Izmer = regex.Replace(indata_0, "");



                GEText.Text = GE5Izmer;
                // MessageBox.Show(GE5Izmer);
                double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
                double OptPlot1 = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5Izmer));
                double OptPlot1_1 = OptPlot1 - Math.Truncate(OptPlot1);
                Table1.CurrentCell.Value = string.Format("{0:0.0000}", OptPlot1_1);
                GAText.Text = string.Format("{0:0.00}", Aser);
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {
                    {
                        for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                        {
                            if (Table1.Rows[j].Cells[i].Value == null)
                            {
                                doNotWrite = true;

                                for (int l = 3; l <= Table1.Rows[Table1.CurrentCell.RowIndex].Cells.Count - 1; l++)
                                {
                                    if (Table1.Rows[rowIndex].Cells[l].Value == null)
                                    {
                                        cellnull++;
                                    }
                                }
                            }


                        }
                    }
                }



                if (!doNotWrite)
                {
                    if (Table1.Rows.Count - 1 == 1)
                    {
                        radioButton1.Enabled = true;
                        radioButton4.Enabled = true;
                        radioButton5.Enabled = true;
                        radioButton3.Enabled = false;
                        radioButton2.Enabled = false;
                    }
                    if (Table1.Rows.Count - 1 == 2)
                    {
                        radioButton1.Enabled = true;
                        radioButton2.Enabled = true;
                        radioButton4.Enabled = true;
                        radioButton5.Enabled = true;
                        radioButton3.Enabled = false;
                    }
                    if (Table1.Rows.Count - 1 >= 3)
                    {
                        radioButton1.Enabled = true;
                        radioButton2.Enabled = true;
                        radioButton3.Enabled = true;
                        radioButton4.Enabled = true;
                        radioButton5.Enabled = true;
                    }
                    sum = 0.0;
                    while (true)
                    {
                        int i = Table1.Columns.Count - 1;//С какого столбца начать
                        if (Table1.Columns.Count == 3 + Convert.ToInt32(CountSeriya2))
                            break;
                        Table1.Columns.RemoveAt(i);
                    }

                    for (int l = startIndexCell + NoCaIzm; l <= endIndexCell; ++l)
                    {
                        if (Table1.Rows[rowIndex].Cells[l].Value == null)
                        {
                            cellnull++;
                        }

                        else
                        {
                            for (int j = 0; j < Table1.Rows.Count - 1; j++)
                            {

                                for (int i1 = startIndexCell + 1; i1 <= endIndexCell; ++i1)
                                {
                                    sum += Convert.ToDouble(Table1.Rows[j].Cells[i1].Value);
                                    Asred1 = sum / NoCaIzm;
                                    // MessageBox.Show(Convert.ToString(Asred1));
                                    Table1.Rows[j].Cells["Asred"].Value = string.Format("{0:0.0000}", Asred1);

                                }
                                sum = 0.0;
                            }
                        }
                        Izmerenie1 = true;
                    }
                    for (int m = 0; m < Table1.Rows.Count - 1; m++)
                    {
                        for (int ml = 0; ml < Table1.Rows[m].Cells.Count; ml++)
                        {
                            if (Table1.Rows[m].Cells[ml].Value == null)
                            { doNotWrite = true; }
                        }
                    }
                    if (!doNotWrite)
                    {
                        while (true)
                        {
                            int ml = Table1.Columns.Count - 1;//С какого столбца начать
                            if (Table1.Columns.Count == 3 + NoCaIzm)
                                break;
                            Table1.Columns.RemoveAt(ml);
                        }
                        functionAsred();

                    }

                }
            }
            else
            {
                ///Измерение
                ///

                int rowIndex2 = Table2.CurrentRow.Index;
                bool doNotWrite1 = false;
                double maxEl;
                double minEl;
                double serValue = 0;
                int cellnull = 0;
                El = new double[NoCaIzm1 + 1];

                double SredValue = 0;
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {
                    {
                        for (int i = 0; i < Table1.Rows[j].Cells.Count; i++)
                        {
                            if (Table1.Rows[j].Cells[i].Value == null)
                            {
                                doNotWrite = true;
                                for (int i1 = startIndexCell; i1 <= startIndexCell + NoCaIzm1 + NoCaIzm1; i1 += 2)
                                {
                                    string SWAnalis = WL_grad1;
                                    string GE5Izmer = "";
                                    string GE5_1_1 = "";
                                    /*  newPort.Write("SW " + SWAnalis + "\r");
                                      int SWAnalisByteRecieved1 = newPort.ReadBufferSize;
                                      Thread.Sleep(100);
                                      byte[] SWAnalisBuffer1 = new byte[SWAnalisByteRecieved1];
                                      newPort.Read(SWAnalisBuffer1, 0, SWAnalisByteRecieved1);*/

                                    newPort.Write("SA " + countSA + "\r");
                                    string indata = newPort.ReadExisting();
                                    string indata_0;
                                    bool indata_bool = true;
                                    while (indata_bool == true)
                                    {

                                        if (indata.Contains(">"))
                                        {

                                            indata_bool = false;

                                        }

                                        else {
                                            indata = newPort.ReadExisting();

                                        }
                                    }

                                    newPort.Write("GE 1\r");

                                    GE5Izmer = "";
                                    int GEbyteRecieved4_1 = newPort.ReadBufferSize;
                                    byte[] GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                                    newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

                                    indata = newPort.ReadExisting();

                                    indata_0 = "";
                                    indata_bool = true;
                                    while (indata_bool == true)
                                    {

                                        if (indata.Contains(">"))
                                        {

                                            indata_bool = false;

                                        }

                                        else {

                                            indata = newPort.ReadExisting();
                                            indata_0 += indata;
                                        }
                                    }
                                    Regex regex = new Regex(@"\W");
                                    GE5Izmer = regex.Replace(indata_0, "");
                                    GEText.Text = GE5Izmer;
                                    // MessageBox.Show(GE5Izmer);
                                    double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
                                    double OptPlot1 = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5Izmer));
                                    double OptPlot1_1 = OptPlot1 - Math.Truncate(OptPlot1);
                                    Table2.CurrentCell.Value = string.Format("{0:0.0000}", OptPlot1_1);
                                    GAText.Text = string.Format("{0:0.00}", Aser);

                                    if (Table2.Rows[rowIndex2].Cells[i1].Value != null)
                                    {

                                        if (aproksim == "Линейная через 0")
                                        {
                                            serValue = Convert.ToDouble(Table2.Rows[rowIndex2].Cells[i1].Value.ToString()) / Convert.ToDouble(textBox5.Text);
                                        }
                                        if (aproksim == "Линейная")
                                        {
                                            serValue = ((Convert.ToDouble(Table2.Rows[rowIndex2].Cells[i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / Convert.ToDouble(textBox5.Text);
                                        }
                                        if (aproksim == "Квадратичная")
                                        {
                                            serValue = ((Convert.ToDouble(Table2.Rows[rowIndex2].Cells[i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / (Convert.ToDouble(textBox5.Text) + Convert.ToDouble(textBox6.Text));
                                        }
                                        double CValue1 = Convert.ToDouble(F1Text.Text);
                                        double CValue2 = Convert.ToDouble(F2Text.Text);

                                        Table2.Rows[rowIndex2].Cells[i1 + 1].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);

                                        SredValue += Convert.ToDouble(Table2.Rows[rowIndex2].Cells[i1 + 1].Value.ToString());


                                        double CCR = SredValue / NoCaIzm1;
                                        if (Convert.ToDouble(textBox7.Text) >= 1)
                                        {
                                            Table2.Rows[rowIndex2].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.0000}", (CCR / Convert.ToDouble(textBox7.Text)));
                                        }
                                        else Table2.Rows[rowIndex2].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                                        //Table2.Rows[j].Cells["d%"].Value = El.Max();
                                        El[cellnull] = Convert.ToDouble(Table2.Rows[rowIndex2].Cells[i1 + 1].Value.ToString());

                                        cellnull++;
                                    }
                                    //NCels++;
                                    Izmerenie1 = true;
                                    // установка типа графика

                                }
                                for (int j1 = 0; j1 < Table2.Rows.Count - 1; j1++)
                                    for (int i1 = 2; i1 <= Table2.Rows[j1].Cells.Count - 2; i1++)
                                        if (Table2.Rows[j1].Cells[i1].Value == null)
                                            doNotWrite = true;
                                if (!doNotWrite)
                                {
                                    Array.Sort(El);
                                    maxEl = El[El.Length - 1];
                                    minEl = El[1];
                                    double a = ((maxEl - minEl) * 100) / Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["Ccr"].Value);
                                    double b = a;

                                    Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = string.Format("{0:0.0000}", b);

                                }
                            }
                        }
                    }
                }
                if (!doNotWrite)
                {
                    sum = 0.0;
                    Table2.CurrentCell.Value = null;
                    serValue = 0;
                    cellnull = 0;
                    El = new double[NoCaIzm1 + 1];

                    SredValue = 0;
                    for (int i1 = startIndexCell; i1 <= startIndexCell + NoCaIzm1 + NoCaIzm1; i1 += 2)
                    {
                        string SWAnalis = WL_grad1;
                        string GE5Izmer = "";
                        string GE5_1_1 = "";
                        /* newPort.Write("SW " + SWAnalis + "\r");
                         int SWAnalisByteRecieved1 = newPort.ReadBufferSize;
                         Thread.Sleep(100);
                         byte[] SWAnalisBuffer1 = new byte[SWAnalisByteRecieved1];
                         newPort.Read(SWAnalisBuffer1, 0, SWAnalisByteRecieved1);*/

                        newPort.Write("SA " + countSA + "\r");
                        string indata = newPort.ReadExisting();
                        string indata_0;
                        bool indata_bool = true;
                        while (indata_bool == true)
                        {

                            if (indata.Contains(">"))
                            {

                                indata_bool = false;

                            }

                            else {
                                indata = newPort.ReadExisting();

                            }
                        }

                        newPort.Write("GE 1\r");

                        GE5Izmer = "";
                        int GEbyteRecieved4_1 = newPort.ReadBufferSize;
                        byte[] GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                        newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

                        indata = newPort.ReadExisting();

                        indata_0 = "";
                        indata_bool = true;
                        while (indata_bool == true)
                        {

                            if (indata.Contains(">"))
                            {

                                indata_bool = false;

                            }

                            else {

                                indata = newPort.ReadExisting();
                                indata_0 += indata;
                            }
                        }
                        Regex regex = new Regex(@"\W");
                        GE5Izmer = regex.Replace(indata_0, "");
                        GEText.Text = GE5Izmer;
                        // MessageBox.Show(GE5Izmer);
                        double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
                        double OptPlot1 = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5Izmer));
                        double OptPlot1_1 = OptPlot1 - Math.Truncate(OptPlot1);
                        Table2.CurrentCell.Value = string.Format("{0:0.0000}", OptPlot1_1);
                        GAText.Text = string.Format("{0:0.00}", Aser);

                        if (Table2.Rows[rowIndex2].Cells[i1].Value != null)
                        {

                            if (aproksim == "Линейная через 0")
                            {
                                serValue = Convert.ToDouble(Table2.Rows[rowIndex2].Cells[i1].Value.ToString()) / Convert.ToDouble(textBox5.Text);
                            }
                            if (aproksim == "Линейная")
                            {
                                serValue = ((Convert.ToDouble(Table2.Rows[rowIndex2].Cells[i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / Convert.ToDouble(textBox5.Text);
                            }
                            if (aproksim == "Квадратичная")
                            {
                                serValue = ((Convert.ToDouble(Table2.Rows[rowIndex2].Cells[i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / (Convert.ToDouble(textBox5.Text) + Convert.ToDouble(textBox6.Text));
                            }
                            double CValue1 = Convert.ToDouble(F1Text.Text);
                            double CValue2 = Convert.ToDouble(F2Text.Text);

                            Table2.Rows[rowIndex2].Cells[i1 + 1].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);

                            SredValue += Convert.ToDouble(Table2.Rows[rowIndex2].Cells[i1 + 1].Value.ToString());


                            double CCR = SredValue / NoCaIzm1;
                            if (Convert.ToDouble(textBox7.Text) >= 1)
                            {
                                Table2.Rows[rowIndex2].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.0000}", (CCR / Convert.ToDouble(textBox7.Text)));
                            }
                            else Table2.Rows[rowIndex2].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                            //Table2.Rows[j].Cells["d%"].Value = El.Max();
                            El[cellnull] = Convert.ToDouble(Table2.Rows[rowIndex2].Cells[i1 + 1].Value.ToString());

                            cellnull++;
                        }
                        //NCels++;
                        Izmerenie1 = true;
                        // установка типа графика

                    }
                    for (int j1 = 0; j1 < Table2.Rows.Count - 1; j1++)
                        for (int i1 = 2; i1 <= Table2.Rows[j1].Cells.Count - 2; i1++)
                            if (Table2.Rows[j1].Cells[i1].Value == null)
                                doNotWrite = true;
                    if (!doNotWrite)
                    {
                        Array.Sort(El);
                        maxEl = El[El.Length - 1];
                        minEl = El[1];
                        double a = ((maxEl - minEl) * 100) / Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["Ccr"].Value);
                        double b = a;
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = string.Format("{0:0.0000}", b);

                    }
                }
            }
        }
        public string GE5Calibr;
        public void калибровкаДляОдноволновогоАнализаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // string GE5Calibr = "";
            //  CalibrovkaGrad(ref countSA);
            //   MessageBox.Show(GE5Calibr);
            SAGE(ref countSA, ref GE5_1_0);
        }
        public void CalibrovkaGrad(ref int countSA)
        {
            string GE5Calibr1 = "";
            string SWAnalis = WL_grad1;
            if (versionPribor.Contains("2"))
            { countSA = 8; }
            else
            {
                countSA = 4;
            }
            newPort.Write("SW " + SWAnalis + "\r");
            Thread.Sleep(500);
            //  byte[] buffer1 = new byte[byteRecieved1];
            string indata = newPort.ReadExisting();

            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = newPort.ReadExisting();
                }
            }

            newPort.Write("RD\r");
            Thread.Sleep(500);
            //  byte[] buffer1 = new byte[byteRecieved1];
            indata = newPort.ReadExisting();

            indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = newPort.ReadExisting();
                }
            }



            newPort.Write("SA " + countSA + "\r");
            indata = newPort.ReadExisting();
            string indata_0;
            indata_bool = true;
            while (indata_bool == true)
            {

                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = newPort.ReadExisting();

                }
            }

            newPort.Write("GE 1\r");

            string GE5Izmer = "";
            int GEbyteRecieved4_1 = newPort.ReadBufferSize;
            byte[] GEbuffer4_1 = new byte[GEbyteRecieved4_1];
            newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

            indata = newPort.ReadExisting();

            indata_0 = "";
            indata_bool = true;
            while (indata_bool == true)
            {

                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {

                    indata = newPort.ReadExisting();
                    indata_0 += indata;
                }
            }
            Regex regex = new Regex(@"\W");
            GE5Calibr1 = regex.Replace(indata_0, "");


            //Thread.Sleep(1500);

            while (Convert.ToInt32(GE5Calibr1) > 30000 && countSA > 1)
            {
                countSA--;
                newPort.Write("SA " + countSA + "\r");
                indata = newPort.ReadExisting();
                indata_0 = "";
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();

                    }
                }

                newPort.Write("GE 1\r");


                GEbyteRecieved4_1 = newPort.ReadBufferSize;
                GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

                indata = newPort.ReadExisting();

                indata_0 = "";
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {

                        indata = newPort.ReadExisting();
                        indata_0 += indata;
                    }
                }
                regex = new Regex(@"\W");
                GE5Calibr1 = regex.Replace(indata_0, "");

                //GE5_1_0_1 = GEarr4_1_1_1[1];
                resulcCountSA1 = countSA;

            }
            //  GW();
            Izmerenie1 = true;
            GE5Calibr = GE5Calibr1;
        }
        public double k0;
        public double k1;
        public double k2;
        double SUM0, SUM1;
        int cellnull = 0;

        private void Table1_CancelRowEdit(object sender, QuestionEventArgs e)
        {

        }

        public void Table1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show(NoCaIzm.ToString());
            //Table1.CurrentCell.Key
            if (Table1.CurrentCell.ReadOnly == true)
            {
                MessageBox.Show("Запись запрещена!");
            }

            double sum;
            sum = 0.0;
            int startIndexCell = 2;
            int endIndexCell = startIndexCell + NoCaIzm;
            int rowIndex = Table1.CurrentRow.Index;
            bool doNotWrite = false;
            int rownull = 0;

            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                {
                    for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                    {
                        if (Table1.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;

                            for (int l = startIndexCell + NoCaIzm; l <= endIndexCell; ++l)
                            {
                                if (Table1.Rows[rowIndex].Cells[l].Value == null)
                                {
                                    cellnull++;
                                }
                            }
                        }


                    }
                }
            }



            if (!doNotWrite)
            {
                //MessageBox.Show(NoCaIzm.ToString());
                // MessageBox.Show("Таблица заполнена");
                if (Table1.Rows.Count - 1 == 1)
                {
                    radioButton1.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                    radioButton3.Enabled = false;
                    radioButton2.Enabled = false;
                }
                if (Table1.Rows.Count - 1 == 2)
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                    radioButton3.Enabled = false;
                }
                if (Table1.Rows.Count - 1 >= 3)
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton3.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                }

                sum = 0.0;
                /*while (true)
                 {
                     int i = Table1.Columns.Count - 1;//С какого столбца начать
                     if (Table1.Columns.Count == 3 + Convert.ToInt32(CountSeriya2))
                         break;
                     //Table1.Columns.RemoveAt(i);
                 }*/

                for (int l = startIndexCell + NoCaIzm; l <= endIndexCell; ++l)
                {
                    if (Table1.Rows[rowIndex].Cells[l].Value == null)
                    {
                        cellnull++;
                    }

                    else
                    {
                        for (int j = 0; j < Table1.Rows.Count - 1; j++)
                        {

                            for (int i1 = startIndexCell + 1; i1 <= endIndexCell; ++i1)
                            {
                                sum += Convert.ToDouble(Table1.Rows[j].Cells[i1].Value);
                                Asred1 = sum / NoCaIzm;
                                // MessageBox.Show(Convert.ToString(Asred1));
                                Table1.Rows[j].Cells["Asred"].Value = string.Format("{0:0.0000}", Asred1);

                            }
                            sum = 0.0;
                        }
                    }
                    Izmerenie1 = true;
                }
                for (int m = 0; m < Table1.Rows.Count - 1; m++)
                {
                    for (int ml = 0; ml < Table1.Rows[m].Cells.Count; ml++)
                    {
                        if (Table1.Rows[m].Cells[ml].Value == null)
                        { doNotWrite = true; }
                    }
                }
                if (!doNotWrite)
                {
                    while (true)
                    {
                        int ml = Table1.Columns.Count - 1;//С какого столбца начать
                        if (Table1.Columns.Count == 3 + NoCaIzm)
                            break;
                        Table1.Columns.RemoveAt(ml);
                    }
                    functionAsred();

                }

            }

        }
        public void functionAsred()
        {
            //Table1.Rows.Add();
            while (true)
            {
                int ml = Table1.Columns.Count - 1;//С какого столбца начать
                if (Table1.Columns.Count == 3 + NoCaIzm)
                    break;
                Table1.Columns.RemoveAt(ml);
            }
            groupBox3.Enabled = true;
            groupBox2.Enabled = true;

            if (radioButton1.Checked == true)
            {
                lineinaya0();
            }
            else
            {
                if (radioButton2.Checked == true)
                {

                    lineinaya();
                }
                else
                {
                    kvadratichnaya();
                }
            }
        }

        public void графикРезультатаОдноволновогоИзмеренияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            lineinaya0();
        }
        double SUMMX;
        double SUMMY;
        double XY;
        double SUMMY2;
        double SredX, SredY;
        public string aproksim;
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            aproksim = "Линейная через 0";
            lineinaya0();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            aproksim = "Квадратичная";
            kvadratichnaya();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            aproksim = "Линейная";
            lineinaya();
        }
        public string Zavisimoct;
        public void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            Zavisimoct = "A(C)";

            /*    while (true)
                 {
                     int i = Table1.Columns.Count - 1;//С какого столбца начать
                     if (Table1.Columns[i].Name == "Asred")
                         break;
                     Table1.Columns.RemoveAt(i);
                 }*/
            if (radioButton1.Checked == true)
            {
                /*  while (true)
                  {
                      int i = Table1.Columns.Count - 1;//С какого столбца начать
                      if (Table1.Columns[i].Name == "Asred")
                          break;
                      Table1.Columns.RemoveAt(i);
                  }
                  */
                lineinaya0();
            }
            else
            {
                if (radioButton2.Checked == true)
                {

                    lineinaya();
                }
                else
                {
                    kvadratichnaya();
                }
            }
        }


        private void файлToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        int circle;

        public void radioButton5_CheckedChanged_1(object sender, EventArgs e)
        {
            Zavisimoct = "C(A)";
            /*  while (true)
              {
                  int i = Table1.Columns.Count - 1;//С какого столбца начать
                  if (Table1.Columns[i].Name == "Asred")
                      break;
                  Table1.Columns.RemoveAt(i);
              }*/
            if (radioButton1.Checked == true)
            {


                //  Table1.Columns.Add("X*X", "Concetr*Concetr");
                // Table1.Columns.Add("X*Y", "Asred*Concetr");
                lineinaya0();
            }
            else
            {
                if (radioButton2.Checked == true)
                {
                    /* while (true)
                     {
                         int i = Table1.Columns.Count - 1;//С какого столбца начать
                         if (Table1.Columns[i].Name == "Asred")
                             break;
                         Table1.Columns.RemoveAt(i);
                     }*/
                    lineinaya();
                }
                else
                {
                    kvadratichnaya();
                }
            }
        }
        public string filepath;
        public string filepath2;
        private void CreateXMLDocument(ref string filepath)
        {

            filepath = saveFileDialog1.FileName;
            XmlTextWriter xtw = new XmlTextWriter(filepath, Encoding.UTF8);

            xtw.WriteStartDocument();
            xtw.WriteStartElement("Data_Izmerenie");
            xtw.WriteEndDocument();
            xtw.Close();
        }
        private void CreateXMLDocument2(ref string filepath2)
        {

            filepath2 = saveFileDialog1.FileName;
            XmlTextWriter xtw = new XmlTextWriter(filepath2, Encoding.UTF8);

            xtw.WriteStartDocument();
            xtw.WriteStartElement("Data_Izmerenie");
            xtw.WriteEndDocument();
            xtw.Close();
        }
        public void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                SaveAs1();

                tabPage4.Parent = tabControl2;

            }
            else
            {
                SaveAs2();
            }

        }
        public void SaveAs1()
        {
            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as XML File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "QS2 файл|*.qs2";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                CreateXMLDocument(ref filepath);
                WriteXml(ref filepath);

            }
        }
        public void SaveAs2()
        {
            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as XML File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "QA2 файл|*.qa2";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                CreateXMLDocument2(ref filepath2);
                WriteXml2(ref filepath2);

            }
        }
        public string[] HeaderCells;
        public string[,] Cells1;
        public void WriteXml(ref string filepath)
        {
            XmlDocument xd = new XmlDocument();
            FileStream fs = new FileStream(filepath, FileMode.Open);
            xd.Load(fs);

            XmlNode Izmerenie = xd.CreateElement("Izmerenie");

            XmlNode Nazvanie = xd.CreateElement("Nazvanie"); // Название вещества
            Nazvanie.InnerText = "Расчет градуировочного графика"; // и значение
            Izmerenie.AppendChild(Nazvanie); // и указываем кому принадлежит

            XmlNode Veshestvo = xd.CreateElement("Veshestvo"); // Название вещества
            Veshestvo.InnerText = Veshestvo1; // и значение
            Izmerenie.AppendChild(Veshestvo); // и указываем кому принадлежит

            XmlNode wavelength = xd.CreateElement("wavelength"); // Длина волны
            wavelength.InnerText = wavelength1; // и значение
            Izmerenie.AppendChild(wavelength); // и указываем кому принадлежит

            XmlNode WidthCuvet1 = xd.CreateElement("WidthCuvet"); // Ширина кюветы
            WidthCuvet1.InnerText = WidthCuvette; // и значение
            Izmerenie.AppendChild(WidthCuvet1); // и указываем кому принадлежит

            XmlNode BottomLine1 = xd.CreateElement("BottomLine"); // Нижняя граница
            BottomLine1.InnerText = BottomLine; // и значение
            Izmerenie.AppendChild(BottomLine1); // и указываем кому принадлежит

            XmlNode TopLine1 = xd.CreateElement("TopLine"); // Верхняя граница
            TopLine1.InnerText = TopLine; // и значение
            Izmerenie.AppendChild(TopLine1); // и указываем кому принадлежит

            XmlNode ND1 = xd.CreateElement("ND"); // НД
            ND1.InnerText = ND; // и значение
            Izmerenie.AppendChild(ND1); // и указываем кому принадлежит

            XmlNode Description1 = xd.CreateElement("Description"); // Примечание
            Description1.InnerText = Description; // и значение
            Izmerenie.AppendChild(Description1); // и указываем кому принадлежит

            XmlNode DateTime1 = xd.CreateElement("DateTime"); // дата создания градуировки
            DateTime1.InnerText = DateTime; // и значение
            Izmerenie.AppendChild(DateTime1); // и указываем кому принадлежит

            XmlNode DateTime1_1 = xd.CreateElement("DateTime1_1"); // Действительно до
            DateTime1_1.InnerText = label6.Text; // и значение
            Izmerenie.AppendChild(DateTime1_1); // и указываем кому принадлежит

            XmlNode DateTime1_1_1 = xd.CreateElement("DateTime1_1_1"); // Действительно до
            DateTime1_1_1.InnerText = numericUpDown1.Value.ToString(); // и значение
            Izmerenie.AppendChild(DateTime1_1_1); // и указываем кому принадлежит

            XmlNode Pogreshnost = xd.CreateElement("Pogreshnost"); // Действительно до
            Pogreshnost.InnerText = textBox3.Text; // и значение
            Izmerenie.AppendChild(Pogreshnost); // и указываем кому принадлежит

            XmlNode Ispolnitel1 = xd.CreateElement("Ispolnitel"); // Примечание
            Ispolnitel1.InnerText = Ispolnitel; // и значение
            Izmerenie.AppendChild(Ispolnitel1); // и указываем кому принадлежит

            XmlNode CountSeriya1 = xd.CreateElement("CountSeriyal"); // Примечание
            CountSeriya1.InnerText = CountSeriya; // и значение
            Izmerenie.AppendChild(CountSeriya1); // и указываем кому принадлежит

            XmlNode CountInSeriya1 = xd.CreateElement("CountInSeriyal"); // Примечание
            CountInSeriya1.InnerText = CountInSeriya; // и значение
            Izmerenie.AppendChild(CountInSeriya1); // и указываем кому принадлежит


            XmlNode TypeYravn1 = xd.CreateElement("TypeYravn"); // Тип уравнения
            if (radioButton1.Checked == true)
            {
                TypeYravn1.InnerText = "Линейное через 0"; // и значение
            }
            else
            {
                if (radioButton2.Checked == true)
                {

                    TypeYravn1.InnerText = "Линейное";
                }
                else
                {
                    TypeYravn1.InnerText = "Квадратичное";
                }
            }

            Izmerenie.AppendChild(TypeYravn1); // и указываем кому принадлежит

            XmlNode TypeIzmer1 = xd.CreateElement("TypeIzmer"); // Тип уравнения
            if (radioButton4.Checked == true)
            {
                TypeIzmer1.InnerText = "A (C) - градуировочное уравнение (стандарт)"; // и значение
            }
            else
            {
                TypeIzmer1.InnerText = "C (A) - расчетное уравнение (прибор)";
            }

            Izmerenie.AppendChild(TypeIzmer1); // и указываем кому принадлежит

            // ЗАбиваем запись в документ  
            xd.DocumentElement.AppendChild(Izmerenie);
            HeaderCells = new string[this.Table1.Columns.Count];
            Cells1 = new string[this.Table1.Rows.Count - 1, this.Table1.Columns.Count];

            /* for (int i = 0; i < this.Table1.Columns.Count; i++)
             {
                 HeaderCells[i] = this.Table1.Columns[i].HeaderText;
                 XmlNode HeaderCells1 = xd.CreateElement(HeaderCells[i]); // Примечание
                 HeaderCells1.InnerText = HeaderCells[i]; // и значение
                 Izmerenie.AppendChild(HeaderCells1); // и указываем кому принадлежит
             }
             */

            for (int i = 0; i < this.Table1.Rows.Count - 1; i++)
            {
                XmlNode Cells2 = xd.CreateElement("Stroka");

                XmlAttribute attribute1 = xd.CreateAttribute("Nomer");
                attribute1.Value = Convert.ToString(i); // устанавливаем значение атрибута
                Cells2.Attributes.Append(attribute1); // добавляем атрибут
                for (int j = 0; j < this.Table1.Columns.Count; j++)
                {
                    //  row[i + 2] = r.Cells[j].Value;
                    // row[this.Table1.Columns[j].HeaderText] = r.Cells[j].Value.ToString();
                    //string[][] Cells1;
                    Cells1[i, j] = Convert.ToString(this.Table1.Rows[i].Cells[j].Value);

                    HeaderCells[j] = this.Table1.Columns[j].HeaderText;
                    XmlNode HeaderCells1 = xd.CreateElement("Stolbec"); // Столбец
                    HeaderCells1.InnerText = Cells1[i, j]; // и значение
                    Cells2.AppendChild(HeaderCells1); // и указываем кому принадлежит
                    XmlAttribute attribute = xd.CreateAttribute("Header");
                    attribute.Value = HeaderCells[j]; // устанавливаем значение атрибута
                    HeaderCells1.Attributes.Append(attribute); // добавляем атрибут                    
                }
                xd.DocumentElement.AppendChild(Cells2);
            }


            //ds.WriteXml(filepath);

            fs.Close();         // Закрываем поток  
            xd.Save(filepath); // Сохраняем файл  

        }


        public void WriteXml2(ref string filepath2)
        {
            XmlDocument xd = new XmlDocument();
            FileStream fs = new FileStream(filepath2, FileMode.Open);
            xd.Load(fs);

            XmlNode Izmerenie = xd.CreateElement("Izmerenie");

            XmlNode Nazvanie = xd.CreateElement("Nazvanie"); // Название вещества
            Nazvanie.InnerText = "Измерения"; // и значение
            Izmerenie.AppendChild(Nazvanie); // и указываем кому принадлежит


            XmlNode WidthCuvet1 = xd.CreateElement("WidthCuvet"); // Ширина кюветы
            WidthCuvet1.InnerText = Opt_dlin_cuvet.Text; // и значение
            Izmerenie.AppendChild(WidthCuvet1); // и указываем кому принадлежит

            XmlNode Description1 = xd.CreateElement("Description"); // Примечание
            Description1.InnerText = textBox8.Text; // и значение
            Izmerenie.AppendChild(Description1); // и указываем кому принадлежит

            XmlNode DateTime1 = xd.CreateElement("DateTime"); // дата создания градуировки
            DateTime1.InnerText = dateTimePicker2.Text; // и значение
            Izmerenie.AppendChild(DateTime1); // и указываем кому принадлежит

            XmlNode Pogreshnost = xd.CreateElement("Pogreshnost"); // Погрешность
            Pogreshnost.InnerText = textBox7.Text; // и значение
            Izmerenie.AppendChild(Pogreshnost); // и указываем кому принадлежит

            XmlNode F1 = xd.CreateElement("F1"); // F1
            F1.InnerText = F1Text.Text; // и значение
            Izmerenie.AppendChild(F1); // и указываем кому принадлежит

            XmlNode F2 = xd.CreateElement("F2"); // ДF2
            F2.InnerText = F2Text.Text; // и значение
            Izmerenie.AppendChild(F2); // и указываем кому принадлежит

            XmlNode CountSeriya1 = xd.CreateElement("CountSeriyal"); // Примечание
            CountSeriya1.InnerText = Convert.ToString(NoCaIzm1); // и значение
            Izmerenie.AppendChild(CountSeriya1); // и указываем кому принадлежит

            XmlNode CountInSeriya1 = xd.CreateElement("CountInSeriyal"); // Примечание
            CountInSeriya1.InnerText = Convert.ToString(NoCaSer1); // и значение
            Izmerenie.AppendChild(CountInSeriya1); // и указываем кому принадлежит

            // ЗАбиваем запись в документ  
            xd.DocumentElement.AppendChild(Izmerenie);
            HeaderCells = new string[this.Table2.Columns.Count];
            Cells1 = new string[this.Table2.Rows.Count - 1, this.Table2.Columns.Count];

            for (int i = 0; i < this.Table2.Rows.Count - 1; i++)
            {
                XmlNode Cells2 = xd.CreateElement("Stroka");

                XmlAttribute attribute1 = xd.CreateAttribute("Nomer");
                attribute1.Value = Convert.ToString(i); // устанавливаем значение атрибута
                Cells2.Attributes.Append(attribute1); // добавляем атрибут
                for (int j = 0; j < this.Table2.Columns.Count; j++)
                {

                    Cells1[i, j] = Convert.ToString(this.Table2.Rows[i].Cells[j].Value);

                    HeaderCells[j] = this.Table2.Columns[j].HeaderText;
                    XmlNode HeaderCells1 = xd.CreateElement("Stolbec"); // Столбец
                    if (Cells1[i, j] != "")
                    {
                        HeaderCells1.InnerText = Cells1[i, j]; // и значение
                    }
                    else
                    {
                        HeaderCells1.InnerText = "-";
                    }
                    Cells2.AppendChild(HeaderCells1); // и указываем кому принадлежит
                    XmlAttribute attribute = xd.CreateAttribute("Header");
                    attribute.Value = HeaderCells[j]; // устанавливаем значение атрибута
                    HeaderCells1.Attributes.Append(attribute); // добавляем атрибут                    
                }
                xd.DocumentElement.AppendChild(Cells2);
            }

            fs.Close();         // Закрываем поток  
            xd.Save(filepath2); // Сохраняем файл  

        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                openFileDialog1.InitialDirectory = "C";
                openFileDialog1.Title = "Open File";
                openFileDialog1.FileName = "";
                openFileDialog1.Filter = "QS2 файл|*.qs2";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    try
                    {
                        // получаем выбранный файл
                        openFile(ref filepath);
                    }
                    catch (Exception t) { MessageBox.Show("exeption" + t.Message); }


                    TableWrite();
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton3.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                }
                tabPage4.Parent = tabControl2;
            }
            else
            {
                openFileDialog1.InitialDirectory = "C";
                openFileDialog1.Title = "Open File";
                openFileDialog1.FileName = "";
                openFileDialog1.Filter = "QA2 файл|*.qa2";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // получаем выбранный файл
                        openFile2(ref filepath2);
                    }
                    catch (Exception t) { MessageBox.Show("exeption" + t.Message); }


                    TableWrite2();
                }

            }


        }
        public string Veshestvo2;
        public string wavelength2;
        public string WidthCuvette2 = "";
        public string BottomLine2 = "";
        public string TopLine2 = "";
        public string ND2 = "";
        public string Description2 = "";
        public string DateTime2 = "";
        public string DateTime2_1 = "";
        public string DateTime2_2_1 = "";
        public string Ispolnitel2 = "";
        public string CountSeriya2 = Convert.ToString(3);
        public string CountInSeriya2 = Convert.ToString(3);
        public string[,] Stolbec;
        public string Stolbec1 = "";
        public string Stroka1 = "";
        public string Pogreshnost2 = "";
        public string TypeYravn1 = "";
        public string TimeIzmer1 = "";
        public void openFile(ref string filepath)
        {
            WLREMOVE1();
            WLREMOVESTR1();
            filepath = openFileDialog1.FileName;
            параметрыToolStripMenuItem.Enabled = true;
            button10.Enabled = true;

            XmlDocument xDoc = new XmlDocument();

            xDoc.Load(@filepath);

            XmlNodeList nodes = xDoc.ChildNodes;
            // Обходим значения
            foreach (XmlNode n in nodes)
            { // Обрабатываем в цикле только Data_Izmerenie
                if ("Data_Izmerenie".Equals(n.Name))
                {
                    // Читаем в цикле вложенные значения Izmerenie
                    for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                    {
                        // Обрабатываем в цикле только Izmerenie
                        if ("Izmerenie".Equals(d.Name))
                        {
                            //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                            for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                            {
                                if ("Veshestvo".Equals(k.Name) && k.FirstChild != null)
                                {
                                    Veshestvo1 = k.FirstChild.Value; //Вещество

                                }
                                if ("wavelength".Equals(k.Name) && k.FirstChild != null)
                                {
                                    wavelength1 = k.FirstChild.Value; //Длина волны


                                }
                                if ("WidthCuvet".Equals(k.Name) && k.FirstChild != null)
                                {
                                    WidthCuvette = k.FirstChild.Value; //Ширина кюветы
                                    textBox2.Text = WidthCuvette;
                                }
                                if ("BottomLine".Equals(k.Name) && k.FirstChild != null)
                                {
                                    BottomLine = k.FirstChild.Value; //Нижняя граница
                                }
                                if ("TopLine".Equals(k.Name) && k.FirstChild != null)
                                {
                                    TopLine = k.FirstChild.Value; //Верхняя граница
                                }

                                if ("ND".Equals(k.Name) && k.FirstChild != null)
                                {
                                    ND = k.FirstChild.Value; //НД
                                }
                                if ("Description".Equals(k.Name) && k.FirstChild != null)
                                {
                                    Description = k.FirstChild.Value; //Примечание
                                    textBox1.Text = Description;
                                }
                                if ("DateTime".Equals(k.Name) && k.FirstChild != null)
                                {
                                    DateTime = k.FirstChild.Value; //Дата
                                    dateTimePicker1.Text = DateTime;
                                }
                                if ("DateTime1_1".Equals(k.Name) && k.FirstChild != null)
                                {
                                    DateTime2_1 = k.FirstChild.Value; //Дата
                                    label6.Text = DateTime2_1;
                                }
                                if ("DateTime1_1_1".Equals(k.Name) && k.FirstChild != null)
                                {
                                    DateTime2_2_1 = k.FirstChild.Value; //Дата
                                    numericUpDown1.Value = Convert.ToInt32(DateTime2_2_1);
                                }
                                if ("Pogreshnost".Equals(k.Name) && k.FirstChild != null)
                                {
                                    Pogreshnost2 = k.FirstChild.Value; //Дата
                                    textBox3.Text = Pogreshnost2;
                                }
                                if ("Ispolnitel".Equals(k.Name) && k.FirstChild != null)
                                {
                                    Ispolnitel = k.FirstChild.Value; //Исполнитель
                                }
                                /*  Stolbec = new string[this.Table1.Rows.Count - 1, this.Table1.Columns.Count];
                                      for (int i = 0; i < this.Table1.Rows.Count - 1; i++)
                                  {
                                      for (int j = 0; j < this.Table1.Columns.Count; j++)
                                      {
                                          if ("Stolbec".Equals(k.Name) && k.FirstChild != null)
                                          {
                                              Stolbec[i, j] = k.FirstChild.Value;
                                              Table1.Rows[i].Cells[j].Value = Stolbec[i, j];
                                          }
                                      }
                                  }*/
                                if ("TypeIzmer".Equals(k.Name) && k.FirstChild != null)
                                {
                                    TimeIzmer1 = k.FirstChild.Value;

                                }
                                if ("TypeYravn".Equals(k.Name) && k.FirstChild != null)
                                {
                                    TypeYravn1 = k.FirstChild.Value; //Исполнитель
                                    if (TypeYravn1 == "Линейное через 0" || TypeYravn1 == "Линейное")
                                    {
                                        //  MessageBox.Show("Линейное");
                                        Table1.Columns.Add("X*X", "Конц*Конц");
                                        Table1.Columns.Add("X*Y", "Асред*Конц");
                                        /*  Table1.Columns["X*X"].Width = 50;
                                          Table1.Columns["X*Y"].Width = 50;
                                          Table1.Columns["X*X*X"].Width = 50;
                                          Table1.Columns["X*X*X*X"].Width = 50;
                                          Table1.Columns["X*X*Y"].Width = 50;*/
                                    }
                                    else
                                    {
                                        // MessageBox.Show("Квадратичное");
                                        Table1.Columns.Add("X*X", "Конц* Конц");
                                        Table1.Columns.Add("X*Y", "Асред* Конц");
                                        Table1.Columns.Add("X*X*X", "Асред ^3");
                                        Table1.Columns.Add("X*X*X*X", "Асред ^4");
                                        Table1.Columns.Add("X*X*Y", "Асред ^2*Конц");
                                        /*    Table1.Columns["X*X"].Width = 50;
                                            Table1.Columns["X*Y"].Width = 50;
                                            Table1.Columns["X*X*X"].Width = 50;
                                            Table1.Columns["X*X*X*X"].Width = 50;
                                            Table1.Columns["X*X*Y"].Width = 50;*/
                                    }

                                }
                                if ("CountSeriyal".Equals(k.Name) && k.FirstChild != null)
                                {
                                    CountSeriya2 = k.FirstChild.Value; //Количество столбцов
                                    CountSeriya = CountSeriya2;
                                    while (true)
                                    {
                                        int i = Table1.Columns.Count - 1;//С какого столбца начать
                                        if (Table1.Columns[i].Name == "Asred")
                                            break;
                                        Table1.Columns.RemoveAt(i);
                                    }
                                    for (int i = 1; i <= Convert.ToInt32(CountSeriya2); i++)
                                    {

                                        DataGridViewTextBoxColumn firstColumn2 =
                                        new DataGridViewTextBoxColumn();
                                        firstColumn2.HeaderText = "A; Сер" + i;
                                        firstColumn2.Name = "A;Ser (" + i;
                                        Table1.Columns.Add(firstColumn2);
                                    }
                                    for (int i = 1; i <= Convert.ToInt32(CountSeriya2); i++)
                                    {
                                        Table1.Columns["A;Ser (" + i].Width = 50;
                                    }
                                }
                                if ("CountInSeriyal".Equals(k.Name) && k.FirstChild != null)
                                {
                                    CountInSeriya2 = k.FirstChild.Value; //Количество строк
                                    CountInSeriya = CountInSeriya2;

                                    for (int i = 0; i < Convert.ToInt32(CountInSeriya2); i++)
                                    {
                                        Table1.Rows.Add();
                                    }


                                }

                            }

                        }
                    }
                    if (TypeYravn1 == "Линейное через 0" || TypeYravn1 == "Линейное")
                    {
                        StolbecCol = 5 + Convert.ToInt32(CountSeriya2);
                    }
                    else
                    {
                        StolbecCol = 8 + Convert.ToInt32(CountSeriya2);
                    }
                    Stolbec = new string[Convert.ToInt32(CountInSeriya2), StolbecCol];
                    int stroka = 0;

                    // Читаем в цикле вложенные значения Stroka

                    for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                    {

                        // Обрабатываем в цикле только Stroka
                        if ("Stroka".Equals(d.Name))
                        {
                            int stolbec = 0;
                            //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                            for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                            {


                                if ("Stolbec".Equals(k.Name) && k.FirstChild != null)
                                {

                                    Stolbec[stroka, stolbec] = k.FirstChild.Value;


                                    stolbec++;
                                }

                            }
                            stroka++;
                        }

                    }
                }
            }
            if (ComPort == true)
            {
                SW2();
            }
            else
            {
                //   GWNew.Text = wavelength1;
            }

        }
        public string F1;
        public string F2;

        public void openFile2(ref string filepath2)
        {
            WLREMOVE2();
            WLREMOVESTR2();
            filepath2 = openFileDialog1.FileName;
            параметрыToolStripMenuItem.Enabled = true;
            button10.Enabled = true;

            XmlDocument xDoc = new XmlDocument();

            xDoc.Load(@filepath2);

            XmlNodeList nodes = xDoc.ChildNodes;
            // Обходим значения
            foreach (XmlNode n in nodes)
            { // Обрабатываем в цикле только Data_Izmerenie
                if ("Data_Izmerenie".Equals(n.Name))
                {
                    // Читаем в цикле вложенные значения Izmerenie
                    for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                    {
                        // Обрабатываем в цикле только Izmerenie
                        if ("Izmerenie".Equals(d.Name))
                        {
                            //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                            for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                            {

                                if ("WidthCuvet".Equals(k.Name) && k.FirstChild != null)
                                {
                                    WidthCuvette = k.FirstChild.Value; //Ширина кюветы
                                    int index = Opt_dlin_cuvet.FindString(WidthCuvette);
                                    //  MessageBox.Show(index.ToString());
                                    Opt_dlin_cuvet.SelectedIndex = index;
                                }

                                if ("Description".Equals(k.Name) && k.FirstChild != null)
                                {
                                    Description = k.FirstChild.Value; //Примечание
                                    textBox8.Text = Description;
                                }
                                if ("DateTime".Equals(k.Name) && k.FirstChild != null)
                                {
                                    DateTime = k.FirstChild.Value; //Дата
                                    dateTimePicker2.Text = DateTime;
                                }

                                if ("Pogreshnost".Equals(k.Name) && k.FirstChild != null)
                                {
                                    Pogreshnost2 = k.FirstChild.Value; //Дата
                                    textBox7.Text = Pogreshnost2;
                                }
                                if ("F1".Equals(k.Name) && k.FirstChild != null)
                                {
                                    F1 = k.FirstChild.Value; //F1
                                    F1Text.Text = F1;
                                }
                                if ("F2".Equals(k.Name) && k.FirstChild != null)
                                {
                                    F2 = k.FirstChild.Value; //F1
                                    F2Text.Text = F2;
                                }

                                if ("CountSeriyal".Equals(k.Name) && k.FirstChild != null)
                                {
                                    CountSeriya2 = k.FirstChild.Value; //Количество столбцов
                                    while (true)
                                    {
                                        int i = Table2.Columns.Count - 1;//С какого столбца начать
                                        if (Table2.Columns[i].Name == "Obrazec")
                                            break;
                                        Table2.Columns.RemoveAt(i);
                                    }

                                    for (int i = 1; i <= Convert.ToInt32(CountSeriya2); i++)
                                    {

                                        DataGridViewTextBoxColumn firstColumn2 =
                                        new DataGridViewTextBoxColumn();
                                        firstColumn2.HeaderText = "A; Сер." + i;
                                        firstColumn2.Name = "A;Ser" + i;
                                        firstColumn2.ValueType = Type.GetType("System.Double");
                                        Table2.Columns.Add(firstColumn2);
                                        DataGridViewTextBoxColumn firstColumn3 =
                                        new DataGridViewTextBoxColumn();
                                        firstColumn3.HeaderText = "C, " + edconctr + "; Сер." + i;
                                        firstColumn3.Name = "C,edconctr;Ser." + i;
                                        firstColumn3.ValueType = Type.GetType("System.Double");
                                        Table2.Columns.Add(firstColumn3);

                                    }
                                    DataGridViewTextBoxColumn firstColumn4 =
                                    new DataGridViewTextBoxColumn();
                                    firstColumn4.HeaderText = "Cср, " + edconctr;
                                    firstColumn4.Name = "Ccr";
                                    firstColumn4.ValueType = Type.GetType("System.Double");
                                    Table2.Columns.Add(firstColumn4);

                                    DataGridViewTextBoxColumn firstColumn5 =
                                    new DataGridViewTextBoxColumn();
                                    firstColumn5.HeaderText = "d, %";
                                    firstColumn5.Name = "d%";
                                    firstColumn5.ValueType = Type.GetType("System.Double");
                                    Table2.Columns.Add(firstColumn5);
                                }
                                if ("CountInSeriyal".Equals(k.Name) && k.FirstChild != null)
                                {
                                    CountInSeriya2 = k.FirstChild.Value; //Количество строк
                                    NoCaSer1 = Convert.ToInt32(CountInSeriya2);

                                    for (int i = 0; i < Convert.ToInt32(CountInSeriya2); i++)
                                    {
                                        Table2.Rows.Add();
                                    }


                                }

                            }

                        }
                    }
                    StolbecCol = 2 + Convert.ToInt32(CountSeriya2) + Convert.ToInt32(CountSeriya2) + 2;

                    Stolbec = new string[Convert.ToInt32(CountInSeriya2), StolbecCol];
                    int stroka = 0;

                    // Читаем в цикле вложенные значения Stroka

                    for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                    {

                        // Обрабатываем в цикле только Stroka
                        if ("Stroka".Equals(d.Name))
                        {
                            int stolbec = 0;
                            //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                            for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                            {


                                if ("Stolbec".Equals(k.Name) && k.FirstChild != null)
                                {

                                    Stolbec[stroka, stolbec] = k.FirstChild.Value;


                                    stolbec++;
                                }

                            }
                            stroka++;
                        }

                    }
                }
            }


        }

        public void TableWrite()
        {

            int StolbecCol = 3 + Convert.ToInt32(CountSeriya2);


            for (int i = 0; i < Convert.ToInt32(CountInSeriya2); i++)
            {
                for (int j = 0; j < StolbecCol; j++)
                {

                    Table1.Rows[i].Cells[j].Value = Stolbec[i, j];
                }

            }
            NoCaIzm = Convert.ToInt32(CountSeriya2);
            if (TimeIzmer1 == "A (C) - градуировочное уравнение (стандарт)")
            {
                radioButton4.Checked = true;
            }
            else
            {
                radioButton5.Checked = true;
            }
            if (TypeYravn1 == "Линейное через 0")
            {
                radioButton1.Checked = true;
                lineinaya0();
            }
            else
            {
                if (TypeYravn1 == "Линейное")
                {
                    radioButton2.Checked = true;
                    lineinaya();
                }
                else
                {
                    radioButton3.Checked = true;
                    kvadratichnaya();
                }
            }
        }
        public void TableWrite2()
        {

            int StolbecCol = 2 + Convert.ToInt32(CountSeriya2) + Convert.ToInt32(CountSeriya2) + 2;


            for (int i = 0; i < Convert.ToInt32(CountInSeriya2); i++)
            {
                for (int j = 0; j < StolbecCol; j++)
                {

                    Table2.Rows[i].Cells[j].Value = Stolbec[i, j];
                }

            }
            NoCaIzm1 = Convert.ToInt32(CountSeriya2);
        }

        private void новыйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                this.textBox4.Text = string.Format("{0:0.0000}", 0);
                this.textBox5.Text = string.Format("{0:0.0000}", 0);
                this.textBox6.Text = string.Format("{0:0.0000}", 0);
                //    Izmerenie1 = true;
                while (true)
                {
                    int i = Table1.Columns.Count - 1;//С какого столбца начать
                    if (Table1.Columns.Count == 3 + NoCaIzm)
                        break;
                    Table1.Columns.RemoveAt(i);
                }
                ParametrsGrad _ParametrsGrad = new ParametrsGrad(this);

                _ParametrsGrad.button1.Click += (ParametrsGrad, eSlave) =>
                {
                    Veshestvo1 = _ParametrsGrad.Veshestvo.Text;
                    wavelength1 = _ParametrsGrad.WL_grad.Text;
                    WidthCuvette = _ParametrsGrad.Opt_dlin_cuvet.Text;
                    BottomLine = _ParametrsGrad.Down.Text;
                    TopLine = _ParametrsGrad.Up.Text;
                    ND = _ParametrsGrad.ND.Text;
                    Description = _ParametrsGrad.Description.Text;
                    DateTime = _ParametrsGrad.dateTimePicker1.Text;
                    Ispolnitel = _ParametrsGrad.Ispolnitel.Text;
                    CountSeriya = _ParametrsGrad.numericUpDown3.Text;
                    CountInSeriya = _ParametrsGrad.numericUpDown4.Text;
                    textBox3.Text = _ParametrsGrad.textBox4.Text;
                    Days = Convert.ToInt32(_ParametrsGrad.numericUpDown1.Value);
                    label6.Text = dateTimePicker1.Value.AddDays(Days).ToString("dd.MM.yyyy");
                    edconctr = _ParametrsGrad.Ed.Text;
                    if (_ParametrsGrad.radioButton7.Checked == true)
                    {
                        this.textBox4.Text = string.Format("{0:0.0000}", _ParametrsGrad.k0Text.Text);
                        this.textBox5.Text = string.Format("{0:0.0000}", _ParametrsGrad.k1Text.Text);
                        this.textBox6.Text = string.Format("{0:0.0000}", _ParametrsGrad.k2Text.Text);


                    }
                    else
                    {
                        if (_ParametrsGrad.radioButton6.Checked == true)
                        {
                            this.textBox4.Text = string.Format("{0:0.0000}", 0);
                            this.textBox5.Text = string.Format("{0:0.0000}", 0);
                            this.textBox6.Text = string.Format("{0:0.0000}", 0);
                        }
                    }
                    if (_ParametrsGrad.radioButton6.Checked == true)
                    {
                        SposobZadan = "По СО";
                    }
                    else
                    {
                        SposobZadan = "Ввод коэффициентов";
                    }
                    if (_ParametrsGrad.radioButton4.Checked == true)
                    {
                        Zavisimoct = "A(C)";
                    }
                    else
                    {
                        Zavisimoct = "C(A)";
                    }
                    if (_ParametrsGrad.radioButton1.Checked == true)
                    {
                        aproksim = "Линейная через 0";
                    }
                    else
                    {
                        if (_ParametrsGrad.radioButton2.Checked == true)
                        {
                            aproksim = "Линейная";
                        }
                        else
                        {
                            aproksim = "Квадратичная";
                        }
                    }
                };

                _ParametrsGrad.ShowDialog();
                dateTimePicker1.Text = DateTime;
                tabControl2.SelectTab(tabPage3);
                chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;

            }
            else
            {
                New _New = new New(this);
                _New.ShowDialog();
            }


        }
        public bool IzmerenieOpen = false;


        public double[] El;

        private void Table2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (Table1.CurrentCell.ReadOnly == true)
            {
                MessageBox.Show("Запись запрещена!");
            }
            double CCR = 0.0;
            bool doNotWrite = false;
            bool doNotWrite1 = false;
            double maxEl;
            double minEl;
            double serValue = 0;
            int cellnull = 0;
            for (int i = 2; i < Table2.Rows[Table2.CurrentCell.RowIndex].Cells.Count - 1; i++)
            {
                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells[i].Value == null)
                {
                    El = new double[NoCaIzm1 + 1];

                    doNotWrite = true;


                    // El = new double[NoCaIzm1 + 1];
                    for (int j = 0; j < Table2.Rows.Count - 1; j++)
                    {
                        // El = new double[NoCaIzm1 + 1];
                        double SredValue = 0;
                        for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                        {
                            if (Table2.Rows[j].Cells["A;Ser" + i1].Value == null)
                            {
                                cellnull++;
                            }
                            else
                            {
                                if (aproksim == "Линейная через 0")
                                {
                                    serValue = Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) / Convert.ToDouble(textBox5.Text);
                                }
                                if (aproksim == "Линейная")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / Convert.ToDouble(textBox5.Text);
                                }
                                if (aproksim == "Квадратичная")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / (Convert.ToDouble(textBox5.Text) + Convert.ToDouble(textBox6.Text));
                                }
                                double CValue1 = Convert.ToDouble(F1Text.Text);
                                double CValue2 = Convert.ToDouble(F2Text.Text);

                                Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);
                                SredValue += Convert.ToDouble(Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value.ToString());

                                CCR = SredValue / NoCaIzm1;
                                if (Convert.ToDouble(textBox7.Text) >= 1)
                                {
                                    Table2.Rows[j].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.00}", (CCR / Convert.ToDouble(textBox7.Text)));
                                }
                                else Table2.Rows[j].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                                //Table2.Rows[j].Cells["d%"].Value = El.Max();
                                //  El[i1] = Convert.ToDouble(Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value.ToString());
                            }
                            //El = new double[NoCaIzm1 + 1];
                            if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value == null)
                            {
                                cellnull++;
                            }
                            else
                            {
                                El[i1] = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                            }
                        }

                        Array.Sort(El);
                        maxEl = El[El.Length - 1];
                        minEl = El[1];
                        double a = ((maxEl - minEl) * 100) / Convert.ToDouble(CCR);
                        double b = a;
                        // b = b * 10;


                        if (minEl == 0)
                        {
                            Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = 0.0000;
                        }
                        else
                        {
                            Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = string.Format("{0:0.00}", b);

                        }

                    }
                }
            }

            if (!doNotWrite)
            {
                El = new double[Convert.ToInt32(CountSeriya2) + 1];
                for (int j = 0; j < Table2.Rows.Count - 1; j++)
                {
                    double SredValue = 0;
                    for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                    {
                        if (Table2.Rows[j].Cells["A;Ser" + i1].Value == null)
                        {
                            cellnull++;
                        }
                        else
                        {
                            if (aproksim == "Линейная через 0")
                            {
                                serValue = Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) / Convert.ToDouble(textBox5.Text);
                            }
                            if (aproksim == "Линейная")
                            {
                                serValue = ((Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / Convert.ToDouble(textBox5.Text);
                            }
                            if (aproksim == "Квадратичная")
                            {
                                serValue = ((Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / (Convert.ToDouble(textBox5.Text) + Convert.ToDouble(textBox6.Text));
                            }
                            double CValue1 = Convert.ToDouble(F1Text.Text);
                            double CValue2 = Convert.ToDouble(F2Text.Text);

                            Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);
                            SredValue += Convert.ToDouble(Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value.ToString());

                            CCR = SredValue / NoCaIzm1;
                            if (Convert.ToDouble(textBox7.Text) >= 1)
                            {
                                Table2.Rows[j].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.0000}", (CCR / Convert.ToDouble(textBox7.Text)));
                            }
                            else Table2.Rows[j].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                            //Table2.Rows[j].Cells["d%"].Value = El.Max();
                            El[i1] = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                        }

                    }

                    Array.Sort(El);
                    maxEl = El[El.Length - 1];
                    minEl = El[1];
                    double a = ((maxEl - minEl) * 100) / Convert.ToDouble(CCR);
                    double b = a;


                    if (minEl == 0)
                    {
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = 0.0000;
                    }
                    else
                    {
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = string.Format("{0:0.00}", b);

                    }

                }
            }
        }

        private void Table2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        public bool nonPort = false;
        private void button2_Click(object sender, EventArgs e)
        {
            // MessageBox.Show(Izmerenie1.ToString());
            SettingPort _SettingPort = new SettingPort(this);
            if (nonPort == true)
            {
                _SettingPort.ShowDialog();
            }
            else
            {
                _SettingPort.Dispose();
            }
            //_SettingPort.Close();
            if (nonPort == true && Izmerenie1 != false)
            {
                newPort = new SerialPort();

                try
                {
                    // настройки порта (Communication interface)
                    newPort.PortName = portsName;
                    newPort.BaudRate = 19200;
                    newPort.DataBits = 8;
                    newPort.Parity = System.IO.Ports.Parity.None;
                    newPort.StopBits = System.IO.Ports.StopBits.One;
                    // Установка таймаутов чтения/записи (read/write timeouts)
                    newPort.ReadTimeout = 20000;
                    newPort.WriteTimeout = 20000;
                    //    newPort.DataReceived += new SerialDataReceivedEventHandler(newPort_DataReceived);
                    newPort.RtsEnable = false;
                    newPort.DtrEnable = true;
                    newPort.Open();
                    // MessageBox.Show("ПОРТ ОТКРЫТ " + newPort.PortName);


                    newPort.DiscardInBuffer();
                    newPort.DiscardOutBuffer();
                }
                catch (Exception)
                {
                    MessageBox.Show("Порт не был выбран!");
                    return;

                }

                //char[] OpenPribor = { Convert.ToChar('C'), Convert.ToChar('O'), Convert.ToChar('\r') };
                //newPort.Write(OpenPribor, 0, OpenPribor.Length);


                Analis__Activated();
                CO();
                // GW();
                RD();
                //GW();
                //SA();
                Izmerenie1 = false;
                ComPodkl = true;
                SAGE(ref countSA, ref GE5_1_0);
                this.подключитьToolStripMenuItem.Enabled = false;
                this.настройкаПортаToolStripMenuItem.Enabled = true;
                this.информацияToolStripMenuItem.Enabled = true;
                this.калибровкаToolStripMenuItem.Enabled = true;
                this.темновойТокToolStripMenuItem.Enabled = true;
                this.измеритьToolStripMenuItem.Enabled = true;
                this.одноволновоеИзмерениеToolStripMenuItem.Enabled = true;
                this.одноволновоеИзмерениеToolStripMenuItem.Enabled = true;
                this.калибровкаДляОдноволновогоАнализаToolStripMenuItem.Enabled = true;
                button1.Enabled = true;
                button2.Enabled = false;
                button13.Enabled = true;
                button12.Enabled = true;
                button14.Enabled = true;
                ComPort = true;
                wavelength1 = GWNew.Text;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ComPort == true)
            {
                char[] ClosePribor = { Convert.ToChar('Q'), Convert.ToChar('U'), Convert.ToChar('\r') };
                newPort.Write("QU\r");
                Thread.Sleep(500);
                //  byte[] buffer1 = new byte[byteRecieved1];
                string indata = newPort.ReadExisting();

                bool indata_bool = true;
                while (indata_bool == true)
                {
                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();
                    }
                }

                GWNew.Text = null;
                GEText.Text = null;
                GAText.Text = null;
                OptichPlot.Text = null;
                Izmerenie1 = true;

                this.подключитьToolStripMenuItem.Enabled = true;
                button2.Enabled = true;
                button13.Enabled = false;
                button12.Enabled = false;
                button14.Enabled = false;
                this.настройкаПортаToolStripMenuItem.Enabled = false;
                this.информацияToolStripMenuItem.Enabled = false;
                this.калибровкаToolStripMenuItem.Enabled = false;
                this.темновойТокToolStripMenuItem.Enabled = false;
                this.измеритьToolStripMenuItem.Enabled = false;
                this.одноволновоеИзмерениеToolStripMenuItem.Enabled = false;
                this.одноволновоеИзмерениеToolStripMenuItem.Enabled = false;
                this.калибровкаДляОдноволновогоАнализаToolStripMenuItem.Enabled = false;
                button1.Enabled = false;
                newPort.Close();


            }
            ComPort = false;
            /* if (ComPort == false)
              {
                  Close();
              }*/
        }

        private void button14_Click(object sender, EventArgs e)
        {
            double sum = 0.0;
            int startIndexCell = 2;
            int endIndexCell = startIndexCell + NoCaIzm;
            int rowIndex = Table1.CurrentRow.Index;

            bool doNotWrite = false;
            if (tabControl2.SelectedIndex == 0)
            {
                string SWAnalis = WL_grad1;
                string GE5Izmer = "";
                string GE5_1_1 = "";
                /*newPort.Write("SW " + SWAnalis + "\r");
                int SWAnalisByteRecieved1 = newPort.ReadBufferSize;
                Thread.Sleep(100);
                byte[] SWAnalisBuffer1 = new byte[SWAnalisByteRecieved1];
                newPort.Read(SWAnalisBuffer1, 0, SWAnalisByteRecieved1);*/

                newPort.Write("SA " + countSA + "\r");
                string indata = newPort.ReadExisting();
                string indata_0;
                bool indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();

                    }
                }

                newPort.Write("GE 1\r");

                GE5Izmer = "";
                int GEbyteRecieved4_1 = newPort.ReadBufferSize;
                byte[] GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

                indata = newPort.ReadExisting();

                indata_0 = "";
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {

                        indata = newPort.ReadExisting();
                        indata_0 += indata;
                    }
                }
                Regex regex = new Regex(@"\W");
                GE5Izmer = regex.Replace(indata_0, "");



                GEText.Text = GE5Izmer;
                // MessageBox.Show(GE5Izmer);
                double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
                double OptPlot1 = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5Izmer));
                double OptPlot1_1 = OptPlot1 - Math.Truncate(OptPlot1);
                if (Table1.CurrentCell.ReadOnly != true)
                {
                    Table1.CurrentCell.Value = string.Format("{0:0.0000}", OptPlot1_1);
                }
                else
                {
                    MessageBox.Show("Запись запрещена!");
                }
                GAText.Text = string.Format("{0:0.00}", Aser);
                for (int j = 0; j < Table1.Rows.Count; j++)
                {
                    {
                        for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                        {
                            if (Table1.Rows[j].Cells[i].Value == null)
                            {
                                doNotWrite = true;

                                for (int l = startIndexCell + NoCaIzm; l <= endIndexCell; ++l)
                                {
                                    if (Table1.Rows[rowIndex].Cells[l].Value == null)
                                    {
                                        cellnull++;
                                    }
                                }
                            }


                        }
                    }
                }



                if (!doNotWrite)
                {
                    if (Table1.Rows.Count == 1)
                    {
                        radioButton1.Enabled = true;
                        radioButton4.Enabled = true;
                        radioButton5.Enabled = true;
                        radioButton3.Enabled = false;
                        radioButton2.Enabled = false;
                    }
                    if (Table1.Rows.Count == 2)
                    {
                        radioButton1.Enabled = true;
                        radioButton2.Enabled = true;
                        radioButton4.Enabled = true;
                        radioButton5.Enabled = true;
                        radioButton3.Enabled = false;
                    }
                    if (Table1.Rows.Count >= 3)
                    {
                        radioButton1.Enabled = true;
                        radioButton2.Enabled = true;
                        radioButton3.Enabled = true;
                        radioButton4.Enabled = true;
                        radioButton5.Enabled = true;
                    }
                    sum = 0.0;
                    while (true)
                    {
                        int i = Table1.Columns.Count - 1;//С какого столбца начать
                        if (Table1.Columns.Count == 3 + Convert.ToInt32(CountSeriya2))
                            break;
                        Table1.Columns.RemoveAt(i);
                    }

                    for (int l = 3; l <= Table1.Rows[Table1.CurrentCell.RowIndex].Cells.Count - 1; l++)
                    {
                        if (Table1.Rows[rowIndex].Cells[l].Value == null)
                        {
                            cellnull++;
                        }

                        else
                        {
                            for (int j = 0; j < Table1.Rows.Count; j++)
                            {

                                for (int i1 = startIndexCell + 1; i1 <= endIndexCell; ++i1)
                                {
                                    sum += Convert.ToDouble(Table1.Rows[j].Cells[i1].Value);
                                    Asred1 = sum / NoCaIzm;
                                    // MessageBox.Show(Convert.ToString(Asred1));
                                    Table1.Rows[j].Cells["Asred"].Value = string.Format("{0:0.0000}", Asred1);

                                }
                                sum = 0.0;
                            }
                        }
                        Izmerenie1 = true;
                    }
                    for (int m = 0; m < Table1.Rows.Count - 1; m++)
                    {
                        for (int ml = 0; ml < Table1.Rows[m].Cells.Count; ml++)
                        {
                            if (Table1.Rows[m].Cells[ml].Value == null)
                            { doNotWrite = true; }
                        }
                    }
                    if (!doNotWrite)
                    {
                        while (true)
                        {
                            int ml = Table1.Columns.Count - 1;//С какого столбца начать
                            if (Table1.Columns.Count == 3 + NoCaIzm)
                                break;
                            Table1.Columns.RemoveAt(ml);
                        }
                        functionAsred();

                    }

                }
            }
            else
            {
                ///Измерение
                ///
                double CCR = 0.0;
                int rowIndex2 = Table2.CurrentRow.Index;
                bool doNotWrite1 = false;
                double maxEl;
                double minEl;
                double serValue = 0;
                int cellnull = 0;
                El = new double[NoCaIzm1 + 1];
                string GE5Izmer = "";
                string GE5_1_1 = "";
                newPort.Write("SA " + countSA + "\r");
                string indata = newPort.ReadExisting();
                string indata_0;
                bool indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();

                    }
                }

                newPort.Write("GE 1\r");

                GE5Izmer = "";
                int GEbyteRecieved4_1 = newPort.ReadBufferSize;
                byte[] GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);

                indata = newPort.ReadExisting();

                indata_0 = "";
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {

                        indata = newPort.ReadExisting();
                        indata_0 += indata;
                    }
                }
                Regex regex = new Regex(@"\W");
                GE5Izmer = regex.Replace(indata_0, "");
                GEText.Text = GE5Izmer;
                // MessageBox.Show(GE5Izmer);
                double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
                double OptPlot1 = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5Izmer));
                double OptPlot1_1 = OptPlot1 - Math.Truncate(OptPlot1);
                if (Table2.CurrentCell.ReadOnly != true)
                {
                    Table2.CurrentCell.Value = string.Format("{0:0.0000}", OptPlot1_1);
                }
                else
                {
                    MessageBox.Show("Запись запрещена!");
                }
                GAText.Text = string.Format("{0:0.00}", Aser);
                double SredValue = 0;
                for (int i = 2; i < Table2.Rows[Table2.CurrentCell.RowIndex].Cells.Count - 1; i++)
                {
                    if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells[i].Value == null)
                    {
                        El = new double[NoCaIzm1 + 1];

                        doNotWrite = true;


                        // El = new double[NoCaIzm1 + 1];
                        for (int j = 0; j < Table2.Rows.Count - 1; j++)
                        {
                            SredValue = 0;
                            for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                            {
                                if (Table2.Rows[j].Cells["A;Ser" + i1].Value == null)
                                {
                                    cellnull++;
                                }
                                else
                                {
                                    if (aproksim == "Линейная через 0")
                                    {
                                        serValue = Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) / Convert.ToDouble(textBox5.Text);
                                    }
                                    if (aproksim == "Линейная")
                                    {
                                        serValue = ((Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / Convert.ToDouble(textBox5.Text);
                                    }
                                    if (aproksim == "Квадратичная")
                                    {
                                        serValue = ((Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / (Convert.ToDouble(textBox5.Text) + Convert.ToDouble(textBox6.Text));
                                    }
                                    double CValue1 = Convert.ToDouble(F1Text.Text);
                                    double CValue2 = Convert.ToDouble(F2Text.Text);

                                    Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);
                                    SredValue += Convert.ToDouble(Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value.ToString());

                                    CCR = SredValue / NoCaIzm1;
                                    if (Convert.ToDouble(textBox7.Text) >= 1)
                                    {
                                        Table2.Rows[j].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.0000}", (CCR / Convert.ToDouble(textBox7.Text)));
                                    }
                                    else Table2.Rows[j].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                                    //Table2.Rows[j].Cells["d%"].Value = El.Max();
                                    // El[i1] = Convert.ToDouble(Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value.ToString());
                                }

                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value == null)
                                {
                                    cellnull++;
                                }
                                else
                                {
                                    El[i1] = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                                }
                            }

                            Array.Sort(El);
                            maxEl = El[El.Length - 1];
                            minEl = El[1];
                            double a = ((maxEl - minEl) * 100) / Convert.ToDouble(CCR);
                            double b = a;


                            if (minEl == 0)
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = 0.0000;
                            }
                            else
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = string.Format("{0:0.00}", b);

                            }

                        }
                    }
                }

                if (!doNotWrite)
                {
                    El = new double[Convert.ToInt32(CountSeriya2) + 1];
                    for (int j = 0; j < Table2.Rows.Count - 1; j++)
                    {
                        SredValue = 0;
                        for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                        {
                            if (Table2.Rows[j].Cells["A;Ser" + i1].Value == null)
                            {
                                cellnull++;
                            }
                            else
                            {
                                if (aproksim == "Линейная через 0")
                                {
                                    serValue = Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) / Convert.ToDouble(textBox5.Text);
                                }
                                if (aproksim == "Линейная")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / Convert.ToDouble(textBox5.Text);
                                }
                                if (aproksim == "Квадратичная")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(textBox4.Text))) / (Convert.ToDouble(textBox5.Text) + Convert.ToDouble(textBox6.Text));
                                }
                                double CValue1 = Convert.ToDouble(F1Text.Text);
                                double CValue2 = Convert.ToDouble(F2Text.Text);

                                Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);
                                SredValue += Convert.ToDouble(Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value.ToString());

                                CCR = SredValue / NoCaIzm1;
                                if (Convert.ToDouble(textBox7.Text) >= 1)
                                {
                                    Table2.Rows[j].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.0000}", (CCR / Convert.ToDouble(textBox7.Text)));
                                }
                                else Table2.Rows[j].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                                //Table2.Rows[j].Cells["d%"].Value = El.Max();
                                El[i1] = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                            }

                        }

                        Array.Sort(El);
                        maxEl = El[El.Length - 1];
                        minEl = El[1];
                        double a = ((maxEl - minEl) * 100) / Convert.ToDouble(CCR);
                        double b = a;


                        if (minEl == 0)
                        {
                            Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = 0.000;
                        }
                        else
                        {
                            Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = string.Format("{0:0.00}", b);

                        }

                    }
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            SAGE(ref countSA, ref GE5_1_0);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Izmerenie1 = true;
            if (tabControl2.SelectedIndex == 0)
            {
                openFileDialog1.InitialDirectory = "C";
                openFileDialog1.Title = "Open File";
                openFileDialog1.FileName = "";
                openFileDialog1.Filter = "QS2 файл|*.qs2";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    try
                    {
                        // получаем выбранный файл
                        openFile(ref filepath);
                    }
                    catch (Exception t) { MessageBox.Show("exeption" + t.Message); }


                    TableWrite();
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton3.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                }
                tabPage4.Parent = tabControl2;
            }
            else
            {
                openFileDialog1.InitialDirectory = "C";
                openFileDialog1.Title = "Open File";
                openFileDialog1.FileName = "";
                openFileDialog1.Filter = "QA2 файл|*.qa2";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // получаем выбранный файл
                        openFile2(ref filepath2);
                    }
                    catch (Exception t) { MessageBox.Show("exeption" + t.Message); }


                    TableWrite2();
                }

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                bool doNotWrite = false;
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {

                    for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                    {
                        if (Table1.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;
                            break;

                        }
                    }
                }
                if (doNotWrite == true)
                {
                    MessageBox.Show("Не вся поля таблицы были заполнены!");
                }
                else
                {
                    SaveAs1();
                    tabPage4.Parent = tabControl2;

                }
            }
            else
            {
                bool doNotWrite = false;
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {

                    for (int i = 2; i < Table1.Rows[j].Cells.Count; i++)
                    {
                        if (Table1.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;
                            break;

                        }
                    }
                }
                if (doNotWrite == true)
                {
                    MessageBox.Show("Не вся поля таблицы были заполнены!");
                }
                else
                {
                    SaveAs2();
                }
            }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                bool doNotWrite = false;
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {

                    for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                    {
                        if (Table1.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;
                            break;

                        }
                    }
                }
                if (doNotWrite == true)
                {
                    MessageBox.Show("Не вся поля таблицы были заполнены!");
                }
                else
                {
                    ExportToPDF1();
                }
            }
            else
            {
                bool doNotWrite = false;
                for (int j = 0; j < Table2.Rows.Count - 1; j++)
                {

                    for (int i = 2; i < Table2.Rows[j].Cells.Count; i++)
                    {
                        if (Table2.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;
                            break;

                        }
                    }
                }
                if (doNotWrite == true)
                {
                    MessageBox.Show("Не вся поля таблицы были заполнены!");
                }
                else
                {
                    ExportToPDF();
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            int GDByteRecieved1 = newPort.ReadBufferSize;
            Thread.Sleep(100);
            byte[] GDBuffer1 = new byte[GDByteRecieved1];
            newPort.Read(GDBuffer1, 0, GDByteRecieved1);
            RD();
            SAGE(ref countSA, ref GE5_1_0);
        }

        public void lineinaya0()
        {

            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            k0 = 0; k1 = 0; k2 = 0;
            circle = 0;
            XY = 0;
            SUMMY2 = 0;
            textBox4.Text = string.Format("{0:0.0000}", 0);
            textBox5.Text = string.Format("{0:0.0000}", 0);
            textBox6.Text = string.Format("{0:0.0000}", 0);

            if (radioButton4.Checked == true)
            {
                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");
                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                }
                catch
                {
                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;


                }
                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                k0 = 0; k1 = 0; k2 = 0;
                circle = 0;
                XY = 0;
                SUMMY2 = 0;

                for (int i = 0; i < Table1.Rows.Count - 1; i++)
                {
                    double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                    double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                    XY += x * y;
                    SUMMY2 += y * y;

                    Table1.Rows[i].Cells["X*X"].Value = y * y;
                    Table1.Rows[i].Cells["X*Y"].Value = x * y;
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "";
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "";
                    Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);

                }
            }
            else
            {
                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");
                    Table1.Columns.Add("X*X", "Асред* Асред");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                }
                catch
                {
                    Table1.Columns.Add("X*X", "Асред* Асред");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;

                }
                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                k0 = 0; k1 = 0; k2 = 0;
                circle = 0;
                XY = 0;
                SUMMY2 = 0;
                for (int i = 0; i < Table1.Rows.Count - 1; i++)
                {
                    double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                    double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                    XY += x * y;
                    SUMMY2 += y * y;
                    Table1.Rows[i].Cells["X*X"].Value = y * y;
                    Table1.Rows[i].Cells["X*Y"].Value = x * y;
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "";
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "";
                    Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);
                }
            }

            k1 = XY / SUMMY2;

            textBox4.Text = string.Format("{0:0.0000}", 0);
            textBox5.Text = string.Format("{0:0.0000}", k1);
            textBox4.Text = string.Format("{0:0.0000}", 0);

            if (radioButton4.Checked == true)
            {
                if (textBox5.Text != string.Format("{0:0.0000}", 0))
                {
                    int Table1_Asred = 0;
                    label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C";
                    for (int i = 0; i < Table1.Rows.Count - 1; i++)
                    {
                        double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                        double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                        chart1.Series[0].Points.AddXY(x1, y1);
                        chart1.Series[0].ChartType = SeriesChartType.Point;
                        chart1.ChartAreas[0].AxisY.Crossing = 0;
                        chart1.ChartAreas[0].AxisX.Crossing = 0;


                        //     chart1.Series[0].Points.Item.Label = Convert.ToString(circle);
                        // chart1.Series[0].IsValueShownAsLabel = true;
                        //chart1.Series[0].IsXValueIndexed = true;
                        // circle++;
                        // double x2 = 0.1 * i;
                        // double y2 = x2 / k1;
                        double x2 = 0;
                        if (Table1_Asred == 0)
                        {
                            x2 = 0;
                        }
                        else
                        {
                            x2 = x1;
                        }
                        Table1_Asred++;
                        double y2 = x2 * k1;
                        chart1.Series[1].Points.AddXY(x2, y2);
                        chart1.Series[1].ChartType = SeriesChartType.Line;
                        chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                        chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                        chart1.ChartAreas[0].AxisX.Minimum = 0;
                        //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                        chart1.ChartAreas[0].AxisY.Minimum = 0;
                        //   chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2;
                        //    chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Concetr"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Concetr"].Value)), 2);
                        //   chart1.ChartAreas[0].AxisX.Interval = 5;
                    }
                }

            }
            else
            {
                if (textBox5.Text != string.Format("{0:0.0000}", 0))
                {
                    label14.Text = "C(A) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*A";
                    for (int i = 0; i < Table1.Rows.Count - 1; i++)
                    {
                        double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                        double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                        chart1.Series[0].Points.AddXY(x1, y1);
                        chart1.Series[0].ChartType = SeriesChartType.Point;
                        chart1.ChartAreas[0].AxisY.Crossing = 0;
                        chart1.ChartAreas[0].AxisX.Crossing = 0;


                        //     chart1.Series[0].Points.Item.Label = Convert.ToString(circle);
                        // chart1.Series[0].IsValueShownAsLabel = true;
                        //chart1.Series[0].IsXValueIndexed = true;
                        // circle++;
                        //double y2 = 0.5 * i;
                        //double x2 = y2 / k1;

                        double x2 = x1;
                        double y2 = x1 * k1;
                        chart1.Series[1].Points.AddXY(x2, y2);
                        chart1.Series[1].ChartType = SeriesChartType.Line;
                        chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                        chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                        chart1.ChartAreas[0].AxisX.Minimum = 0;
                        //   chart1.ChartAreas[0].AxisX.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2), 2);
                        chart1.ChartAreas[0].AxisY.Minimum = 0;
                        // chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                        // chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value)), 2);
                        //    chart1.ChartAreas[0].AxisX.Interval = 5;
                    }
                }
            }

        }

        public void lineinaya()
        {

            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            k0 = 0; k1 = 0; k2 = 0;
            SUM0 = 0; SUM1 = 0;
            SUMMX = 0; SUMMY = 0; XY = 0; SUMMY2 = 0;
            textBox4.Text = string.Format("{0:0.0000}", 0);
            textBox5.Text = string.Format("{0:0.0000}", 0);
            textBox4.Text = string.Format("{0:0.0000}", 0);
            if (radioButton4.Checked == true)
            {

                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");
                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                }



                catch
                {
                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;

                }

                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                k0 = 0; k1 = 0; k2 = 0;
                SUM0 = 0; SUM1 = 0;
                SUMMX = 0; SUMMY = 0; XY = 0; SUMMY2 = 0;
                for (int i = 0; i < Table1.Rows.Count - 1; i++)
                {
                    double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                    double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                    double y0 = Convert.ToDouble(Table1.Rows[0].Cells["Concetr"].Value);
                    double y1 = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value);
                    double x0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
                    double x1 = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value);
                    SUMMX += x; SUMMY += y;
                    XY += x * y;
                    SUMMY2 += y * y;
                    Table1.Rows[i].Cells["X*X"].Value = y * y;
                    Table1.Rows[i].Cells["X*Y"].Value = x * y;
                    Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMX);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMY);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);
                }
            }
            else
            {
                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");
                    Table1.Columns.Add("X*X", "Асред* Асред");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;

                }
                catch
                {
                    Table1.Columns.Add("X*X", "Асред* Асред");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;

                }
                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                k0 = 0; k1 = 0; k2 = 0;
                SUM0 = 0; SUM1 = 0;
                SUMMX = 0; SUMMY = 0; XY = 0; SUMMY2 = 0;
                for (int i = 0; i < Table1.Rows.Count - 1; i++)
                {
                    double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                    double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                    double x0 = Convert.ToDouble(Table1.Rows[0].Cells["Concetr"].Value);
                    double x1 = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value);
                    double y0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
                    double y1 = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value);
                    SUMMX += x; SUMMY += y;
                    XY += x * y;
                    SUMMY2 += y * y;
                    Table1.Rows[i].Cells["X*X"].Value = y * y;
                    Table1.Rows[i].Cells["X*Y"].Value = x * y;
                    Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMX);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMY);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);
                }
            }

            k0 = (SUMMY2 * SUMMX - SUMMY * XY) / ((Table1.Rows.Count - 1) * SUMMY2 - SUMMY * SUMMY);
            k1 = ((Table1.Rows.Count - 1) * XY - SUMMY * SUMMX) / ((Table1.Rows.Count - 1) * SUMMY2 - SUMMY * SUMMY);
            textBox4.Text = string.Format("{0:0.0000}", k0);
            textBox5.Text = string.Format("{0:0.0000}", k1);
            textBox6.Text = string.Format("{0:0.0000}", 0);
            if (radioButton4.Checked == true)
            {
                if (textBox4.Text != string.Format("{0:0.0000}", 0) && textBox5.Text != string.Format("{0:0.0000}", 0))
                {
                    label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C " + k0.ToString("+ 0.0000 ;- 0.0000 ");
                    for (int i = 0; i < Table1.Rows.Count - 1; i++)
                    {
                        circle = 1;
                        double x1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                        double y1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                        // chart1.ChartAreas[0].AxisY.Crossing = k0;
                        chart1.Series[0].Points.AddXY(y1_1, x1_1);
                        chart1.Series[0].ChartType = SeriesChartType.Point;
                        chart1.ChartAreas[0].AxisY.Crossing = 0;
                        chart1.ChartAreas[0].AxisX.Crossing = 0;
                        //  double x2 = 0.1 * i;
                        //double y2 = (x2 - k0) / k1;
                        double y2 = y1_1;
                        double x2 = y1_1 * k1 + k0;
                        chart1.Series[1].Points.AddXY(y2, x2);
                        chart1.Series[1].ChartType = SeriesChartType.Line;
                        chart1.ChartAreas[0].AxisX.Minimum = 0;
                        chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                        chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                        //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(string.Format("{0:0.0000}", Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2);
                        chart1.ChartAreas[0].AxisY.Minimum = 0;
                        //   chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(string.Format("{0:0.0000}", Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2);
                        //   chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Concetr"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Concetr"].Value)), 2);

                        //       chart1.ChartAreas[0].AxisX.Interval = 5;
                    }
                }
            }
            else
            {
                if (textBox4.Text != string.Format("{0:0.0000}", 0) && textBox5.Text != string.Format("{0:0.0000}", 0))
                {
                    label14.Text = "C(A) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*A " + k0.ToString("+ 0.0000 ;- 0.0000 ");
                    for (int i = 0; i < Table1.Rows.Count - 1; i++)
                    {
                        circle = 1;
                        double x1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                        double y1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                        // chart1.ChartAreas[0].AxisY.Crossing = k0;
                        chart1.Series[0].Points.AddXY(x1_1, y1_1);
                        chart1.Series[0].ChartType = SeriesChartType.Point;
                        chart1.ChartAreas[0].AxisY.Crossing = 0;
                        chart1.ChartAreas[0].AxisX.Crossing = 0;
                        // double y2 = 0.5 * i;
                        //     double x2 = (y2 - k0) / k1;
                        //  double y2 = k1 * x1_1 + k0;
                        double x2 = x1_1;
                        double y2 = x1_1 * k1 + k0;
                        chart1.Series[1].Points.AddXY(x2, y2);
                        chart1.Series[1].ChartType = SeriesChartType.Line;
                        chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                        chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                        chart1.ChartAreas[0].AxisX.Minimum = 0;
                        //  chart1.ChartAreas[0].AxisX.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2), 2);
                        chart1.ChartAreas[0].AxisY.Minimum = 0;
                        //      chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                        //   chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value)), 2);
                        //     chart1.ChartAreas[0].AxisX.Interval = 5;
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Izmerenie1 = true;
            if (tabControl2.SelectedIndex == 0)
            {
                this.textBox4.Text = string.Format("{0:0.0000}", 0);
                this.textBox5.Text = string.Format("{0:0.0000}", 0);
                this.textBox6.Text = string.Format("{0:0.0000}", 0);
                //    Izmerenie1 = true;
                while (true)
                {
                    int i = Table1.Columns.Count - 1;//С какого столбца начать
                    if (Table1.Columns.Count == 3 + NoCaIzm)
                        break;
                    Table1.Columns.RemoveAt(i);
                }
                ParametrsGrad _ParametrsGrad = new ParametrsGrad(this);

                _ParametrsGrad.button1.Click += (ParametrsGrad, eSlave) =>
                {
                    Veshestvo1 = _ParametrsGrad.Veshestvo.Text;
                    wavelength1 = _ParametrsGrad.WL_grad.Text;
                    WidthCuvette = _ParametrsGrad.Opt_dlin_cuvet.Text;
                    BottomLine = _ParametrsGrad.Down.Text;
                    TopLine = _ParametrsGrad.Up.Text;
                    ND = _ParametrsGrad.ND.Text;
                    Description = _ParametrsGrad.Description.Text;
                    DateTime = _ParametrsGrad.dateTimePicker1.Text;
                    Ispolnitel = _ParametrsGrad.Ispolnitel.Text;
                    CountSeriya = _ParametrsGrad.numericUpDown3.Text;
                    CountInSeriya = _ParametrsGrad.numericUpDown4.Text;
                    textBox3.Text = _ParametrsGrad.textBox4.Text;
                    Days = Convert.ToInt32(_ParametrsGrad.numericUpDown1.Value);
                    label6.Text = dateTimePicker1.Value.AddDays(Days).ToString("dd.MM.yyyy");
                    edconctr = _ParametrsGrad.Ed.Text;
                    if (_ParametrsGrad.radioButton7.Checked == true)
                    {
                        this.textBox4.Text = string.Format("{0:0.0000}", _ParametrsGrad.k0Text.Text);
                        this.textBox5.Text = string.Format("{0:0.0000}", _ParametrsGrad.k1Text.Text);
                        this.textBox6.Text = string.Format("{0:0.0000}", _ParametrsGrad.k2Text.Text);


                    }
                    else
                    {
                        if (_ParametrsGrad.radioButton6.Checked == true)
                        {
                            this.textBox4.Text = string.Format("{0:0.0000}", 0);
                            this.textBox5.Text = string.Format("{0:0.0000}", 0);
                            this.textBox6.Text = string.Format("{0:0.0000}", 0);
                        }
                    }
                    if (_ParametrsGrad.radioButton6.Checked == true)
                    {
                        SposobZadan = "По СО";
                    }
                    else
                    {
                        SposobZadan = "Ввод коэффициентов";
                    }
                    if (_ParametrsGrad.radioButton4.Checked == true)
                    {
                        Zavisimoct = "A(C)";
                    }
                    else
                    {
                        Zavisimoct = "C(A)";
                    }
                    if (_ParametrsGrad.radioButton1.Checked == true)
                    {
                        aproksim = "Линейная через 0";
                    }
                    else
                    {
                        if (_ParametrsGrad.radioButton2.Checked == true)
                        {
                            aproksim = "Линейная";
                        }
                        else
                        {
                            aproksim = "Квадратичная";
                        }
                    }
                };

                _ParametrsGrad.ShowDialog();
                dateTimePicker1.Text = DateTime;
                tabControl2.SelectTab(tabPage3);
                chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;

            }
            else
            {
                New _New = new New(this);
                _New.ShowDialog();
            }
        }

        private void информацияToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void Table2_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

        }

        private void печать1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintDocument def = new PrintDocument();
            def.PrintPage += new PrintPageEventHandler(PRD);
            def.DocumentName = "Document1";
            def.PrinterSettings = printDialog1.PrinterSettings;
            PrintPreviewDialog dlg = new PrintPreviewDialog();
            dlg.WindowState = FormWindowState.Maximized;
            dlg.Document = def;
            dlg.ShowDialog();
        }
        void PRD(object sender, PrintPageEventArgs e)
        {

            filepath = openFileDialog1.FileName;
            // printPreviewDialog1.Document = @"filepath";
            Graphics g = e.Graphics;
            g.DrawString(@filepath, Font, new SolidBrush(System.Drawing.Color.Black), 0, 0);

        }
        Bitmap bitmap;
        private void printPreviewDialog1_Load(object sender, EventArgs e)
        {

        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            ExportToPDF1();
        }

        public void kvadratichnaya()
        {
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            double x2 = 0; double x3 = 0; double x4 = 0; double xy = 0; double SUMX = 0;
            double SUMY = 0; double x2y = 0;
            double Opred; double OpredA; double OpredB; double OpredC;
            k0 = 0; k1 = 0; k2 = 0;
            if (radioButton4.Checked == true)
            {
                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");
                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Асред* Конц");
                    Table1.Columns.Add("X*X*X", "Асред ^3");
                    Table1.Columns.Add("X*X*X*X", "Асред ^4");
                    Table1.Columns.Add("X*X*Y", "Асред ^2*Конц");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X*X"].Width = 50;
                    Table1.Columns["X*X*X*X"].Width = 50;
                    Table1.Columns["X*X*Y"].Width = 50;
                }
                catch
                {
                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Асред* Конц");
                    Table1.Columns.Add("X*X*X", "Асред ^3");
                    Table1.Columns.Add("X*X*X*X", "Асред ^4");
                    Table1.Columns.Add("X*X*Y", "Асред ^2*Конц");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X*X"].Width = 50;
                    Table1.Columns["X*X*X*X"].Width = 50;
                    Table1.Columns["X*X*Y"].Width = 50;
                }
                for (int i = 0; i < Table1.Rows.Count - 1; i++)
                {

                    double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                    double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                    x2 += x * x;
                    x3 += x * x * x;
                    x4 += x * x * x * x;
                    xy += x * y;
                    SUMX += x;
                    SUMY += y;
                    x2y += x * x * y;
                    Table1.Rows[i].Cells["X*X"].Value = string.Format("{0:0.0000}", x * x);
                    Table1.Rows[i].Cells["X*Y"].Value = string.Format("{0:0.0000}", x * y);
                    Table1.Rows[i].Cells["X*X*X"].Value = string.Format("{0:0.0000}", x * x * x);
                    Table1.Rows[i].Cells["X*X*X*X"].Value = string.Format("{0:0.0000}", x * x * x * x);
                    Table1.Rows[i].Cells["X*X*Y"].Value = string.Format("{0:0.0000}", x * x * y);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMX);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMY);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(x2);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X"].Value = "СУММА = " + Convert.ToString(x3);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X*X"].Value = "СУММА = " + Convert.ToString(x4);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*Y"].Value = "СУММА = " + Convert.ToString(x2y);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(xy);

                }
            }
            else
            {
                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");
                    Table1.Columns.Add("X*X", "Асред ^2");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns.Add("X*X*X", "Асред ^3");
                    Table1.Columns.Add("X*X*X*X", "Асред ^4");
                    Table1.Columns.Add("X*X*Y", "Асред ^2*Конц");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X*X"].Width = 50;
                    Table1.Columns["X*X*X*X"].Width = 50;
                    Table1.Columns["X*X*Y"].Width = 50;
                }
                catch
                {
                    Table1.Columns.Add("X*X", "Асред ^2");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns.Add("X*X*X", "Асред ^3");
                    Table1.Columns.Add("X*X*X*X", "Асред ^4");
                    Table1.Columns.Add("X*X*Y", "Асред ^2*Конц");

                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X*X"].Width = 50;
                    Table1.Columns["X*X*X*X"].Width = 50;
                    Table1.Columns["X*X*Y"].Width = 50;
                }
                for (int i = 0; i < Table1.Rows.Count - 1; i++)
                {
                    double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                    double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                    x2 += x * x;
                    x3 += x * x * x;
                    x4 += x * x * x * x;
                    xy += x * y;
                    SUMX += x;
                    SUMY += y;
                    x2y += x * x * y;
                    Table1.Rows[i].Cells["X*X"].Value = string.Format("{0:0.0000}", x * x);
                    Table1.Rows[i].Cells["X*X*X"].Value = string.Format("{0:0.0000}", x * x * x);
                    Table1.Rows[i].Cells["X*X*X*X"].Value = string.Format("{0:0.0000}", x * x * x * x);
                    Table1.Rows[i].Cells["X*X*Y"].Value = string.Format("{0:0.0000}", x * x * y);
                    Table1.Rows[i].Cells["X*Y"].Value = string.Format("{0:0.0000}", x * y);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMX);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMY);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(x2);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X"].Value = "СУММА = " + Convert.ToString(x3);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X*X"].Value = "СУММА = " + Convert.ToString(x4);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*Y"].Value = "СУММА = " + Convert.ToString(x2y);
                    Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(xy);
                }
            }
            Opred = x2 * x2 * x2 + SUMX * SUMX * x4 + (Table1.Rows.Count - 1) * x3 * x3 - (Table1.Rows.Count - 1) * x2 * x4 - x2 * SUMX * x3 - SUMX * x3 * x2;
            OpredA = SUMY * x2 * x2 + SUMX * SUMX * x2y + (Table1.Rows.Count - 1) * xy * x3 - (Table1.Rows.Count - 1) * x2 * x2y - SUMY * SUMX * x3 - SUMX * xy * x2;
            OpredB = x2 * xy * x2 + SUMY * SUMX * x4 + (Table1.Rows.Count - 1) * x3 * x2y - (Table1.Rows.Count - 1) * xy * x4 - x2 * SUMX * x2y - SUMY * x3 * x2;
            OpredC = x2 * x2 * x2y + SUMX * xy * x4 + SUMY * x3 * x3 - SUMY * x2 * x4 - x2 * xy * x3 - SUMX * x3 * x2y;

            k2 = OpredA / Opred;
            k1 = OpredB / Opred;
            k0 = OpredC / Opred;
            textBox4.Text = string.Format("{0:0.0000}", k0);
            textBox5.Text = string.Format("{0:0.0000}", k1);
            textBox6.Text = string.Format("{0:0.0000}", k2);
            if (radioButton4.Checked == true)
            {
                if (textBox4.Text != string.Format("{0:0.0000}", 0) && textBox5.Text != string.Format("{0:0.0000}", 0) && textBox6.Text != string.Format("{0:0.0000}", 0))
                {
                    label14.Text = "A(C) = " + k0.ToString("0.0000 ;- 0.0000 ") + k1.ToString("+ 0.0000 ;- 0.0000 ") + "*C " + k2.ToString("+ 0.0000 ;- 0.0000 ") + "*C^2";
                    for (int i = 0; i < Table1.Rows.Count - 1; i++)
                    {
                        double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                        double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                        chart1.Series[0].Points.AddXY(x, y);
                        chart1.Series[0].ChartType = SeriesChartType.Point;
                        chart1.ChartAreas[0].AxisY.Crossing = 0;
                        chart1.ChartAreas[0].AxisX.Crossing = 0;

                        // double x2_1 = 0.3 * i;
                        double x2_1 = x;
                        double y2_1 = k0 + k1 * x2_1 + k2 * x2_1 * x2_1;

                        chart1.Series[1].Points.AddXY(x2_1, y2_1);
                        chart1.Series[1].ChartType = SeriesChartType.Line;
                        chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                        chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                        chart1.ChartAreas[0].AxisX.Minimum = 0;
                        //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + x2_1;
                        chart1.ChartAreas[0].AxisY.Minimum = 0;
                        //   chart1.ChartAreas[0].AxisY.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + y2_1), 2);
                        //chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Concetr"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Concetr"].Value)), 2);
                    }
                }
            }
            else
            {
                if (textBox4.Text != string.Format("{0:0.0000}", 0) && textBox5.Text != string.Format("{0:0.0000}", 0) && textBox6.Text != string.Format("{0:0.0000}", 0))
                {
                    label14.Text = "C(A) = " + k0.ToString("0.0000 ;- 0.0000 ") + k1.ToString("+ 0.0000;- 0.0000") + "*A " + k2.ToString("+ 0.0000;- 0.0000") + "*A^2";
                    for (int i = 0; i < Table1.Rows.Count - 1; i++)
                    {
                        double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                        double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                        chart1.Series[0].Points.AddXY(x, y);
                        chart1.Series[0].ChartType = SeriesChartType.Point;
                        chart1.ChartAreas[0].AxisY.Crossing = 0;
                        chart1.ChartAreas[0].AxisX.Crossing = 0;
                        double x2_1 = x;
                        double y2_1 = k0 + k1 * x2_1 + k2 * x2_1 * x2_1;

                        chart1.Series[1].Points.AddXY(x2_1, y2_1);
                        chart1.Series[1].ChartType = SeriesChartType.Line;
                        chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                        chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                        chart1.ChartAreas[0].AxisX.Minimum = 0;
                        //  chart1.ChartAreas[0].AxisX.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2_1), 2);
                        chart1.ChartAreas[0].AxisY.Minimum = 0;
                        //   chart1.ChartAreas[0].AxisY.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2_1), 2);
                        //   chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value)+ (Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value))), 2);
                        //  chart1.ChartAreas[0].AxisX.Interval = 5;
                    }
                }
            }
        }
        int countOpen = 0;
        private void tabControl2_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (countOpen == 0)
            {
                параметрыToolStripMenuItem.Enabled = false;
                button10.Enabled = false;
            }
            countOpen++;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                //    this.textBox4.Text = string.Format("{0:0.0000}", 0);
                //   this.textBox5.Text = string.Format("{0:0.0000}", 0);
                //   this.textBox6.Text = string.Format("{0:0.0000}", 0);
                //    Izmerenie1 = true;
                while (true)
                {
                    int i = Table1.Columns.Count - 1;//С какого столбца начать
                    if (Table1.Columns.Count == 3 + NoCaIzm)
                        break;
                    Table1.Columns.RemoveAt(i);
                }


                NewGraduirovka _NewGraduirovka = new NewGraduirovka(this);

                _NewGraduirovka.button1.Click += (NewGraduirovka, eSlave) =>
                {
                    Veshestvo1 = _NewGraduirovka.Veshestvo.Text;
                    wavelength1 = _NewGraduirovka.WL_grad.Text;
                    WidthCuvette = _NewGraduirovka.Opt_dlin_cuvet.Text;
                    BottomLine = _NewGraduirovka.Down.Text;
                    TopLine = _NewGraduirovka.Up.Text;
                    ND = _NewGraduirovka.ND.Text;
                    Description = _NewGraduirovka.Description.Text;
                    DateTime = _NewGraduirovka.dateTimePicker1.Text;
                    Ispolnitel = _NewGraduirovka.Ispolnitel.Text;
                    textBox3.Text = _NewGraduirovka.textBox4.Text;
                    CountSeriya = _NewGraduirovka.numericUpDown3.Text;
                    CountInSeriya = _NewGraduirovka.numericUpDown4.Text;
                    edconctr = _NewGraduirovka.Ed.Text;
                    Days = Convert.ToInt32(_NewGraduirovka.numericUpDown1.Value);
                    label6.Text = dateTimePicker1.Value.AddDays(Days).ToString("dd.MM.yyyy");
                    if (_NewGraduirovka.radioButton7.Checked == true)
                    {
                        this.textBox4.Text = string.Format("{0:0.0000}", _NewGraduirovka.k0Text.Text);
                        this.textBox5.Text = string.Format("{0:0.0000}", _NewGraduirovka.k1Text.Text);
                        this.textBox6.Text = string.Format("{0:0.0000}", _NewGraduirovka.k2Text.Text);


                    }
                    else
                    {
                        if (_NewGraduirovka.radioButton6.Checked == true)
                        {
                            //  this.textBox4.Text = string.Format("{0:0.0000}", 0);
                            //   this.textBox5.Text = string.Format("{0:0.0000}", 0);
                            ///   this.textBox6.Text = string.Format("{0:0.0000}", 0);
                        }
                    }

                };
                _NewGraduirovka.ShowDialog();
                dateTimePicker1.Text = DateTime;
                tabControl2.SelectTab(tabPage3);
                chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                //       Table1.Columns.Add("X*X", "Concetr*Concetr");
                //Table1.Columns.Add("X*Y", "Asred*Concetr");
                //CalibrovkaGrad();
                // MessageBox.Show(Convert.ToString(Days));

            }
            else
            {

                Parametr1 _Parametr1 = new Parametr1(this);
                _Parametr1.ShowDialog();
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (ComPort == true)
            {
                Izmerenie1 = true;
                char[] ClosePribor = { Convert.ToChar('Q'), Convert.ToChar('U'), Convert.ToChar('\r') };
                newPort.Write("QU\r");
                Thread.Sleep(500);
                //  byte[] buffer1 = new byte[byteRecieved1];
                string indata = newPort.ReadExisting();

                bool indata_bool = true;
                while (indata_bool == true)
                {
                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();
                    }
                }

                newPort.Close();
                ComPort = false;
                SWF.Application.Exit();
            }
            else
            {
                SWF.Application.Exit();
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                bool doNotWrite = false;
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {

                    for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                    {
                        if (Table1.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;
                            break;

                        }
                    }
                }
                if (doNotWrite == true)
                {
                    MessageBox.Show("Не вся поля таблицы были заполнены!");
                }
                else
                {
                    ExportToExcel();
                }
            }
            else
            {
                bool doNotWrite = false;
                for (int j = 0; j < Table2.Rows.Count - 1; j++)
                {

                    for (int i = 2; i < Table2.Rows[j].Cells.Count; i++)
                    {
                        if (Table2.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;
                            break;

                        }
                    }
                }
                if (doNotWrite == true)
                {
                    MessageBox.Show("Не вся поля таблицы были заполнены!");
                }
                else
                {
                    ExportToExcel2();
                }
            }
        }

        private void Table1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (Table1.CurrentCell.ColumnIndex >= 3 && Table1.CurrentCell.ReadOnly != true)

            {
                TextBox Table1 = (TextBox)e.Control;
                Table1.KeyPress += new KeyPressEventHandler(tb_KeyPress);
            }
            else
            {
                TextBox Table1 = (TextBox)e.Control;
                Table1.KeyPress -= tb_KeyPress;
            }
            if (Table1.CurrentCell.ReadOnly == true)
            {
                MessageBox.Show("Редактирование ячейки запрещено!");
            }
        }
        private void tb_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("Только цифры!");
            }
        }

        private void Table2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (Table2.CurrentCell.ColumnIndex >= 2 && Table2.CurrentCell.ReadOnly != true)

            {
                TextBox Table2 = (TextBox)e.Control;
                Table2.KeyPress += new KeyPressEventHandler(tb_KeyPress1);
            }
            else
            {
                TextBox Table2 = (TextBox)e.Control;
                Table2.KeyPress -= tb_KeyPress1;
            }
            if (Table2.CurrentCell.ReadOnly == true)
            {
                MessageBox.Show("Редактирование ячейки запрещено!");
            }
        }
        private void tb_KeyPress1(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("Только цифры!");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();

        }

        private void printDocument1_PrintPage_1(object sender, PrintPageEventArgs e)
        {
            e.Graphics.DrawString("Расчет линейного градуировочного графика\n", new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 15, FontStyle.Bold), Brushes.Black, new Point(170, 60));
            e.Graphics.DrawString("Вещество: " + Veshestvo1, new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(50, 90));
            e.Graphics.DrawString("Длина волны: " + wavelength1, new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(50, 120));
            e.Graphics.DrawString("Ширина кюветы: " + WidthCuvette, new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(50, 150));
            e.Graphics.DrawString("Нижняя нраница обнаружения: " + BottomLine, new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(50, 180));
            e.Graphics.DrawString("Верхняя нраница обнаружения: " + TopLine, new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(50, 210));
            e.Graphics.DrawString("НД: " + ND, new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(500, 90));
            e.Graphics.DrawString("Примечание: " + Description, new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(50, 240));
            e.Graphics.DrawString("Таблица исходных данных\n\n", new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(50, 270));


            int height = Table1.Height;
            Table1.Height = (Table1.RowCount * Table1.RowTemplate.Height) + 90;
            Bitmap Table1bmp = new Bitmap(Table1.Width, Table1.Height);
            Table1.DrawToBitmap(Table1bmp, new System.Drawing.Rectangle(0, 0, Table1.Width, Table1.Height));
            Table1.Height = height;
            e.Graphics.DrawImage(Table1bmp, 50, 300);

            e.Graphics.DrawString("Градуировочное уравнение: " + label14.Text, new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(50, 430));
            Bitmap Chart1bmp = new Bitmap(chart1.Size.Width + 10, chart1.Size.Height + 70);
            chart1.DrawToBitmap(Chart1bmp, chart1.Bounds);
            e.Graphics.DrawImage(Chart1bmp, 50, 460);

            e.Graphics.DrawString("Дата: " + DateTime, new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(50, 930));
            e.Graphics.DrawString("Исполнитель: " + Ispolnitel, new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(50, 960));



        }

        private void printPreviewDialog1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Dispose();
        }

        private void Table1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                bool doNotWrite = false;
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {

                    for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                    {
                        if (Table1.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;
                            break;

                        }
                    }
                }
                if (doNotWrite == true)
                {
                    MessageBox.Show("Не вся поля таблицы были заполнены!");
                }
                else
                {
                    printPreviewDialog1.Document = printDocument1;
                    printPreviewDialog1.ShowDialog();
                }
            }
        }
        int cordY = 0;

        private void printDocument1_PrintPage_2(object sender, PrintPageEventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                e.Graphics.DrawString("Расчет линейного градуировочного графика\n\n", new System.Drawing.Font("Times New Roman", 20, FontStyle.Bold), Brushes.Black, 100, 50);
                e.Graphics.DrawString("Вещесво:", new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 25, 110);
                e.Graphics.DrawString(Veshestvo1, new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular), Brushes.Black, 115, 110);
                e.Graphics.DrawString("Длина волны:", new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 25, 130);
                e.Graphics.DrawString(wavelength1, new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular), Brushes.Black, 155, 130);
                e.Graphics.DrawString("Ширина кюветы:", new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 25, 150);
                e.Graphics.DrawString(WidthCuvette, new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular), Brushes.Black, 190, 150);
                e.Graphics.DrawString("Нижняя граница обнаружения:", new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 25, 170);
                e.Graphics.DrawString(BottomLine, new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular), Brushes.Black, 315, 170);
                e.Graphics.DrawString("Верхняя граница обнаружения:", new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 25, 190);
                e.Graphics.DrawString(TopLine, new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular), Brushes.Black, 315, 190);
                e.Graphics.DrawString("НД:", new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 450, 110);
                e.Graphics.DrawString(ND, new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular), Brushes.Black, 500, 110);
                e.Graphics.DrawString("Примечание:", new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 25, 210);
                e.Graphics.DrawString(Description, new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular), Brushes.Black, 155, 210);
                e.Graphics.DrawString("Таблица исходных данных", new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 25, 230);


                
                if (NoCaIzm <= 3)
                {
                    Table1PrintViewer1(sender, e);
                }
                else
                {
                    if (NoCaIzm > 3 && NoCaIzm <= 7)
                    {
                        Table1PrintViewer2(sender, e);
                    }
                    else
                    {
                        Table1PrintViewer3(sender, e);
                    }
                }

           

                e.Graphics.DrawString("Градуировочное уравнение:", new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 25, cordY+30);
                e.Graphics.DrawString(label14.Text, new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular), Brushes.Black, 285, cordY+30);
                int height = chart1.Height;
                Bitmap bmp = new Bitmap(chart1.Width, chart1.Height);
                chart1.DrawToBitmap(bmp, new System.Drawing.Rectangle(0,0, chart1.Width, chart1.Height));
                e.Graphics.DrawImage(bmp, 25, cordY+60);
                cordY = cordY + chart1.Height + 60;
                e.Graphics.DrawString("Дата:", new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 25, cordY);
                e.Graphics.DrawString(DateTime, new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular), Brushes.Black, 80, cordY);
                e.Graphics.DrawString("Исполнитель:", new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 25, cordY + 30);
                e.Graphics.DrawString(Ispolnitel, new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular), Brushes.Black, 160, cordY + 30);

                //  Paragraph Ispolnitel2 = new Paragraph("Исполнитель: " + Ispolnitel, font);

            }


        }
        ///Если меньше или равно 3
        public void Table1PrintViewer1(object sender, PrintPageEventArgs e)
        {
            int height = 270;
            int width = 25;
            Pen p = new Pen(Brushes.Black, 2.5f);

            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[0].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[0].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[1].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[1].Width;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[2].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[2].Width + 5;
            for (int i = 3; i <= Table1.Columns.Count - NoCaIzm; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[i].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[i].Width + 10;
            }
            for (int i = Table1.Columns.Count - NoCaIzm + 1; i < Table1.Columns.Count; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[i].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[i].Width + 10;
                // height = height + Table1.Rows[i].Height;
            }
            height = height + Table1.Rows[0].Height * 2;
            width = 25;
            int height1 = height;
            int width1_1 = width;
           
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                    e.Graphics.DrawString(Table1.Rows[j].Cells[0].Value.ToString(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                    // width = width + Table1.Columns[0].Width;
                    height += Table1.Rows[j].Height;
                }
                height = height1;
                width = width + Table1.Columns[0].Width + 5;
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                    e.Graphics.DrawString(Table1.Rows[j].Cells[1].Value.ToString(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                    // width = width + Table1.Columns[1].Width;
                    height += Table1.Rows[j].Height;
                }
                height = height1;
                width = width + Table1.Columns[1].Width;
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[2].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[2].Value.ToString(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                    }
                    // width = width + Table1.Columns[2].Width;
                    height += Table1.Rows[j].Height;
                }
                height = height1;
                width = width + Table1.Columns[2].Width + 5;
                int width1 = width;
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {
                    for (int i = 3; i <= Table1.Columns.Count - NoCaIzm; i++)
                    {
                        e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                        e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                        if (Table1.Rows[j].Cells[i].Value != null)
                        {
                            e.Graphics.DrawString(Table1.Rows[j].Cells[i].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                        }
                        else
                        {
                            e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                        }
                        width = width + Table1.Columns[i].Width + 10;
                        width1_1 = width;
                    }
                    height += Table1.Rows[j].Height;
                    width = width1;
                }

                height = height1;
                width1 = width1_1;
                width = width1;
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {
                    for (int i = Table1.Columns.Count - NoCaIzm + 1; i < Table1.Columns.Count; i++)
                    {
                        e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                        e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                        if (Table1.Rows[j].Cells[i].Value != null)
                        {
                            e.Graphics.DrawString(Table1.Rows[j].Cells[i].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                        }
                        else
                        {
                            e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                        }
                        width = width + Table1.Columns[i].Width + 10;
                    }
                    cordY = height;
                    height += Table1.Rows[j].Height;
                    width = width1;
                }
            

        }
        ///Если больше 3 и меньше или равно 7
        public void Table1PrintViewer2(object sender, PrintPageEventArgs e)
        {
            int height = 270;
            int width = 25;
            Pen p = new Pen(Brushes.Black, 2.5f);

            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[0].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[0].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[1].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[1].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[1].Width;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[2].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[2].Width + 5;
            int k = 3;
            for (int i = 0; i < NoCaIzm; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[k].Width + 10;
                k++;
            }
            height = height + Table1.Rows[0].Height * 2;
            /* Формируем значения */
            width = 25;
            int height1 = height;
            int width1_1 = width;


            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawString(Table1.Rows[j].Cells[0].Value.ToString(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                // width = width + Table1.Columns[0].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[0].Width + 5;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                e.Graphics.DrawString(Table1.Rows[j].Cells[1].Value.ToString(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                // width = width + Table1.Columns[1].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[1].Width;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                if (Table1.Rows[j].Cells[2].Value != null)
                {
                    e.Graphics.DrawString(Table1.Rows[j].Cells[2].Value.ToString(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                }
                else
                {
                    e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                }
                // width = width + Table1.Columns[2].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[2].Width + 5;
            int width1 = width;
            k = 3;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = 0; i < NoCaIzm; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[k].Width + 10;
                    k++;
                    //width1_1 = width;
                }
                height += Table1.Rows[j].Height;
                width = width1;
                k = 3;
            }
            /*Cancel*/
            height = height + 10;
            width = 25;
            k = NoCaIzm + 3;
            for (int i = 0; i < Table1.Columns.Count - NoCaIzm - 3; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[k].Width + 10;
                k++;
                // height = height + Table1.Rows[i].Height;
            }

            /*Формируем вторую часть значений*/

            height = height + Table1.Rows[0].Height * 2;
            width1 = 25;
            width = 25;
            k = NoCaIzm + 3;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = 0; i < Table1.Columns.Count - NoCaIzm - 3; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[k].Width + 10;
                    k++;
                }
                cordY = height;
                height += Table1.Rows[j].Height;
                width = width1;
                k = NoCaIzm + 3;
            }
            /*Cancel*/           

        }


        /*Если больше 7*/
        public void Table1PrintViewer3(object sender, PrintPageEventArgs e)
        {
            int height = 270;
            int width = 25;
            Pen p = new Pen(Brushes.Black, 2.5f);

            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[0].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[0].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[1].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[1].Height * 2));
            e.Graphics.DrawString(Table1.Columns[1].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[1].Width;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[2].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[2].Width + 5;
            int k = 3;
            for (int i = 0; i < 7; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[k].Width + 10;
                k++;
            }
            
            
            height = height + Table1.Rows[0].Height * 2;
            /* Формируем значения */
            width = 25;
            int height1 = height;
            int width1_1 = width;


            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawString(Table1.Rows[j].Cells[0].Value.ToString(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                // width = width + Table1.Columns[0].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[0].Width + 5;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                e.Graphics.DrawString(Table1.Rows[j].Cells[1].Value.ToString(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                // width = width + Table1.Columns[1].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[1].Width;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                if (Table1.Rows[j].Cells[2].Value != null)
                {
                    e.Graphics.DrawString(Table1.Rows[j].Cells[2].Value.ToString(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                }
                else
                {
                    e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                }
                // width = width + Table1.Columns[2].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[2].Width + 5;
            int width1 = width;
            k = 3;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = 0; i < 7; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[k].Width + 10;
                    k++;
                    //width1_1 = width;
                }
                height += Table1.Rows[j].Height;
                width = width1;
                k = 3;
            }
            /*Cancel*/

            height = height + 10;
            width = 25;
            k = 10;
            //k = 11;
            for (int i = 0; i < NoCaIzm - 7; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[k].Width + 10;
                k++;
            }
           
            for (int i = 0; i < Table1.Columns.Count - NoCaIzm - 3; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[k].Width + 10;
                k++;
                // height = height + Table1.Rows[i].Height;
            }

            /*Формируем вторую часть значений*/

            height = height + Table1.Rows[0].Height * 2;
            height1 = height;
            width1 = 25;
            width = 25;
            k = 10;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = 0; i < NoCaIzm - 7; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[k].Width + 10;
                    k++;
                    //width1_1 = width;
                }
                height += Table1.Rows[j].Height;
                width = width1;
                k = 10;
            }
            width = (Table1.Columns[10].Width + 10)* (NoCaIzm - 7) + width;
            width1 = width;
            height = height1;
            k = 10 + NoCaIzm - 7;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = 0; i < Table1.Columns.Count - NoCaIzm - 3; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[k].Width + 10;
                    k++;
                }
                cordY = height;
                height += Table1.Rows[j].Height;
                width = width1;
                k = 10 + NoCaIzm - 7;
            }
            /*Cancel*/

        }


        public string filename;
        public void ExportToPDF1()
        {
            string head = @"Расчет линейного градуировочного графика";

            //Creating iTextSharp Table from the DataTable data
            PdfPTable pdfTable = new PdfPTable(Table1.ColumnCount);

            pdfTable.DefaultCell.Padding = 5;
            pdfTable.WidthPercentage = 100;

            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
            pdfTable.DefaultCell.BorderWidth = 1;
            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\georgia.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 12f, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font fontBold = new iTextSharp.text.Font(baseFont, 18f, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font fontBold1 = new iTextSharp.text.Font(baseFont, 10f, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font font1 = new iTextSharp.text.Font(baseFont, 5f, iTextSharp.text.Font.BOLD);
            // iTextSharp.text.Font fontLeft = new iTextSharp.text.Font(baseFont, 9f, iTextSharp.text.Font.NORMAL);

            //Adding Header row
            foreach (DataGridViewColumn column in Table1.Columns)
            {

                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText, fontBold1));
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                cell.Padding = 1;
                cell.PaddingBottom = 5;
                pdfTable.AddCell(cell);

            }

            //Adding DataRow
            foreach (DataGridViewRow row in Table1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    pdfTable.AddCell(new Phrase(Convert.ToString(cell.Value), font));


                }
            }




            var chartimage = new MemoryStream();
            chart1.SaveImage(chartimage, ChartImageFormat.Png);
            iTextSharp.text.Image Chart_Image = iTextSharp.text.Image.GetInstance(chartimage.GetBuffer());
            Chart_Image.ScalePercent(70f);




            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Pdf File |*.pdf";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Document doc = new Document(PageSize.A4, 50f, 20f, 20f, 0f);
                PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));

                doc.Open();
                //iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance("Image.jpeg");

                Paragraph welcomeParagraph = new Paragraph("Расчет линейного градуировочного графика\n", fontBold);
                welcomeParagraph.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                Paragraph Veshestvo2 = new Paragraph("Вещество: " + Veshestvo1, font);
                Paragraph wavelength2 = new Paragraph("Длина волны: " + wavelength1, font);
                Paragraph WidthCuvette2 = new Paragraph("Ширина кюветы: " + WidthCuvette, font); ;
                Paragraph BottomLine2 = new Paragraph("Нижняя граница обнаружения: " + BottomLine, font);
                Paragraph TopLine2 = new Paragraph("Верхняя граница обнаружения: " + TopLine, font);
                Paragraph ND2 = new Paragraph("НД: " + ND, font);
                Paragraph Description2 = new Paragraph("Примечание: " + Description, font);
                Paragraph DateTime2 = new Paragraph("Дата: " + DateTime, font);
                Paragraph Ispolnitel2 = new Paragraph("Исполнитель: " + Ispolnitel, font);
                Paragraph GradYrav = new Paragraph("Градуировочное уравнение: " + label14.Text, font);
                Paragraph Table1 = new Paragraph("Таблица исходных данных\n\n", font);


                PdfPTable table = new PdfPTable(5);
                PdfPCell cell = new PdfPCell(Veshestvo2);
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BorderWidth = 0;
                table.AddCell(cell);

                cell = new PdfPCell(ND2);
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                table.AddCell(cell);

                cell = new PdfPCell(wavelength2);
                cell.BorderWidth = 0;
                cell.Colspan = 5;
                table.AddCell(cell);

                cell = new PdfPCell(WidthCuvette2);
                cell.BorderWidth = 0;
                cell.Colspan = 5;
                table.AddCell(cell);

                cell = new PdfPCell(BottomLine2);
                cell.BorderWidth = 0;
                cell.Colspan = 5;
                table.AddCell(cell);

                cell = new PdfPCell(TopLine2);
                cell.BorderWidth = 0;
                cell.Colspan = 5;
                table.AddCell(cell);



                PdfPTable table1 = new PdfPTable(5);
                PdfPCell cell1 = new PdfPCell(DateTime2);
                cell1.BorderWidth = 0;
                cell1.Colspan = 5;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(Ispolnitel2);
                cell1.BorderWidth = 0;
                cell1.Colspan = 5;
                table1.AddCell(cell1);

                Paragraph welcomeParagraph1 = new Paragraph("\n", fontBold);

                doc.Add(welcomeParagraph);
                doc.Add(welcomeParagraph1);
                doc.Add(table);
                //  doc.Add(Veshestvo2);
                //  doc.Add(wavelength2);
                // doc.Add(WidthCuvette2);
                // doc.Add(BottomLine2);
                //  doc.Add(TopLine2);
                doc.Add(Description2);
                // doc.Add(welcomeParagraph1);
                doc.Add(Table1);
                //  doc.Add(welcomeParagraph1);
                doc.Add(pdfTable);
                doc.Add(welcomeParagraph1);
                doc.Add(GradYrav);
                doc.Add(welcomeParagraph1);
                doc.Add(Chart_Image);
                //  doc.Add(welcomeParagraph1);
                doc.Add(table1);
                //  doc.Add(ND2);

                // doc.Add(DateTime2);
                // doc.Add(Ispolnitel2);

                doc.Close();
                /*   string filename = Application.StartupPath;
                   filename = Path.GetFullPath(Path.Combine(filename, ".\\Test.pdf"));
                   wbrPdf.Navigate(filename);*/
                filename = sfd.FileName;

            }

            /*   Spire.Pdf.PdfDocument pdfdocument = new Spire.Pdf.PdfDocument();
               pdfdocument.LoadFromFile(filename);        
               pdfdocument.PrintDocument.Print();
               pdfdocument.Dispose();
               */
        }
    }
}