﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace Analis200
{
    public partial class PriborInformation : Form
    {
        Analis _Analis;
        public PriborInformation(Analis parent)
        {
            InitializeComponent();
            this._Analis = parent;
        }
        string Information = "InformationPribor.xml";
        private void PriborInformation_Load(object sender, EventArgs e)
        {
            /* string Model;
             XmlDocument xDoc = new XmlDocument();
             xDoc.Load(@Information);
             XmlNodeList nodes = xDoc.ChildNodes;
             foreach (XmlNode n in nodes)
             { // Обрабатываем в цикле только Pribors
                 if ("Pribors".Equals(n.Name))
                 {
                     // Читаем в цикле вложенные значения Izmerenie
                     for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                     {
                         //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                         for (XmlNode k = n.FirstChild; k != null; k = n.NextSibling)
                         {
                             if ("Model".Equals(k.Name) && k.FirstChild != null)
                             {
                                 Model = k.FirstChild.Value;
                                 Model1.Items.Add(Model);
                             }

                         }
                     }
                 }
             }*/
            ///    Model1.SelectedIndex = 0;
            textBox3.Enabled = false;
            Pribor();
        }
        public void Pribor()
        {
            string model = @"pribor/model";
          //  string s = Model1.SelectedItem.ToString();
            string SerNomer_Text = @"pribor/SerNomer";
            string InventarNomer_Text = @"pribor/InventarNomer";
            string SrokIstech_Text = @"pribor/SrokIstech";
            string Poveren_Text = @"pribor/Poveren";
            StreamReader fs = new StreamReader(model);
            string model1;
            model1 = fs.ReadLine();
            int index = Model1.FindString(model1);
            if (index != -1)
            {
                Model1.SelectedIndex = index;

            }
            else
            {
                Model1.SelectedIndex = 0;

            }
            fs.Close();

            StreamReader fs1 = new StreamReader(SerNomer_Text);            
            textBox1.Text = fs1.ReadLine();            
            fs1.Close();

            StreamReader fs2 = new StreamReader(InventarNomer_Text);
            textBox2.Text = fs2.ReadLine();
            fs2.Close();

            StreamReader fs3 = new StreamReader(SrokIstech_Text);
            textBox3.Text = fs3.ReadLine();
            fs3.Close();

            if(textBox3.Text != "")
            {
                checkBox1.Checked = true;
            }
            else
            {
                textBox3.Enabled = false;
            }

            StreamReader fs4 = new StreamReader(Poveren_Text);
            dateTimePicker1.Text = fs4.ReadLine();
            fs4.Close();


        }

        private void Model1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s1 = "";
            string model = @"pribor/model";
            string s = Model1.SelectedItem.ToString();

            File.WriteAllText(model, string.Empty);
            File.AppendAllText(model, s, Encoding.UTF8);

            string SerNomer = textBox1.Text;
            string InventarNomer = textBox2.Text;
            string SrokIstech = textBox3.Text;
            string SerNomer_Text = @"pribor/SerNomer";
            string InventarNomer_Text = @"pribor/InventarNomer";
            string SrokIstech_Text = @"pribor/SrokIstech";
            string Poveren_Text = @"pribor/Poveren";

            File.WriteAllText(SerNomer_Text, string.Empty);
            File.AppendAllText(SerNomer_Text, textBox1.Text, Encoding.UTF8);
            File.WriteAllText(InventarNomer_Text, string.Empty);
            File.AppendAllText(InventarNomer_Text, textBox2.Text, Encoding.UTF8);
            File.WriteAllText(SrokIstech_Text, string.Empty);
            File.AppendAllText(SrokIstech_Text, textBox3.Text, Encoding.UTF8);
            File.WriteAllText(Poveren_Text, string.Empty);
            File.AppendAllText(Poveren_Text, dateTimePicker1.Value.ToString("dd.MM.yyyy"), Encoding.UTF8);
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked == true)
            {
                textBox3.Enabled = true;
            }
            else
            {
                textBox3.Enabled = false;
            }
        }
    }
}
