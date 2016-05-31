using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using Microsoft.VisualBasic;

namespace Grouper
{
    public partial class Form1 : Form
    {
        Dictionary<string, XmlDocument> groups = new Dictionary<string, XmlDocument>();
        string subject = "class";
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                string[] fls = Directory.GetFiles("Data");
                foreach (string fn in fls)
                {
                    if (fn.Contains(subject))
                    {
                        XmlDocument xmldoc = new XmlDocument();
                        xmldoc.Load(fn);
                        XmlNode root = xmldoc.SelectSingleNode(xmldoc.DocumentElement.Name);
                        if (root.Attributes["id"].Value == subject && root.Name == "Info")
                        {
                            if (checkBox1.Checked)
                                foreach (XmlNode node in root.SelectNodes(subject))
                                {
                                    listBox1.Items.Add(node.Attributes["id"].Value);
                                }
                        }
                        else if (root.Attributes["id"].Value == subject && root.Name == "Group")
                        {
                            if (checkBox2.Checked)
                            {
                                listBox1.Items.Add(root.Attributes["name"].Value);
                                groups[root.Attributes["name"].Value] = xmldoc;
                            }
                        }
                    }
                }
                listBox1.SelectedIndex = 0;
            }
            catch { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (groups.ContainsKey(listBox1.SelectedItem.ToString()))
                {
                    XmlNode root = groups[listBox1.SelectedItem.ToString()].SelectSingleNode(groups[listBox1.SelectedItem.ToString()].DocumentElement.Name);
                    foreach (XmlNode node in root.SelectNodes(subject))
                        if (!listBox2.Items.Contains(node.Attributes["id"].Value))
                            listBox2.Items.Add(node.Attributes["id"].Value);
                }
                else
                {
                    if (!listBox2.Items.Contains(listBox1.SelectedItem))
                        listBox2.Items.Add(listBox1.SelectedItem);
                }
                if (listBox1.SelectedIndex + 1 < listBox1.Items.Count)
                    ++listBox1.SelectedIndex;
                else
                    listBox1.SelectedIndex = 0;
            }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                int temp = listBox2.SelectedIndex;
                listBox2.Items.Remove(listBox2.SelectedItem);
                listBox2.SelectedIndex = temp;
            }
            catch { }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            Form1_Load(sender, e);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string array = "";
            foreach(string item in listBox2.Items)
            {
                array += item;
                array += ",";
            }
            MessageBox.Show(array);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string fn;
            do
            {
                fn = Interaction.InputBox("input a group name:", "Inquiry");
            }
            while (fn.Contains("<"));
            int i = 2;
            for (; File.Exists("Data/" + subject + i + ".xml"); ++i) ;
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.CreateXmlDeclaration("1.0", "utf-8", null);
            XmlElement root = xmldoc.CreateElement("Group");
            root.SetAttribute("id", subject);
            root.SetAttribute("name", fn);
            xmldoc.AppendChild(root);
            foreach(string item in listBox2.Items)
            {
                XmlElement node = xmldoc.CreateElement("class");
                node.SetAttribute("id", item);
                root.AppendChild(node);
            }
            xmldoc.Save("Data/" + subject + i + ".xml");
            MessageBox.Show("You might be willing to click \"reload\"!");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            Form1_Load(sender, e);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            Form1_Load(sender, e);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                if (groups.ContainsKey(listBox1.SelectedItem.ToString()))
                {
                    File.Delete(groups[listBox1.SelectedItem.ToString()].BaseURI.Replace("file:///",""));
                    groups.Remove(listBox1.SelectedItem.ToString());
                    listBox1.Items.Clear();
                    Form1_Load(sender, e);
                }
                else
                {
                    MessageBox.Show("This program can only delete groups!");
                }
            }
            catch { }
        }
    }
}
