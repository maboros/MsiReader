using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MsiReader;

namespace MsiReaderForm
{
    public partial class Form1 : Form
    {
        private string fullPathName;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            listView1.Items.Clear();
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Browse MSI Files",
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "msi",
                Filter = "msi files (*.msi)|*.msi|All files (*.*)|*.*",
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                
                String myString= fileDialog.FileName;
                myString = myString.Replace(@"\", "/");
                fullPathName = myString;
                FillView();
            }

        }

        private void FillView()
        {
            treeView1.Nodes.Clear();
            listView1.Items.Clear();
            String fileName = Path.GetFileNameWithoutExtension(fullPathName);
            treeView1.BeginUpdate();
            treeView1.Nodes.Add(fileName);
            treeView1.EndUpdate();
            
            List<String> allFileNames = new List<String>();
            MsiPull.DrawFromMsi(fullPathName, ref allFileNames);
            foreach (var name in allFileNames)
            {
                treeView1.Nodes[0].Nodes.Add(name.ToString());
            }
            treeView1.Enabled = true;
           var propertyList = MsiPull.getSummaryInformation(fullPathName);

            foreach (var property in propertyList)
            {
                var row = new string[] { property.name, property.value };
                var viewItem = new ListViewItem(row);
                viewItem.Tag = property;
                listView1.Items.Add(viewItem);
            }

        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Level == 0)
            {
                return;
            }
            List<String> dataString = new List<String>();
            textBox1.Text = MsiPull.GetItemData(fullPathName,ref dataString,(uint)e.Node.Index);
            foreach(var item in dataString)
            {
                textBox1.Text += item.ToString();
            }
        }
    }
}
