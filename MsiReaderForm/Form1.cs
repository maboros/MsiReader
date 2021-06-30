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
using OpenMcdf.Extensions;
using OpenMcdf.Extensions.OLEProperties;
using OpenMcdf;

namespace MsiReaderForm
{
    public partial class Form1 : Form
    {
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
                FillView(myString);
            }

        }

        private void FillView(String fullPathName)
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
           // var propertyList = MsiPull.getSummaryInformation(fullPathName);


           //// *Ovaj getSummaryInfromation se krši.



           // foreach (var property in propertyList)
           // {
           //     var row = new string[] { property.name, property.value };
           //     var viewItem = new ListViewItem(row);
           //     viewItem.Tag = property;
           //     listView1.Items.Add(viewItem);
           // }

        }
    }
}
