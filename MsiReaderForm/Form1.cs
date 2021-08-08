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
            listView2.Items.Clear();
            listView2.Columns.Clear();
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
            listView2.Items.Clear();
            listView2.Columns.Clear();
            if (e.Node.Level == 0)
            {
                return;
            }
            List<String> columnString = new List<String>();
            List<String> dataString = new List<String>();
            int columnCount=0;
            MsiPull.GetItemData(fullPathName,e.Node.Text,ref columnString,ref columnCount,ref dataString);
            foreach(var item in columnString)
            {
                listView2.Columns.Add(item.ToString());
            }
            
            for(int i = 0; i< dataString.Count; i=i+columnCount)
            {
                int count = i;
                ListViewItem item = new ListViewItem(dataString[count]);
                count++;
                while (count < i + columnCount)
                {
                    item.SubItems.Add(dataString[count]);
                    count++;
                }
                listView2.Items.AddRange(new ListViewItem[] { item });
            }
            listView2.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            listView2.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
            //TODO:: replace textbox with another listbox variable sizes and have GetItemData fetch listview items or keep it a list but have it return a number of columns
        }
    }
}
