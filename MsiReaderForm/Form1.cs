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
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
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
                FillTreeView(myString);
            }
        }

        private void FillTreeView(string myString)
        {
            String fileName = Path.GetFileNameWithoutExtension(myString);
            treeView1.BeginUpdate();
            treeView1.Nodes.Add(fileName);
            treeView1.EndUpdate();
            
            List<String> allFileNames = new List<String>();
            MsiPull.DrawFromMsi(myString, ref allFileNames);
            //TODO: Skontat kako zap
            foreach (var name in allFileNames)
            {
                treeView1.Nodes[0].Nodes.Add(name.ToString());
            }
            treeView1.Enabled = true;
            
        }
        //this is where the magic happens

    }
}
