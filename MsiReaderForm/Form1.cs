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
    static class Win32Error
    {
        public const int NO_ERROR = 0;
        public const int ERROR_NO_MORE_ITEMS = 259;

    }
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
                textBox1.Text = myString;
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
            
        }
        //this is where the magic happens

    }
}
