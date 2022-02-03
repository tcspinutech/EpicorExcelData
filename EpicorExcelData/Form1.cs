using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EpicorExcelData.Models;

namespace EpicorExcelData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            var document = new Excel();
            document.ProcessFiles();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
