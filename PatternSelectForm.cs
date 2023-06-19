using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class PatternSelectForm : Form
    {
        public string selectedName = "";
        
        public PatternSelectForm(Word.Documents doc)
        {
            InitializeComponent();

            for (int i = 1; i < Globals.ThisAddIn.Application.ActiveDocument.Styles.Count; i++)
            {
                comboBox1.Items.Add(Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal);
            }
            comboBox1.Text = comboBox1.Items[0].ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            selectedName = comboBox1.SelectedItem.ToString();
            this.Close();
        }
    }
}
