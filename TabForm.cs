using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class TabForm : UserControl
    {
        const int sizeX = 129;
        const int sizeY = 28;
        int DocNums;
        Color ColorSelected = Color.FromArgb(255, 255, 255);
        Color ColorNoSelected = Color.FromArgb(230, 231, 232);

        public TabForm(Word.Documents Docs)
        {
            InitializeComponent();


            DocNums = Docs.Count;

            if(DocNums>1)
            {
                Button[] buttonNew = new Button[DocNums-1];
                for(int i=0;i< DocNums - 1;i++)
                {
                    buttonNew[i] = new Button();
                    buttonNew[i].Text = System.IO.Path.GetFileNameWithoutExtension(Docs[DocNums - 1-i].Name);
                    buttonNew[i].Size = new Size(sizeX, sizeY);
                    buttonNew[i].Name = "checkBox" + (i + 2).ToString();
                    buttonNew[i].BackColor = ColorNoSelected;
                    buttonNew[i].Location = new Point((i + 1) * sizeX, 0);
                    buttonNew[i].Visible = true;
                    buttonNew[i].Parent = this;
                    buttonNew[i].Click += new System.EventHandler(this.button1_Click);
                    this.Controls.Add(buttonNew[i]);
                    toolTip1.SetToolTip(buttonNew[i], Docs[DocNums - 1 - i].Path + "\\" + Docs[DocNums - 1 - i].Name);

                }

            }


            button1.Text = System.IO.Path.GetFileNameWithoutExtension(Docs[DocNums].Name);
            button1.BackColor = ColorNoSelected;
            toolTip1.SetToolTip(button1, Docs[DocNums].Path + "\\" + Docs[DocNums].Name);


            //遍历标签
            foreach (Control ctrl in this.Controls)
            {
                //if (ctrl.Text.Equals(System.IO.Path.GetFileNameWithoutExtension(Docs.Application.ActiveDocument.Name)))
                //{
                //    ctrl.BackColor = ColorSelected;
                //}
                if (ctrl is Button)
                {
                    if (Docs.Application.ActiveDocument.Path + "\\" + Docs.Application.ActiveDocument.Name == toolTip1.GetToolTip(ctrl))
                    {
                        ctrl.BackColor = ColorSelected;
                    }
                }
            }

        }


        private void button1_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;

            //选中、跳转标签
            foreach(Word.Window wd in Globals.ThisAddIn.Application.Windows)
            {
                if (wd.Document.Path + "\\" + wd.Document.Name == toolTip1.GetToolTip(button))
                {
                    wd.Activate();
                }
            }


            //遍历标签
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is Button)
                {
                    ctrl.BackColor = ColorNoSelected;
                }
                if (ctrl.Name == button.Name)
                {
                    ctrl.BackColor = ColorSelected;
                }
            }

            //Globals.ThisAddIn.Application.Windows[int.Parse(checkBox.Name.Substring(checkBox.Name.Length - 1, 1))].Activate();

            //MessageBox.Show(button.Name.Substring(button.Name.Length-1,1));
        }
    }
}
