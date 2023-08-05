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

        static public List<string> DocNamesList_pane = new List<string>();


        public TabForm(List<string> Docs)
        {
            InitializeComponent();


            DocNums = Docs.Count;
            DocNamesList_pane = Docs;


            if (DocNums>1)
            {
                Button[] buttonNew = new Button[DocNums-1];
                for (int i = 0; i < DocNums - 1; i++)
                {
                    buttonNew[i] = new Button();
                    buttonNew[i].Text = System.IO.Path.GetFileNameWithoutExtension(Docs[DocNums - 2 - i]);
                    buttonNew[i].Size = new Size(sizeX, sizeY);
                    buttonNew[i].Name = "button" + (i + 2).ToString();
                    buttonNew[i].BackColor = ColorNoSelected;
                    buttonNew[i].Location = new Point((i + 1) * sizeX, 0);
                    buttonNew[i].Visible = true;
                    buttonNew[i].Parent = this;
                    buttonNew[i].Click += new System.EventHandler(this.button1_Click);
                    this.Controls.Add(buttonNew[i]);
                    toolTip1.SetToolTip(buttonNew[i], Docs[DocNums - 2 - i]);

                    buttonNew[i].MouseDown += button1_MouseDown;
                    buttonNew[i].MouseMove += button1_MouseMove;
                }

            }


            button1.Text = System.IO.Path.GetFileNameWithoutExtension(Docs[DocNums - 1]);
            button1.BackColor = ColorNoSelected;
            toolTip1.SetToolTip(button1, Docs[DocNums - 1]);

            button1.MouseDown += button1_MouseDown;
            button1.MouseMove += button1_MouseMove;


            //遍历标签
            foreach (Control ctrl in this.Controls)
            {
                //if (ctrl.Text.Equals(System.IO.Path.GetFileNameWithoutExtension(Docs.Application.ActiveDocument.Name)))
                //{
                //    ctrl.BackColor = ColorSelected;
                //}
                if (ctrl is Button)
                {
                    if (Globals.ThisAddIn.Application.ActiveDocument.Path + "\\" + Globals.ThisAddIn.Application.ActiveDocument.Name == toolTip1.GetToolTip(ctrl))
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

        private Button b;
        private Point location;

        private void button1_MouseDown(object sender, MouseEventArgs e)
        {
            location = e.Location;
            b = sender as Button;

            //右键关闭文档
            if (e.Button == MouseButtons.Right)
            {
                Globals.ThisAddIn.Application.ActiveDocument.Save();

                Object saveChanges = Word.WdSaveOptions.wdSaveChanges;
                Object originalFormat = Type.Missing;
                Object routeDocument = Type.Missing;

                Globals.ThisAddIn.Application.ActiveDocument.Close(ref saveChanges, ref originalFormat, ref routeDocument);

                //若已经没有文档存在，则关闭应用程序
                if (Globals.ThisAddIn.Application.Documents.Count == 0)
                {
                    Globals.ThisAddIn.Application.Quit(Type.Missing, Type.Missing, Type.Missing);
                }
                else
                {
                    foreach (Word.Window wd in Globals.ThisAddIn.Application.Windows)
                    {
                        if (wd.Document.Path + "\\" + wd.Document.Name == DocNamesList_pane[0])
                        {
                            wd.Activate();
                        }
                    }
                }
            }
        }


        private void button1_MouseMove(object sender, MouseEventArgs e)
        {
            Button button = (Button)sender;

            string doc0 = toolTip1.GetToolTip(button);
            string doc1;

            int pos_x0 = button.Location.X;
            int pos_x;

            if (DocNums > 1)
            {
                if (e.Button == MouseButtons.Left)
                {
                    pos_x = b.Location.X + (e.X - location.X);

                    pos_x = ((int)(pos_x / sizeX)) * sizeX;  //移动整数个标签位置

                    foreach (Control ctrl in this.Controls)
                    {
                        if (ctrl is Button)
                        {
                            if (ctrl.Location.X == pos_x)
                            {
                                ctrl.Location = new Point(pos_x0, 0);

                                b.Location = new Point(pos_x, 0);


                                doc1 = toolTip1.GetToolTip(ctrl);

                                int index0 = DocNamesList_pane.FindIndex(x => x == doc0);
                                int index1 = DocNamesList_pane.FindIndex(x => x == doc1);

                                string temp = DocNamesList_pane[index0];
                                DocNamesList_pane[index0] = DocNamesList_pane[index1];
                                DocNamesList_pane[index1] = temp;
                            }
                        }
                    }
                }
            }
       
        }


    }
}
