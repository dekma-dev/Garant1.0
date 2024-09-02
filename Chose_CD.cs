using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Garant1._0
{
    public partial class Chose_CD : Form
    {
        ListBox cbb;
        public Chose_CD(string[] codes, ref ListBox cb1)
        {
            InitializeComponent();
            cbb = cb1;
            CheckBox[] cbs = new CheckBox[codes.Length];
            int x = 5,
                y = 0;
            for (int i = 0; i < cbs.Length; i++)
            {
                cbs[i] = new CheckBox();
                cbs[i].Text = codes[i];
                if (codes[i][codes[i].Length - 1] == '!')
                {
                    cbs[i].Font = new Font(cbs[i].Font.FontFamily, cbs[i].Font.Size, FontStyle.Bold);
                }
                cbs[i].Location = new Point(x, y);
                cbs[i].Width = 200;
                y += 25;
                if (y > 700-cbs[i].Height*4)
                {
                    x += 250;
                    y = 0;
                }
                cbs[i].CheckedChanged += Chose_CD_CheckedChanged;
            }
            this.Controls.AddRange(cbs);

            string codes_braka = "";
            foreach (string d in cbb.Items)
            {
                codes_braka += d + '\n';
            }
            for (int i = 0; i < cbs.Length; i++)
            {
                if(codes_braka.IndexOf(cbs[i].Text)!=-1)
                {
                    cbs[i].Checked = true;
                }
            }

            void Chose_CD_CheckedChanged(object sender, EventArgs e)
            {
                string text = "";
                //bool f = true;
                cbb.Items.Clear();
                for (int i = 0; i < cbs.Length; i++)
                {
                    if (cbs[i].Checked)
                    {
                        //if (f == false) text += '\n';
                        // f = false;
                        cbb.Items.Add(cbs[i].Text);
                    }
                    cbb.Text = text;
                }
            }
        }

        private void Chose_CD_Load(object sender, EventArgs e)
        {

        }
    }
}
