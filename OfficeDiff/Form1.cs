﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeDiff
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] filename = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            textBox1.Text = filename[0];
        }

        private void textBox2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }

        }

        private void textBox2_DragDrop(object sender, DragEventArgs e)
        {
            string[] filename = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            textBox2.Text = filename[0];

        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string cmdbase = "C:\\Program Files\\TortoiseSVN\\Diff-Scripts";
            string cmd = cmdbase + "\\diff-doc.js";
            System.Diagnostics.Process.Start(cmd, textBox1.Text + " " + textBox2.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string cmdbase = "C:\\Program Files\\TortoiseSVN\\Diff-Scripts";
            string cmd = cmdbase + "\\diff-xls.js";
            System.Diagnostics.Process.Start(cmd, textBox1.Text + " " + textBox2.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string cmdbase = "C:\\Program Files\\TortoiseSVN\\Diff-Scripts";
            string cmd = cmdbase + "\\diff-ppt.js";
            System.Diagnostics.Process.Start(cmd, textBox1.Text + " " + textBox2.Text);
        }
    }
}
