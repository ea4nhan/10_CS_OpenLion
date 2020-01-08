using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.IO;

namespace OPL
{
    public partial class ToBeType : DevExpress.XtraEditors.XtraForm
    {
        public ToBeType()
        {
            InitializeComponent();
        }

        private void ToBeType_Load(object sender, EventArgs e)
        {
            LabNewFileName.ForeColor = Color.Gray;
            LabNewFileName.BackColor = NewFileName.BackColor;
            NewFileName.Text = PowerMainForm.saveFilePath_N;
        }

        private void NewFileName_Click(object sender, EventArgs e)
        {
            if (sender.Equals(LabNewFileName))
            {
                NewFileName.Focus();
            }
            //elseif(sender.Equals(labelPwd))
            //{
            //   PwdTxt.Focus();
            //   //PwdTxt.PasswordChar = '*';
            //}
        }

        private void NewFileName_TextChanged(object sender, EventArgs e)
        {
            if (sender.Equals(NewFileName))
            {
                LabNewFileName.Visible = NewFileName.Text.Length < 1;
            }
            //else if(sender.Equals(PwdTxt))
            //{
            //   labelPwd.Visible = PwdTxt.Text.Length < 1;
            //   PwdTxt.PasswordChar = '*';//隐藏输入的密码
            //}
        }

        private void NewFileCanel_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void NewFileOK_Click(object sender, EventArgs e)
        {
            if (NewFileName.Text.Trim() != "")
            {
                string[] filespath;
                PowerMainForm.saveFilePath_D = @".\Save\" + NewFileName.Text.Trim() + ".ini";
                PowerMainForm.saveFilePath_N = NewFileName.Text.Trim();

                try { filespath = Directory.GetFiles(@".\Save\"); }
                catch { return; }
                HYQFileInfoList fileList = new HYQFileInfoList(filespath);
                foreach (FileInfoWithIcon file in fileList.list)
                {
                    if (System.IO.Path.GetFileNameWithoutExtension(file.fileInfo.Name) == NewFileName.Text.Trim())
                    {
                        DialogResult result = DevExpress.XtraEditors.XtraMessageBox.Show("已经存在重复的OPL文件名，请重新输入！", "Info", MessageBoxButtons.OK, MessageBoxIcon.Question);
                        if (result == DialogResult.OK)
                        return;
                    }
                }
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void NewFileName_KeyPress(object sender, KeyPressEventArgs e)
        {
            //阻止从键盘输入键
            e.Handled = false;
            //当输入为0-9的数字、小数点、回车和退格键时不阻止
            if (e.KeyChar == '[' ||e.KeyChar == ']' ||e.KeyChar == '.' ||
                e.KeyChar == '=' ||e.KeyChar == '#' ||e.KeyChar == '@' ||
                e.KeyChar == '^' ||e.KeyChar == '!' ||e.KeyChar == '*' ||
                e.KeyChar == '|' || e.KeyChar == '/' || e.KeyChar == '?')
            {
                e.Handled = true;
                DevExpress.XtraEditors.XtraMessageBox.Show("请不要输入下列字符：[].=#@^!*|/?",
                "Warning", MessageBoxButtons.OK);
            } 
        }




    }
}