using System;
using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using DevExpress.XtraEditors.Repository;
//using System.Drawing.Imaging;
using Aspose.Cells;
using System.Text.RegularExpressions;
using System.Linq;

namespace OPL
{
    #region MsgBox定义
    public class MessageUtil
    {
        /// <summary>
        /// 显示一般的提示信息
        /// </summary>
        /// <param name="message">提示信息</param>
        public static DialogResult ShowTips(string message)
        {
            return MessageBox.Show(message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 显示警告信息
        /// </summary>
        /// <param name="message">警告信息</param>
        public static DialogResult ShowWarning(string message)
        {
            return MessageBox.Show(message, "警告信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// 显示错误信息
        /// </summary>
        /// <param name="message">错误信息</param>
        public static DialogResult ShowError(string message)
        {
            return MessageBox.Show(message, "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// 显示询问用户信息，并显示错误标志
        /// </summary>
        /// <param name="message">错误信息</param>
        public static DialogResult ShowYesNoAndError(string message)
        {
            return MessageBox.Show(message, "错误信息", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
        }

        /// <summary>
        /// 显示询问用户信息，并显示提示标志
        /// </summary>
        /// <param name="message">错误信息</param>
        public static DialogResult ShowYesNoAndTips(string message)
        {
            return MessageBox.Show(message, "提示信息", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 显示询问用户信息，并显示警告标志
        /// </summary>
        /// <param name="message">警告信息</param>
        public static DialogResult ShowYesNoAndWarning(string message)
        {
            return MessageBox.Show(message, "警告信息", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// 显示询问用户信息，并显示提示标志
        /// </summary>
        /// <param name="message">错误信息</param>
        public static DialogResult ShowYesNoCancelAndTips(string message)
        {
            return MessageBox.Show(message, "提示信息", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
        }
    }
    #endregion

    #region HyqIni定义
    /// <summary>
    /// 修改Ini
    /// </summary>
    class HyqIni
    {
        //static string IniFileName = @".\Default.ini";
        static char[] TrimChar = { ' ', '\t' };
        static ArrayList ls = new ArrayList();

        #region 子函数 string[] GetSects()：获取全“段数”，通过符号“[”和“]”来判断是否为有效段名
        public static string[] GetSects(string IniFileName)
        {
            string[] Sects = null;

            if (File.Exists(IniFileName))
            {
                string str;
                ls.Clear();
                StreamReader sr = new StreamReader(IniFileName, Encoding.GetEncoding("GB2312"));
                /* 逐行读取 */
                while ((str = sr.ReadLine()) != null)
                {
                    str = str.Trim();
                    /* 每一段的开始标志符和结束标志符 */
                    if ((str.StartsWith("[")) && (str.EndsWith("]")))
                    {
                        try
                        {
                            str = str.Substring(1, str.Length - 2);
                            ls.Add(str);
                        }
                        catch { }
                    }
                }
                sr.Close();
                /* Array转换为String[] */
                if (ls.Count > 0)
                {
                    Sects = new string[ls.Count];
                    for (int i = 0; i < ls.Count; i++)
                    {
                        Sects[i] = ls[i].ToString();
                    }
                }
            }
            return Sects;
        }
        #endregion

        #region 子函数 string[] GetItems()：获取某“段数”下的行数
        public static string[] GetItems(string sect, string IniFileName)
        {
            string[] Items = null;
            ArrayList result = new ArrayList();
            bool secOK = false;

            if (File.Exists(IniFileName))
            {
                string str;
                ls.Clear();
                StreamReader sr = new StreamReader(IniFileName, Encoding.GetEncoding("GB2312"));
                /* 逐行读取 */
                while ((str = sr.ReadLine()) != null)
                {
                    str = str.Trim();

                    if ((secOK == true) && (str != null) && (str != "")
                        &&(!((str.StartsWith("[")) && (str.EndsWith("]")))))
                    {
                        try
                        {
                            result.Add(str);
                        }
                        catch { }
                    }
                    
                    /* 每一段的开始标志符和结束标志符 */
                    if (str.StartsWith("[" + sect + "]"))
                    {
                        secOK = true;
                    }
                    else if ((str.StartsWith("[")) && (str.EndsWith("]")))
                    {
                        secOK = false;
                    }
                }
                sr.Close();
                /* Array转换为String[] */
                if (result.Count > 0)
                {
                    Items = new string[result.Count];
                    for (int i = 0; i < result.Count; i++)
                    {
                        Items[i] = result[i].ToString();
                    }
                }
            }
            return Items;
        }
        #endregion

        #region 子函数 void PutINI(sect, keystr, valuestr, IniFileName)：在[段]中替换关键字的参数（字符）或新增
        //[段sect]
        //关键字keystr = valuestr;（替换此参数）
        //null则删除此段
        public static void PutINI(string sect, string keystr, string valuestr, string IniFileName)
        {
            if ((keystr.Trim() == "") || (keystr == null)) return; //如果对象为空，则退出
            if (File.Exists(IniFileName))
            {
                /*打开文件并读取*/
                PutINI_read(IniFileName);

                /*处理*/
                PutINI_main(sect, keystr, valuestr, IniFileName);

            } //如果文件不存在，则需要建立文件。
            else
            {
                ls.Clear();
                ls.Add("##File created by Han Yiqi：" + DateTime.Now.ToString() + "##");
                ls.Add("[" + sect + "]");
                ls.Add(keystr.Trim() + "=" + valuestr);
            }

            /*处理完之后保存*/
            PutINI_write(IniFileName);

            //File.WriteAllLines(IniFileName, strList);
        }

        #region 子函数 PutINI_read：在[段]中替换关键字的参数（字符）或新增
        public static void PutINI_read(string IniFileName)
        {
            /*打开文件并读取*/
            string str;
            StreamReader sr = new StreamReader(IniFileName, Encoding.GetEncoding("GB2312"));
            ls.Clear();
            while ((str = sr.ReadLine()) != null)
            {
                ls.Add(str);
            }
            sr.Close();
        }
        #endregion

        #region 子函数 PutINI_main：在[段]中替换关键字的参数（字符）或新增
        public static int PutINI_main(string sect, string keystr,string valuestr ,string IniFileName)
        {
            /*处理ls文件*/
            string str;
            bool SectOK = false;
            bool SetOK = false;
            int pos1;
            string substr;
            int result = -1;
            //开始寻找关键字，如果找不到，则在这段的最后一行插入，然后再整体的保存一下INI文件。
            for (int i = 0; i < ls.Count; i++)
            {
                str = ls[i].ToString();
                if (str.StartsWith("[") && str.EndsWith("]") && SectOK) //先判断是否到下一段中了,如果本来就是最后一段，那就有可能永远也不会发生了。
                {
                    SetOK = true; //如果在这一段中没有找到，并且已经要进入下一段了，就直接在这一段末添加了。
                    if (valuestr != null)//add by HYQ，如果null则删除，不需要添加
                    {
                        ls.Insert(i, keystr.Trim() + "=" + valuestr);
                        result = i;
                    }
                    break;//如果到下一段了，则直接退出就好。
                }
                if (SectOK)
                {
                    pos1 = str.IndexOf("=");
                    if (pos1 > 1)
                    {
                        substr = str.Substring(0, pos1);
                        substr.Trim(TrimChar);
                        //如果在这一段中找到KEY了，直接修改就好了。
                        if (substr.Equals(keystr, StringComparison.OrdinalIgnoreCase) && SectOK) //是在此段中，并且KEYSTR前段也能匹配上。
                        {
                            SetOK = true;
                            if (valuestr != null)//add by HYQ，如果null则删除，不需要添加
                            {
                                ls[i] = keystr.Trim() + "=" + valuestr;
                            }
                            else
                            {
                                ls[i] = null;
                            }
                            result = i;
                            break;
                        }
                    }
                }
                if (str.StartsWith("[" + sect + "]")) //判断是否到需要的段中了。
                    SectOK = true;
            }
            if (SetOK == false)
            {
                SetOK = true;
                if (!SectOK) //如果没有找到段，则需要再添加段。
                {
                    ls.Add("[" + sect + "]");
                }
                if (valuestr != null)//add by HYQ，如果null则删除，不需要添加
                {
                    ls.Add(keystr.Trim() + "=" + valuestr);
                    result = 0;
                }
            }
            return result;
        }
        #endregion

        #region 子函数 PutINI_write：在[段]中替换关键字的参数（字符）或新增
        public static void PutINI_write(string IniFileName)
        {
            //删除源文件。
            if (File.Exists(IniFileName))
            {
                File.Delete(IniFileName);
            }
           
            FileStream fs = new FileStream(IniFileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding("GB2312"));
            //string[] strList = new string[ls.Count];
            for (int i = 0; i < ls.Count; i++)
            {
                //strList[i] = ls[i].ToString();
                if (ls[i] != null)//add by HYQ，如果null则删除，不需要添加
                {
                    sw.WriteLine(ls[i].ToString());
                }
            }
            try
            {
                sw.Flush();
                fs.Close();
                sw.Close();
            }
            catch { }
        }
        #endregion
        #endregion

        #region 子函数 string GetINI(sect, keystr, defaultstr, IniFileName)：在[段]搜索关键字的参数(字符)并且返回
        public static string GetINI(string sect, string keystr, string defaultstr, string IniFileName)
        {
            //[段sect]
            //关键字keystr = retstr;（替换此参数）
            string retstr = defaultstr;
            if (File.Exists(IniFileName))
            {
                bool SectOK = false;
                int pos1;
                string substr;
                string str;
                ls.Clear();
                StreamReader sr = new StreamReader(IniFileName, Encoding.GetEncoding("GB2312"));
                while ((str = sr.ReadLine()) != null)
                {
                    str = str.Trim();
                    if (str.StartsWith("[") && SectOK) //先判断是否到下一段中了。
                    {
                        break;//如果到下一段了，则直接退出就好。
                    }
                    if (SectOK)
                    {
                        pos1 = str.IndexOf("=");
                        if (pos1 > 1)
                        {
                            substr = str.Substring(0, pos1);
                            substr.Trim(TrimChar);
                            if (substr.Equals(keystr, StringComparison.OrdinalIgnoreCase)) //是在此段中，并且KEYSTR前段也能匹配上。
                            {
                                retstr = str.Substring(pos1 + 1).Trim(TrimChar);
                                break;
                            }
                        }
                    }
                    if (str.StartsWith("[" + sect + "]")) //判断是否到需要的段中了。
                        SectOK = true;
                }
                sr.Close();
            }
            return retstr;
        }
        #endregion

        #region 子函数 int GetINI(sect, keystr, defaultint, IniFileName)：在[段]搜索关键字的参数（int）并且返回
        public static int GetINI(string sect, string keystr, int defaultint, string IniFileName)
        {
            string intStr = GetINI(sect, keystr, Convert.ToString(defaultint), IniFileName);
            try
            {
                return Convert.ToInt32(intStr);
            }
            catch
            {
                return defaultint;
            }
        }
        #endregion

        #region 子函数 void PutINI(sect, keystr, valueint, IniFileName)：在[段]替换关键字的参数（int）
        public static void PutINI(string sect, string keystr, int valueint, string IniFileName)
        {
            PutINI(sect, keystr, valueint.ToString(), IniFileName);
        }
        #endregion

        #region 子函数 bool GetINI(sect, keystr, defaultbool, IniFileName)：在[段]搜索关键字的参数（bool）并且返回
        //读布尔 
        public static bool GetINI(string sect, string keystr, bool defaultbool, string IniFileName)
        {
            try
            {
                return Convert.ToBoolean(GetINI(sect, keystr, Convert.ToString(defaultbool), IniFileName));
            }
            catch
            {
                return defaultbool;
            }
        }
        #endregion

        #region 子函数 void PutINI(sect, keystr, valuebool, IniFileName)：在[段]修改关键字的参数（bool）
        public static void PutINI(string sect, string keystr, bool valuebool, string IniFileName)
        {
            PutINI(sect, keystr, Convert.ToString(valuebool), IniFileName);
        }
        #endregion

        /////////////////////////////////////////////////////////////////////////
        //判断如果是默认值，则替换；不然则不替换
        //使用此INI文件的特例（自己使用）
        public string GetParam(string KeyStr, string Default, string IniFileName)
        {
            string str;
            str = GetINI("Params", KeyStr, "???", IniFileName);
            if (str == "???")
            {
                PutINI("Params", KeyStr, Default, IniFileName);
                str = Default;
            }
            return str;
        }
        public void UpdateParam(string KeyStr, string ValueStr, string IniFileName)
        {
            PutINI("Params", KeyStr, ValueStr, IniFileName);
        }
    }
    #endregion

    #region HyqDG定义
    /// <summary>
    /// 为了Dategridview定义的函数
    /// </summary>
    public class HyqDG
    {
        public static int Dev_AddNewRow(DevExpress.XtraGrid.Views.Grid.GridView DG, List<string> sArry)
        {
            DevExpress.XtraGrid.Views.Grid.GridRow row = new DevExpress.XtraGrid.Views.Grid.GridRow();

            int rowcnt = DG.DataRowCount;
            DG.AddNewRow();
            try
            {
                int i = 0;
                foreach (string str in sArry)
                {
                    try { DG.SetRowCellValue(rowcnt, DG.Columns[i], sArry[i].ToString()); }
                    catch { }
                    
                    i++;
                }
            }
            catch { }
            return rowcnt;
        }

        public static void Dev_EditCol(DevExpress.XtraGrid.Columns.GridColumn col, string name, bool multiline, int width)
        {
            col.FieldName = name;
            col.Width = width;
            //col.MaxWidth = width;
            col.AppearanceHeader.Font = new System.Drawing.Font("STXihei", 11, FontStyle.Regular);
            col.AppearanceCell.Font = new System.Drawing.Font("STXihei", 11, FontStyle.Regular);
            if (multiline)
            {
                RepositoryItemMemoEdit repoMemo = new RepositoryItemMemoEdit();
                repoMemo.WordWrap = true;
                repoMemo.AutoHeight = true;
                col.ColumnEdit = repoMemo;
                col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                col.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            }
            col.Visible = true;
        }
    }
    #endregion

    #region HyqCtrl定义
    /// <summary>
    /// 为了通用控件定义的函数
    /// </summary>
    public class HyqCtrl
    {
        public static Control ResetCtrl(Control form, string name, int font,
            double X, double Y, double Y0, double W, double H)
        {
            Color back = Color.FromArgb(45, 45, 48);
            Color white = Color.FromArgb(122, 193, 255);
            Control Ctrl = Func.FindControl(form, name);
            Ctrl.Font = Func.FontAdjust(form, font + 4, font + 2, font);            
            Ctrl.Location = Func.PositionAdjust(form, Ctrl, X, Y , Y0);
            Ctrl.Visible = true;
            if (W != 0) Ctrl.Width = (int)(form.Width * W);
            if (H != 0) Ctrl.Height = (int)((form.Height - Y0) * H);
            Ctrl.Anchor = AnchorStyles.None;
            if (Ctrl.GetType().Equals(typeof(PictureBox)))
            {
                Ctrl.MouseLeave += new EventHandler(Func.General_MouseLeave);
                Ctrl.MouseMove += new MouseEventHandler(Func.General_MouseMove);
                Ctrl.Tag = 0;
            }
            else 
            {
                Ctrl.BackColor = back;
                Ctrl.ForeColor = white;    
            }
            return Ctrl;
        }

        public static void NewLabel(Control form, string name, string text, int font, Color color, ContentAlignment allign, double X, double Y, double Y0)
        {
            Label Ctrl = new Label();
            Ctrl.Text = text;
            Ctrl.Font = Func.FontAdjust(form, font + 4, font + 2, font);
            Ctrl.Name = name;
            Ctrl.AutoSize = true;
            Ctrl.ForeColor = color;
            Ctrl.TextAlign = allign;
            Ctrl.Location = Func.PositionAdjust(form, Ctrl, X, Y, Y0);
            form.Controls.Add(Ctrl);
           
        }


        ///
        /// 
        /// 冒泡提示
        ///
        /// System.Windows.Forms的一个控件，在其上面提示显示
        /// 提示的标题默认（温馨提示）
        /// 提示的信息默认（???）
        /// 提示显示等待时间
        /// DevExpress.Utils.ToolTipType 显示的类型
        /// DevExpress.Utils.ToolTipLocation 在控件显示的位置
        /// 是否自动隐藏提示信息
        /// DevExpress.Utils.ToolTipIconType 显示框图表的类型
        /// 一个System.Windows.Forms.ImageList 装载Icon图标的List，显示的ToolTipIconType上，可以为Null
        /// 图标在ImageList上的索引，ImageList为Null时传0进去
        public static void NewToolTip(Control ctl, string title, string content, int showTime,
            System.Windows.Forms.ImageList imgList, int imgIndex)
        {
            DevExpress.Utils.ToolTipController MyToolTipClt = null;
            DevExpress.Utils.ToolTipControllerShowEventArgs args = null;
            try
            {
                MyToolTipClt = new DevExpress.Utils.ToolTipController();
                args = MyToolTipClt.CreateShowArgs();
                content = (string.IsNullOrEmpty(content) ? "???" : content);
                title = string.IsNullOrEmpty(title) ? "温馨提示" : title;
                MyToolTipClt.ImageList = imgList;
                MyToolTipClt.ImageIndex = (imgList == null ? 0 : imgIndex);
                args.AutoHide = true;
                MyToolTipClt.ShowBeak = true;
                MyToolTipClt.ShowShadow = true;
                MyToolTipClt.Rounded = false;
                MyToolTipClt.AutoPopDelay = (showTime == 0 ? 2000 : showTime);
                MyToolTipClt.SetToolTip(ctl, content);
                MyToolTipClt.SetTitle(ctl, title);
                MyToolTipClt.ToolTipType = DevExpress.Utils.ToolTipType.SuperTip;
                MyToolTipClt.SetToolTipIconType(ctl, DevExpress.Utils.ToolTipIconType.None);
                MyToolTipClt.Active = true;
                MyToolTipClt.AppearanceTitle.Font = new System.Drawing.Font("STXihei", 12, FontStyle.Bold);
                MyToolTipClt.Appearance.Font = new System.Drawing.Font("STXihei", 12);
                MyToolTipClt.HideHint();

                //MyToolTipClt.ShowHint(content, title, ctl, tipLocation);
            }
            catch 
            {
                //CommonFunctionHeper.CommonFunctionHeper.CreateLogFiles(ex);
            }

        }

 


        //public static ToolTip NewToolTip(Control parent, string text)
        //{
        //    ToolTip Ctrl = new ToolTip();
        //    // Set up the delays for the ToolTip.
        //    Ctrl.AutoPopDelay = 10 * 1000;
        //    Ctrl.InitialDelay = 0;
        //    Ctrl.OwnerDraw = true;
        //    Ctrl.ReshowDelay = 25;
        //    // Force the ToolTip text to be displayed whether or not the form is active.
        //    Ctrl.ShowAlways = true;
        //    Ctrl.IsBalloon = false;
        //    //Ctrl.Draw +=new DrawToolTipEventHandler(Ctrl_Draw);
        //    Ctrl.SetToolTip(parent, text);
        //    return Ctrl;
        //}

        //public static void Ctrl_Draw(object sender, DrawToolTipEventArgs e)
        //{
        //    e.DrawBackground();
        //    Font f = new System.Drawing.Font("STXihei",12);//, 24, FontStyle.Regular);
        //    e.Graphics.DrawString(e.ToolTipText, f, Brushes.Blue, new PointF(0, 0));
        //}

        public static void NewTextBox(Control form, string name, string text, int font, Color color, HorizontalAlignment allign, double X, double Y, double Y0)
        {
            TextBox Ctrl = new TextBox();
            Ctrl.Text = text;
            Ctrl.Font = Func.FontAdjust(form, font + 4, font + 2, font);
            Ctrl.Name = name;
            Ctrl.AutoSize = true;
            Ctrl.ForeColor = color;
            Ctrl.TextAlign = allign;
            Ctrl.Location = Func.PositionAdjust(form, Ctrl, X, Y, Y0);
            form.Controls.Add(Ctrl);
        }

        public static Aspose.Cells.Style creatStyle(Workbook wb, System.Drawing.Color fontcolor, System.Drawing.Color color, TextAlignmentType HAlign)
        {
            Aspose.Cells.Style txtStyle = wb.CreateStyle();
            txtStyle.Borders[Aspose.Cells.BorderType.LeftBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 左边界线  
            txtStyle.Borders[Aspose.Cells.BorderType.RightBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 右边界线   
            txtStyle.Borders[Aspose.Cells.BorderType.TopBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 上边界线   
            txtStyle.Borders[Aspose.Cells.BorderType.BottomBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 下边界线    
            txtStyle.Font.Name = "STXihei"; //字体 
            txtStyle.Font.Color = fontcolor; //颜色
            txtStyle.HorizontalAlignment = HAlign;
            txtStyle.VerticalAlignment = TextAlignmentType.Center;
            txtStyle.IsTextWrapped = true;
            //txtStyle.Font.IsBold = true; //设置粗体 
            txtStyle.Font.Size = 11; //设置字体大小 

            txtStyle.ForegroundColor = color; //背景色 
            txtStyle.Pattern = BackgroundType.Solid;
            //txtStyle.Pattern = Aspose.Cells.BackgroundType.Solid;
            return txtStyle;
        }
    }
    #endregion
    
    #region HYQFileInfoList
    /// <summary>
    /// Listview调用图标
    /// </summary>
    class HYQFileInfoList
        {
            public List<FileInfoWithIcon> list;
            public ImageList imageListLargeIcon;
            public ImageList imageListSmallIcon;


            /// <summary>
            /// 根据文件路径获取生成文件信息，并提取文件的图标
            /// </summary>
            /// <param name="filespath"></param>
            public HYQFileInfoList(string[] filespath)
            {
                list = new List<FileInfoWithIcon>();
                imageListLargeIcon = new ImageList();
                imageListLargeIcon.ImageSize = new Size(32, 32);
                imageListSmallIcon = new ImageList();
                imageListSmallIcon.ImageSize = new Size(16, 16);
                foreach (string path in filespath)
                {
                    try
                    {
                        FileInfoWithIcon file = new FileInfoWithIcon(path);
                        imageListLargeIcon.Images.Add(file.largeIcon);
                        imageListSmallIcon.Images.Add(file.smallIcon);
                        file.iconIndex = imageListLargeIcon.Images.Count - 1;
                        list.Add(file);
                    }
                    catch { }
                    
                }
            }
        }

        class FileInfoWithIcon
        {
            public FileInfo fileInfo;
            public Icon largeIcon;
            public Icon smallIcon;
            public int iconIndex;
            public FileInfoWithIcon(string path)
            {
                fileInfo = new FileInfo(path);
                largeIcon = GetSystemIcon.GetIconByFileName(path, true);
                if (largeIcon == null)
                    largeIcon = GetSystemIcon.GetIconByFileType(Path.GetExtension(path), true);


                smallIcon = GetSystemIcon.GetIconByFileName(path, false);
                if (smallIcon == null)
                    smallIcon = GetSystemIcon.GetIconByFileType(Path.GetExtension(path), false);
            }
        }

        public static class GetSystemIcon
        {
            /// <summary>
            /// 依据文件名读取图标，若指定文件不存在，则返回空值。  
            /// </summary>
            /// <param name="fileName">文件路径</param>
            /// <param name="isLarge">是否返回大图标</param>
            /// <returns></returns>
            public static Icon GetIconByFileName(string fileName, bool isLarge = true)
            {
                int[] phiconLarge = new int[1];
                int[] phiconSmall = new int[1];
                //文件名 图标索引 
                Win32.ExtractIconEx(fileName, 0, phiconLarge, phiconSmall, 1);
                IntPtr IconHnd = new IntPtr(isLarge ? phiconLarge[0] : phiconSmall[0]);

                if (IconHnd.ToString() == "0")
                    return null;
                return Icon.FromHandle(IconHnd);
            }


            /// <summary>  
            /// 根据文件扩展名（如:.*），返回与之关联的图标。
            /// 若不以"."开头则返回文件夹的图标。  
            /// </summary>  
            /// <param name="fileType">文件扩展名</param>  
            /// <param name="isLarge">是否返回大图标</param>  
            /// <returns></returns>  
            public static Icon GetIconByFileType(string fileType, bool isLarge)
            {
                if (fileType == null || fileType.Equals(string.Empty)) return null;


                RegistryKey regVersion = null;
                string regFileType = null;
                string regIconString = null;
                string systemDirectory = Environment.SystemDirectory + "\\";


                if (fileType[0] == '.')
                {
                    //读系统注册表中文件类型信息  
                    regVersion = Registry.ClassesRoot.OpenSubKey(fileType, false);
                    if (regVersion != null)
                    {
                        regFileType = regVersion.GetValue("") as string;
                        regVersion.Close();
                        regVersion = Registry.ClassesRoot.OpenSubKey(regFileType + @"\DefaultIcon", false);
                        if (regVersion != null)
                        {
                            regIconString = regVersion.GetValue("") as string;
                            regVersion.Close();
                        }
                    }
                    if (regIconString == null)
                    {
                        //没有读取到文件类型注册信息，指定为未知文件类型的图标  
                        regIconString = systemDirectory + "shell32.dll,0";
                    }
                }
                else
                {
                    //直接指定为文件夹图标  
                    regIconString = systemDirectory + "shell32.dll,3";
                }
                string[] fileIcon = regIconString.Split(new char[] { ',' });
                if (fileIcon.Length != 2)
                {
                    //系统注册表中注册的标图不能直接提取，则返回可执行文件的通用图标  
                    fileIcon = new string[] { systemDirectory + "shell32.dll", "2" };
                }
                Icon resultIcon = null;
                try
                {
                    //调用API方法读取图标  
                    int[] phiconLarge = new int[1];
                    int[] phiconSmall = new int[1];
                    uint count = Win32.ExtractIconEx(fileIcon[0], Int32.Parse(fileIcon[1]), phiconLarge, phiconSmall, 1);
                    IntPtr IconHnd = new IntPtr(isLarge ? phiconLarge[0] : phiconSmall[0]);
                    resultIcon = Icon.FromHandle(IconHnd);
                }
                catch { }
                return resultIcon;
            }
        }


        /// <summary>  
        /// 定义调用的API方法  
        /// </summary>  
        class Win32
        {
            [DllImport("shell32.dll")]
            public static extern uint ExtractIconEx(string lpszFile, int nIconIndex, int[] phiconLarge, int[] phiconSmall, uint nIcons);
        }
    #endregion

    #region Func
    public class Func
    {
        #region 基础函数 FontAdjust()：调整字体和字体大小
        public static System.Drawing.Font FontAdjust(Control form, int BigSize, int MediumSize, int SmallSize)
        {
            if (form.Width > 1600)
            {
                return new System.Drawing.Font("STXihei", BigSize, FontStyle.Regular);
            }
            else if (form.Width > 1200)
            {
                return new System.Drawing.Font("STXihei", MediumSize, FontStyle.Regular);
            }
            else
            {
                return new System.Drawing.Font("STXihei", SmallSize, FontStyle.Regular);
            }
        }
        #endregion

        #region 基础函数 PositionAdjust()：调整相对位置
        public static Point PositionAdjust(Control form, Control ctrl, double X, double Y, double Y0)
        {
            Point tmp = new System.Drawing.Point((int)((form.Width) * X), (int)((form.Height - Y0) * Y + Y0));
            return tmp;
        }
        #endregion

        #region 基础函数 Dev过滤
        ////EditValueChanging这个值改变事件
        //private void filter_column_text(object sender, DevExpress.XtraEditors.Controls.ChangingEventArgs e)
        //{
        //    this.gridView1.OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never;
        //    this.gridView1.ActiveFilterCriteria = BulidFilterCriteria();
        //    this.gridView1.RefreshData();
        //}

        ////过滤方法，filter_text.Text是需要过滤的文本内容
        //private GroupOperator BulidFilterCriteria()
        //{
        //    CriteriaOperatorCollection filterCollection = new CriteriaOperatorCollection();
        //    //四列中都参与数据过滤
        //    filterCollection.Add(CriteriaOperator.Parse(string.Format("MTNO LIKE '%{0}%'", filter_text.Text)));
        //    filterCollection.Add(CriteriaOperator.Parse(string.Format("MTNAME LIKE '%{0}%'", filter_text.Text)));
        //    filterCollection.Add(CriteriaOperator.Parse(string.Format("PYCODE LIKE '%{0}%'", filter_text.Text)));
        //    filterCollection.Add(CriteriaOperator.Parse(string.Format("WBCODE LIKE '%{0}%'", filter_text.Text)));
        //    return new GroupOperator(GroupOperatorType.Or, filterCollection);
        //}
        #endregion
 
        #region 基础函数 WidthCal()：计算相对宽度
        public static double WidthCal(Control form, int W)
        {
            double result = (double)W / (double)form.Width;
            return result;
        }
        #endregion

        #region 基础函数 HeightCal()：计算相对高度
        public static double HeightCal(Control form, int H, double Y0)
        {
            double result = (double)H / (double)(form.Height - Y0);
            return result;
        }
        #endregion

        #region 基础函数 FindControl：动态查找控件_Label
        public static Control FindControl(Control parentControl, string findCtrlName)
        {
            Control _findedControl = null;
            if (!string.IsNullOrEmpty(findCtrlName) && parentControl != null)
            {
                foreach (Control ctrl in parentControl.Controls)
                {
                    if (ctrl.Name.Equals(findCtrlName))
                    {
                        _findedControl = ctrl;
                        break;
                    }
                }
            }
            return _findedControl;
        }
        #endregion

        #region 基础函数 CalItemsIndex：计算最后一个Item的Index
        public static int CalItemsIndex(string[] Items, int index)
        {
            int ind;
            try
            {
                if (Items.Length > 0)
                {
                    string substr = CalItemsHead(Items, index);
                    int pos = substr.IndexOf("_");
                    string ind_s = substr.Substring(pos + 1);
                    int.TryParse(ind_s, out ind);
                    //ind++;
                }
                else
                {
                    ind = 0;
                }
            }
            catch
            {
                ind = 0;
            }
            return ind;
        }
        #endregion

        #region 基础函数 CalItemsHead：计算=前面的Item值
        public static string CalItemsHead(string[] Items, int index)
        {
            string substr = "";
            try
            {
                string lastItems = Items[index];
                int k = lastItems.IndexOf("=");
                substr = lastItems.Substring(0, k);
                return substr;
            }
            catch
            {
                return substr;
            }

        }
        #endregion

        #region 基础函数 ConvertLine2List:将行数转换成列表
        public static List<string> ConvertLine2List(string content)
        {
            string[] sArry;
            List<string> slist = new List<string>();
            sArry = Func.ConvertItems(content, new string[] { Environment.NewLine });

            ///转换为List
            for (int j = 0; j < sArry.Length; j++)
            {
                try
                {
                    if ((slist[j].ToString() == null) || (slist[j].ToString() == ""))
                    {
                        slist.RemoveAt(j);
                    }
                    else
                    {
                        slist.Add(sArry[j]);
                    }
                }
                catch 
                {
                    slist.Add(sArry[j]);
                }
                    
                //result + sArry[j] + "\n";
            }
            return slist;
        }
        #endregion

        #region 基础函数：寻找点击项对应的Item
        /// <summary>
        /// 寻找点击项对应的Item
        /// </summary>
        /// <param name="tmp">被搜索变量</param>
        /// <param name="Item">存储的Openlist Item</param>
        public static int FindItemInIni(string Text_No, string[] Items)
        {
            int ind, ind1,result = -1;
            try
            {
                int.TryParse(Text_No, out ind1);
                for (int i = 0; i < Items.Length; i++)
                {
                    ind = Func.CalItemsIndex(Items, i);
                    if (ind == ind1)
                    {
                        result = i;
                    }
                    else
                    { }
                }
            }
            catch
            {
                result = -1;
            }
            return result;
        }
        #endregion

        #region 基础函数：删除文件夹及其内容
        /// <summary>
        /// 清空指定的文件夹，但不删除文件夹
        /// </summary>
        /// <param name="dir"></param>
        public static void ClearFolder(string dir)
        {
            foreach (string d in Directory.GetFileSystemEntries(dir))
            {
                if (File.Exists(d))
                {
                    try
                    {
                        FileInfo fi = new FileInfo(d);
                        if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                            fi.Attributes = FileAttributes.Normal;
                        File.Delete(d);//直接删除其中的文件 
                    }
                    catch
                    {

                    }
                }
                else
                {
                    try
                    {
                        DirectoryInfo d1 = new DirectoryInfo(d);
                        if (d1.GetFiles().Length != 0)
                        {
                            ClearFolder(d1.FullName);////递归删除子文件夹
                        }
                        Directory.Delete(d);
                    }
                    catch
                    {

                    }
                }
            }
        }
        /// <summary>
        /// 删除文件夹及其内容
        /// </summary>
        /// <param name="dir"></param>
        public static void DeleteFolder(string dir)
        {
            foreach (string d in Directory.GetFileSystemEntries(dir))
            {
                if (File.Exists(d))
                {
                    FileInfo fi = new FileInfo(d);
                    if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                        fi.Attributes = FileAttributes.Normal;
                    File.Delete(d);//直接删除其中的文件 
                }
                else
                    ClearFolder(d);////递归删除子文件夹
                Directory.Delete(d);
            }
        }
        #endregion

        #region 基础函数：将拖动的文件添加入ListView
        public static void ListViewShowIcon(DevExpress.XtraEditors.ImageListBoxControl lsv, string path)
        {
            string[] filespath;
            try
            {
                filespath = Directory.GetFiles(path);
            }
            catch
            {
                return;
            }
            HYQFileInfoList fileList = new HYQFileInfoList(filespath);
            lsv.BeginUpdate();
            lsv.Items.Clear();
            int i = 0;
            foreach (FileInfoWithIcon file in fileList.list)
            {
                lsv.Items.Add(file.fileInfo.Name);//.Split('.')[0]);
                lsv.Items[i].ImageIndex = i;
                i++;
            }
            if (i <= 4)
            {
                lsv.ImageList = fileList.imageListLargeIcon;//imageListLargeIcon;
            }
            else
            {
                lsv.ImageList = fileList.imageListSmallIcon;//imageListLargeIcon;
            }
            //lsv.ImageList = fileList.imageListSmallIcon;//imageListLargeIcon;
            lsv.EndUpdate();
            //lsv.BeginUpdate();
        }


        public static byte[] ImgToByte(Image img)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    byte[] imagedata = null;
                    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                    imagedata = ms.GetBuffer();
                    return imagedata;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString());
                return null;
            }
        }

        #endregion

        #region 基础函数 DeleteTail
        /// <summary>
        /// 从Str删除最后一个subStr
        /// </summary>
        /// <param name="Str">完整字符</param>
        /// <param name="subStr">去尾字符</param>
        /// <returns></returns>
        public static string DeleteTail(string Str, string subStr)
        {
            try
            {
                return Str.Substring(0, Str.Length - subStr.Length);
            }
            catch
            {
                return null;
            }
                
        }
        #endregion

        #region 基础函数 CalItemsTail：计算=后面的Item值
        public static string CalItemsTail(string[] Items, int index)
        {
            string lastItems;
            string substr;
            try
            {
                if (index < Items.Length)
                {
                    lastItems = Items[index];
                }
                else
                {
                    lastItems = Items[Items.Length - 1];

                }
                int k = lastItems.IndexOf("=");
                return substr = lastItems.Substring(k + 1);
            }
            catch 
            {
                return substr = null;
            }
                
        }
        #endregion

        #region 基础函数 ConvertItems：计算=转换的Item值
        public static string[] ConvertItems(string Items, string[] str)
        {
            string[] sArray; 
            try
            {
                sArray = Items.Split(str, StringSplitOptions.None);
            }
            catch
            {
                sArray = new string[1] { "null" };
            }
                
            //Regex.Split(Items, "<SPLIT>", RegexOptions.IgnoreCase);
            return sArray;
        }
        #endregion

        #region 界面函数 通用鼠标悬浮、离开
        //图片离开
        public static void General_MouseLeave(object sender, EventArgs e)
        {
            PictureBox var = new PictureBox(); ;
            if (sender.GetType().Equals(typeof(PictureBox)))
            {
                var = sender as PictureBox;
            }
            else
            {
                return;
            }
            try
            { 
                /// 0 为默认 1为离开 2为悬挂
                if (((int)var.Tag) != 1)
                {
                    var.Location = new Point(var.Location.X,var.Location.Y + 2);
                    var.Tag = 1;
                }
            }
            catch
            {
                var.Tag = 1;
            }
                
        }

        //图片移动
        public static void General_MouseMove(object sender, MouseEventArgs e)
        {
            PictureBox var = new PictureBox(); ;
            if (sender.GetType().Equals(typeof(PictureBox)))
            {
                var = sender as PictureBox;
            }
            else
            {
                return;
            }
            try
            { 
                /// 0 为默认 1为离开 2为悬挂
                if (((int)var.Tag) != 2)
                {
                    var.Location = new Point(var.Location.X,var.Location.Y - 2);
                    var.Tag = 2;
                }
            }
            catch
            {
                var.Tag = 2;
            }
        }
        #endregion

        #region 界面函数 图片离开
        public static void TrackSelectedIndex(object sender, EventArgs e)
        {
            //try
            //{
            //    if (TxtB_Track.SelectedIndex != -1)
            //    {
            //        content = sTrack[TxtB_Track.SelectedIndex].ToString().Replace("<ENTER>", "\n");
            //    }
            //    else
            //    {
            //        content = sTrack[0].ToString().Replace("<ENTER>", "\n");
            //    }
            //    content = Func.ConvertItems(content, new string[] { "<DATESPACE_R>" })[1];
            //}
            //catch { }
            //ctrl.Text = content;
        }
        #endregion
    }
    #endregion

    #region HYQUndoStack

    interface IUndoableOperate
    {
        string Redo();
        string Undo();
    }
    /// <summary>
    /// 撤销重复操作管理器
    /// </summary>
    class HYQUndoStack
    {
        /// <summary>
        /// 撤销栈
        /// </summary>
        Stack<IUndoableOperate> un_stack = new Stack<IUndoableOperate>();
        /// <summary>
        /// 重复栈
        /// </summary>
        Stack<IUndoableOperate> re_stack = new Stack<IUndoableOperate>();


        public void ClearStack()
        {
            un_stack.Clear();
            re_stack.Clear();
        }

        /// <summary>
        /// 获取一个值，指示是否有可撤销的操作
        /// </summary>
        public bool CanUndo()
        {
            return un_stack.Count != 0;
        }

        /// <summary>
        /// 获取一个值，指示是否有可重复的操作
        /// </summary>
        public bool CanRedo()
        {
            return re_stack.Count != 0;

        }

        /// <summary>
        /// 撤销上一操作
        /// </summary>
        public string Undo()
        {
            if (this.CanUndo())
            {
                IUndoableOperate op = un_stack.Pop();
                string str = op.Undo();
                re_stack.Push(op);
                return str;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 重复被撤销的操作
        /// </summary>
        public string Redo()
        {
            if (this.CanRedo())
            {
                IUndoableOperate op = re_stack.Pop();
                string str = op.Redo();
                un_stack.Push(op);
                return str;
            }
            else
            {
                return null;
            }
        }


        /// <summary>
        /// 将某一操作存放到撤销栈中
        /// </summary>
        /// <param name="op"></param>
        public void PushToUndoStack(IUndoableOperate op)
        {
            un_stack.Push(op);
            re_stack.Clear();
        }
    }
    #endregion

    #region HYQAutoComplete
    /// <summary>
/// 自动填充
/// </summary>
    public class HYQAutoComplete
    {
        List<TextBox> _CompleteObjectList = new List<TextBox>();
        Dictionary<string, AutoCompleteStringCollection> _Source = new Dictionary<string, AutoCompleteStringCollection>();
        //SqlConnection conn = new SqlConnection("Data Source=.;Initial Catalog=TestDB;Integrated Security=True");
        public HYQAutoComplete()
        {
            //conn.Open();
            //SqlCommand cmd = new SqlCommand("select * from AutoComplete", conn);
            //SqlDataReader read = cmd.ExecuteReader();
            //while (read.Read())
            //{
            //    string key = read["name"].ToString();
            //    if (!_Source.ContainsKey(key))
            //        _Source.Add(key, new AutoCompleteStringCollection());
            //    _Source[key].Add(read["str"].ToString());
            //}
            //read.Close();
            //conn.Close();
                
        }

        public static void AddAll(Control item, Dictionary<string, AutoCompleteStringCollection> _Source)
        {
            for (int i = 0; i < item.Controls.Count; i++)
            {
                Control var = item.Controls[i];
                if (var.GetType().Equals(typeof(TextBox)))
                {
                    Add(var as DevExpress.XtraEditors.TextEdit, _Source);
                }
            }
        }
        public static void Add(DevExpress.XtraEditors.TextEdit text, Dictionary<string, AutoCompleteStringCollection> _Source)
        {
            //_CompleteObjectList.Add(text);
            text.Tag = _Source;
            text.Leave += new EventHandler(text_Leave);
            //text.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            //text.AutoCompleteSource = AutoCompleteSource.CustomSource;
            if (!_Source.ContainsKey(text.Name))
            {
                _Source.Add(text.Name, new AutoCompleteStringCollection());
            }
            //text.AutoCompleteCustomSource = _Source[text.Name];
        }

        public void Delete(TextBox text)
        {
            _CompleteObjectList.Remove(text);
        }

        public void DeleteAll(Control item)
        {
            for (int i = 0; i < item.Controls.Count; i++)
            {
                Control var = item.Controls[i];
                if (var.GetType().Equals(typeof(TextBox)))
                {
                    Delete(var as TextBox);
                }
            }
        }

        public void AutoCompleteClear()
        {
            foreach (AutoCompleteStringCollection var in _Source.Values)
            {
                var.Clear();
            }
            //SqlCommand cmd = new SqlCommand("Delete AutoComplete", conn);
            //conn.Open();
            //cmd.ExecuteNonQuery();
            //conn.Close();
        }

        public static void text_Leave(object sender, EventArgs e)
        {
            TextBox text = sender as TextBox;
            try { 
                Dictionary<string, AutoCompleteStringCollection> _Source = text.Tag as Dictionary<string, AutoCompleteStringCollection>;
                if (text.Text == "")
                    return;
                string key = text.Name;
                if (!_Source.ContainsKey(key))
                {
                    _Source.Add(key, new AutoCompleteStringCollection());
                }
                if (!_Source[key].Contains(text.Text))
                {
                    //SqlCommand cmd = new SqlCommand("insert into AutoComplete select '" + key.Replace("'", "''") + "', '" + text.Text.Replace("'", "''") + "'", conn);
                    //conn.Open();
                    //cmd.ExecuteNonQuery();
                    _Source[key].Add(text.Text);
                    //conn.Close();
                }
            }catch{}

                
        }
    }
#endregion
    
    #region HYQStringHelper
    /// <summary>
    /// 读取首字母
    /// </summary>
    public class HYQStringHelper  
    {  
        #region " LetterItem "  
        private class LetterItem  
        {  
            private String fLetter;  
            private Int64 fMinValue;  
            private Int64 fMaxValue;  
            public String Letter { get { return fLetter; } }  
            public Int64 MinValue { get { return fMinValue; } }  
            public Int64 MaxValue { get { return fMaxValue; } }  
            public LetterItem(String fLetter, Int64 fMinValue, Int64 fMaxValue)  
            {  
                this.fLetter = fLetter;  
                this.fMinValue = fMinValue;  
                this.fMaxValue = fMaxValue;  
            }  
        }  
        #endregion  

        /// <summary>  
        /// 获取一段中文中每个中文拼音的第一个字母  
        /// String.Format("{0}:{1}", "王子涵", HYQStringHelper.GetFirstLetterOfChinese("王子涵", false));
        /// </summary>  
        /// <param name="fInputChinese">需要获取字母的中文</param>  
        /// <returns>中文拼音的第一个字母</returns>  
        public static string GetFirstLetterOfChinese(string fInputChinese)  
        {  
            return GetFirstLetterOfChinese(fInputChinese, false);  
        }  
        /// <summary>  
        /// 获取一段中文中每个中文拼音的第一个字母  
        /// </summary>  
        /// <param name="fInputChinese">需要获取字母的中文</param>  
        /// <param name="fReutrnEmptyWhenFailure">当输入不是中文时是否返回空值。True:返回空值；False：返回传入参数的大写</param>  
        /// <returns>中文拼音的第一个字母</returns>  
        public static string GetFirstLetterOfChinese(string fInputChinese, Boolean fReutrnEmptyWhenFailure)  
        {  
            string letters = "";  
              
            foreach (char c in fInputChinese.ToCharArray())  
                letters += GetFirstLetterOfPinyin(c.ToString(), fReutrnEmptyWhenFailure);  
            return letters;  
        }  
        /// <summary>  
        /// 获取一个中文拼音的第一个字母。  
        /// </summary>  
        /// <param name="fInputSingleChinese">需要获取字母的一个中文</param>  
        /// <param name="fReutrnEmptyWhenFailure">当输入不是中文时是否返回空值。True:返回空值；False：返回传入参数的大写</param>  
        /// <returns>中文拼音的第一个字母</returns>  
        private static string GetFirstLetterOfPinyin(String fInputSingleChinese, Boolean fReutrnEmptyWhenFailure)  
        {  
            byte[] byteArray = System.Text.Encoding.Default.GetBytes(fInputSingleChinese);  
            //如果是字母，则直接返回   
            //建立dictionaryLetter
            List<LetterItem> dictionaryLetter = new List<LetterItem>();
            // 没有 U、V   
            dictionaryLetter.Add(new LetterItem("A", 45217, 45252));
            dictionaryLetter.Add(new LetterItem("B", 45253, 45760));
            dictionaryLetter.Add(new LetterItem("C", 45761, 46317));
            dictionaryLetter.Add(new LetterItem("D", 46318, 46825));
            dictionaryLetter.Add(new LetterItem("E", 46826, 47009));
            dictionaryLetter.Add(new LetterItem("F", 47010, 47296));
            dictionaryLetter.Add(new LetterItem("G", 47297, 47613));
            dictionaryLetter.Add(new LetterItem("H", 47614, 48118));
            dictionaryLetter.Add(new LetterItem("J", 48119, 49061));
            dictionaryLetter.Add(new LetterItem("K", 49062, 49323));
            dictionaryLetter.Add(new LetterItem("L", 49324, 49895));
            dictionaryLetter.Add(new LetterItem("M", 49896, 50370));
            dictionaryLetter.Add(new LetterItem("N", 50371, 50613));
            dictionaryLetter.Add(new LetterItem("O", 50614, 50621));
            dictionaryLetter.Add(new LetterItem("P", 50622, 50905));
            dictionaryLetter.Add(new LetterItem("Q", 50906, 51386));
            dictionaryLetter.Add(new LetterItem("R", 51387, 51445));
            dictionaryLetter.Add(new LetterItem("S", 51446, 52217));
            dictionaryLetter.Add(new LetterItem("T", 52218, 52697));
            dictionaryLetter.Add(new LetterItem("W", 52698, 52979));
            dictionaryLetter.Add(new LetterItem("X", 52980, 53640));
            dictionaryLetter.Add(new LetterItem("Y", 53689, 54480));
            dictionaryLetter.Add(new LetterItem("Z", 54481, 55289));  
            if (byteArray.Length == 1)  
            {  
                return fReutrnEmptyWhenFailure  
                    ? fInputSingleChinese.ToUpper()   
                    : String.Empty;  
            }  
            // 获取范围  
            short minValue = (short)(byteArray[0]);  
            short maxValue = (short)(byteArray[1]);  
            Int64 value = minValue * 256 + maxValue;  
            foreach (LetterItem letterItem in dictionaryLetter)  
            {  
                if (value >= letterItem.MinValue &&  
                    value <= letterItem.MaxValue)  
                    return letterItem.Letter;  
            }  
            return "?"; // 未知  
        }
    }
    #endregion

    #region HYQReNameHelper
    /// <summary>
    /// 文件已存在，重命名操作类
    /// </summary>
    public class HYQReNameHelper
    {
        /// <summary>
        /// 对文件进行重命名
        /// </summary>
        /// <param name="strFilePath"></param>
        /// <returns></returns>
        public static string FileReName(string strFilePath)
        {
            //判断该文件是否存在，存在则返回新名字，否则返回原来的名
            if (!File.Exists(strFilePath))
            {
                return Path.GetFileName(strFilePath);
            }
            else
            {
                //获取不带扩展名的文件名称
                string strFileNameWithoutExtension = Path.GetFileNameWithoutExtension(strFilePath);
                //获取扩展名
                string strFileExtension = Path.GetExtension(strFilePath);
                //获取目录名
                string strDirPath = Path.GetDirectoryName(strFilePath);
                //以文件名开头和结尾的正则
                string strRegex = "^" + strFileNameWithoutExtension + "(\\d+)?";
                Regex regex = new Regex(strRegex);
                //获取该路径下类似的文件名
                string[] strFilePaths = Directory.GetFiles(strDirPath, "*" + strFileExtension).Where(path => regex.IsMatch(Path.GetFileNameWithoutExtension(path))).ToArray();
                //获得新的文件名
                return strFileNameWithoutExtension + "(" + (strFilePaths.Length + 1).ToString() + ")" + strFileExtension;
            }
        }
        /// <summary>
        /// 文件夹已存在，重命名
        /// </summary>
        /// <param name="strFolderPath"></param>
        /// <returns></returns>
        public static string FolderReName(string strFolderPath)
        {
            //判断该文件夹是否存在，存在则返回新名字，否则返回原来的名
            if (!Directory.Exists(strFolderPath))
            {
                return Path.GetFileName(strFolderPath);
            }
            else
            {
                //获取文件夹名
                string strFolderName = Path.GetFileName(strFolderPath);
                //获取目录名
                string strDirPath = Path.GetDirectoryName(strFolderPath);
                //以文件夹名开头和结尾的正则
                string strRegex = "^" + strFolderName + "(\\d+)?";
                Regex regex = new Regex(strRegex);
                //获取该路径下类似的文件夹名
                string[] strFilePaths = Directory.GetDirectories(strDirPath).Where(path => regex.IsMatch(Path.GetFileName(path))).ToArray();
                //获得新的文件名
                return strFolderName + "(" + (strFilePaths.Length + 1).ToString() + ")";
            }
        }
    }
    #endregion

}
