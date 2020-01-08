using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using System.IO;
using System.Net.Mail;
using System.Net;
using System.Text.RegularExpressions;
using System.Collections;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using Aspose.Cells;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraBars.Ribbon;
using DevExpress.Images;
using DevExpress.Data.Filtering;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Diagnostics;
using System.Reflection;

namespace OPL
{
    public partial class PowerMainForm : DevExpress.XtraEditors.XtraForm//DevExpress.XtraBars.Ribbon.RibbonForm
    {
        protected static DevExpress.LookAndFeel.DefaultLookAndFeel defaultLookAndFeel = new DevExpress.LookAndFeel.DefaultLookAndFeel();

        public PowerMainForm()
        {
            InitializeComponent();
        }

        #region 全局变量定义
        public static PowerMainForm MainForm = null;
        public static ToBeType SubForm1 = null;
        public string saveFilePath = @".\Save\";
        public static string saveFilePath_N = "日常记录";
        public static string saveFilePath_D = @".\Save\日常记录.ini";
        public static string saveFilePath_M = @".\User\Main.ini";
        DevExpress.XtraGrid.GridControl mainGrid;
        DevExpress.XtraGrid.Views.Grid.GridView mainGridView;
        List<Control> ls_Edit = new List<Control>();    //全部可见控件集合
        List<Control> lsCtrl_Visible = new List<Control>();    //全部可见控件集合
        List<Control> lsCtrl_AutoSave = new List<Control>();   //自动保存控件集合
        List<Control> lsCtrl_BackGroud = new List<Control>();   //背景控件集合
        List<string> lb_HeaderName = new List<string>();   //表头集合
        List<string> lb_HeaderName2 = new List<string>();   //表头集合
        List<string> copyStack = new List<string>();   //表头集合
        HYQUndoStack RecallStack = new HYQUndoStack();
        bool beginMove = false;
        bool newTrackFlag = false;
        bool IniFlag = true;
        bool recallFlag = false;
        bool sort1Flag = false;
        int UpdateTimer_Ticks;
        int curX;
        int curY;
        int winformSize = 1;
        int curSelectionOPL = -1;
        int curSelectionTRK = -1;
        //用于记录，鼠标是否已按下
        bool isMouseDown = false;
        //用于鼠标拖动多选，标记是否记录开始行
        bool isSetStartRow = false;
        //用于鼠标拖动多选，记录开始行
        private int StartRowHandle = -1;
        //用于鼠标拖动多选，记录现在行
        private int CurrentRowHandle = -1;

        public class itemRecall : IUndoableOperate
        {
            private string local_oldVal;
            private string loca_newVal;
            private string local_key;
            /// <summary>
            /// 执行olditem与newitem的保存
            /// </summary>
            /// <param name="oldval">执行保存前的item</param>
            /// <param name="newval">执行保存后的item</param>
            public void saveRecall(string oldValue, string newValue, string key)
            {
                local_oldVal = oldValue;
                loca_newVal = newValue;
                local_key = key;
            }

            public string Undo()
            {
                return local_key + "=" + local_oldVal;
            }

            public string Redo()
            {
                return local_key + "=" + loca_newVal;
            }
        }


        #region 快捷键定义
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterHotKey(
        IntPtr hWnd,             //要定义热键的窗口的句柄   
        int id,                        //定义热键ID（不能与其它ID重复）  
        KeyModifiers fsModifiers,         //标识热键是否在按Alt、Ctrl、Shift、Windows等键时才会生效 
        Keys vk                   //定义热键的内容   
        );
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnregisterHotKey(
        IntPtr hWnd,             //要取消热键的窗口的句柄   
        int id                        //要取消热键的ID  
        );
        public enum KeyModifiers
        {
            None = 0,
            Alt = 1,
            Ctrl = 2,
            Shift = 4,
            WindowsKey = 8,
            CtrlAndShift = 6
        }
        #endregion
        #endregion

        #region 全局配置字
        /// indicate wether to show foolish Lion head or not
        /// 1 for show
        /// 0 for not show
        private int showLionHeadOrNot = 1;

        #endregion

        #region 主函数 Load
        private void PowerMainForm_Load(object sender, EventArgs e)
        {
            #region StartDebug
            //Initializing
            Stopwatch sw1 = new Stopwatch();
            Stopwatch sw2 = new Stopwatch();
            sw1.Start();//主程序开始计时
            #endregion

            #region 注册按键
            //注册热键Ctrl + Up，Id号为100。
            RegisterHotKey(Handle, 100, KeyModifiers.Ctrl, Keys.Up);
            //注册热键Ctrl + Down，Id号为101。  
            RegisterHotKey(Handle, 101, KeyModifiers.Ctrl, Keys.Down);
            //注册热键Ctrl + Left，Id号为102。
            RegisterHotKey(Handle, 102, KeyModifiers.Ctrl, Keys.Left);
            //注册热键Ctrl + Right，Id号为103。  
            RegisterHotKey(Handle, 103, KeyModifiers.Ctrl, Keys.Right);
            //注册热键Ctrl + Z，Id号为104。  
            RegisterHotKey(Handle, 104, KeyModifiers.Ctrl, Keys.F12); //reserved
            #endregion

            ///主程序
            InitialAtOpening(false);

            IniFlag = false;

            #region StartDebug
            sw1.Stop();
            //sw1.ElapsedMilliseconds即为总耗时（毫秒），计时器运用
            sw1.ToString();
            #endregion
        }
        #endregion

        #region 主函数 Close
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            #region 注销按键
            //注销Id号为100的热键设定   
            UnregisterHotKey(Handle, 100);
            //注销Id号为101的热键设定   
            UnregisterHotKey(Handle, 101);
            //注销Id号为102的热键设定   
            UnregisterHotKey(Handle, 102);
            //注销Id号为103的热键设定   
            UnregisterHotKey(Handle, 103);
            //注销Id号为104的热键设定   
            UnregisterHotKey(Handle, 104);
            #endregion
        }
        #endregion

        #region -1界面函数 快捷键定义
        ///ref功能:
        ///ref 关键字使参数按引用传递。
        ///其效果是，当控制权传递回调用方法时，在
        ///方法中对参数所做的任何更改都将反映在该变量中。
        ///简单点说就是,使用了ref和out的效果就几乎和C中使用了指针变量一样。
        ///它能够让你直接对原数进行操作，而不是对那个原数的Copy进行操作。
        protected override void WndProc(ref Message m)
        {
            const int WM_HOTKEY = 0x0312;            
            const int WM_NCLBUTTONDBLCLK = 0xA3;
            const int WM_NCLBUTTONDOWN = 0x00A1;
            const int HTCAPTION = 2;
            int tmp = curSelectionOPL;
            //按快捷键   
            switch (m.Msg)
            {
                case WM_HOTKEY:
                    switch (m.WParam.ToInt32())
                    {
                        case 100: //按下的是Ctrl + ↑  
                            try
                            {
                                
                                this.mainGridView.MovePrev();
                                OpenList_SelectionChanged(null, null);
                            }
                            catch { }
                            break;
                        case 101: //按下的是Ctrl + ↓   
                            try
                            {
                                this.mainGridView.MoveNext();
                                OpenList_SelectionChanged(null, null);
                            }
                            catch { }
                            break;
                        case 102: //按下的是Ctrl + ←   
                            try
                            {
                                TxtB_Track.SelectedIndex = TxtB_Track.SelectedIndex - 1;
                            }
                            catch { }
                            break;
                        case 103: //按下的是Ctrl + →   
                            try
                            {
                                TxtB_Track.SelectedIndex = TxtB_Track.SelectedIndex + 1;
                            }
                            catch { }
                            break;
                        case 104: 
                            
                            break;
                    }
                    break;
            }

            if (m.Msg == WM_NCLBUTTONDOWN && m.WParam.ToInt32() == HTCAPTION)
            {
                if (this.WindowState == FormWindowState.Maximized)
                {
                    this.SuspendLayout();
                    this.WindowState = FormWindowState.Normal;
                    int X = this.Location.X;
                    this.Location = new Point(0, 0);
                    this.ResumeLayout();
                }
            }

            if (m.Msg == WM_NCLBUTTONDBLCLK)
            {
                if((this.WindowState == FormWindowState.Normal)&&(this.Location.X == 0))return;
            }

            base.WndProc(ref m);
        }
        #endregion

        #region 0界面函数 InitialAtOpening()：全部初始化
        /// <summary>
        /// 全部初始化
        /// </summary>
        /// <param name="isMiniSize">是否是Mini尺寸，如果是，则offset为0</param>
        private void InitialAtOpening(bool isMiniSize)
        {
            ///赋初始值
            this.Width = 1760;
            this.Height = 990;
            this.Left = 20;
            this.Top = 20;
            MainForm = this;
            IniFlag = true;
            mainGrid = this.OpenList;
            mainGridView = this.OpenListView;
            SubForm1 = new ToBeType();
            PowerMainForm.defaultLookAndFeel.LookAndFeel.SkinName = "Office 2013";
            try
            {
                Directory.CreateDirectory(".\\Data\\");
                Directory.CreateDirectory(".\\User\\");
                Directory.CreateDirectory(".\\Save\\");
            }
            catch { };

            saveFilePath_N = HyqIni.GetINI("Main","saveFilePath_N","日常记录",saveFilePath_M);
            saveFilePath_D = HyqIni.GetINI("Main","saveFilePath_D",@".\Save\日常记录.ini",saveFilePath_M);
            curSelectionOPL = HyqIni.GetINI("Main", "curSelectionOPL", -1, saveFilePath_M);
            sort1Flag = false;

            #region 增加Visbile控件集合
            /*所有输入框控件*/
            ls_Edit.Add(TxtB_No);      ///1
            ls_Edit.Add(TxtB_Date);    ///2
            ls_Edit.Add(TxtB_Source);  ///3
            ls_Edit.Add(TxtB_Type);    ///4
            ls_Edit.Add(TxtB_Descrip); ///5
            ls_Edit.Add(TxtB_Due);     ///6
            ls_Edit.Add(TxtB_Plan);    ///7
            ls_Edit.Add(TxtB_EDate);   ///8
            ls_Edit.Add(TxtB_EDate2);  ///9
            ls_Edit.Add(TxtB_Status);  ///10
            ls_Edit.Add(TxtB_ADate);   ///11
            //ls_Edit.Add(TxtB_Track);   ///12
            ls_Edit.Add(TxtB_File);    ///13
            //ls_Edit.Add(TxtB_FileList);///14
            //ls_Edit.Add(OpenList);     ///15
            //ls_Edit.Add(ribbon);       ///16
            //ls_Edit.Add(PanelA);       ///17
            //ls_Edit.Add(PanelB);       ///18
            //ls_Edit.Add(PanelC);       ///19

            /*所有可见控件*/
            lsCtrl_Visible.Add(TxtB_No);      ///1
            //lsCtrl_Visible.Add(TxtB_Date);    ///2
            lsCtrl_Visible.Add(TxtB_Source);  ///3
            lsCtrl_Visible.Add(TxtB_Type);    ///4
            lsCtrl_Visible.Add(TxtB_Descrip); ///5
            lsCtrl_Visible.Add(TxtB_Due);     ///6
            lsCtrl_Visible.Add(TxtB_Plan);    ///7
            lsCtrl_Visible.Add(MC_EDate);   ///8
            lsCtrl_Visible.Add(MC_EDate2);  ///9
            //lsCtrl_Visible.Add(TxtB_Status);  ///10
            //lsCtrl_Visible.Add(TxtB_ADate);   ///11
            lsCtrl_Visible.Add(TxtB_Track);   ///12
            //lsCtrl_Visible.Add(TxtB_File);    ///13
            lsCtrl_Visible.Add(TxtB_FileList);///14
            lsCtrl_Visible.Add(OpenList);     ///15
            lsCtrl_Visible.Add(ribbon);       ///16
            lsCtrl_Visible.Add(PanelA);       ///17
            lsCtrl_Visible.Add(PanelB);       ///18
            lsCtrl_Visible.Add(PanelC);       ///19

            /*所有自动保存控件*/
            lsCtrl_AutoSave.Add(TxtB_No);      ///1
            lsCtrl_AutoSave.Add(TxtB_Date);    ///2
            lsCtrl_AutoSave.Add(TxtB_Source);  ///3
            lsCtrl_AutoSave.Add(TxtB_Type);    ///4
            lsCtrl_AutoSave.Add(TxtB_Descrip); ///5
            lsCtrl_AutoSave.Add(TxtB_Due);     ///6
            lsCtrl_AutoSave.Add(TxtB_Plan);    ///7
            lsCtrl_AutoSave.Add(TxtB_EDate);   ///8
            lsCtrl_AutoSave.Add(TxtB_EDate2);  ///9
            lsCtrl_AutoSave.Add(TxtB_Status);  ///10            
            lsCtrl_AutoSave.Add(TxtB_ADate);   ///11
            //lsCtrl_AutoSave.Add(TxtB_Track);   ///12
            ///lsCtrl_AutoSave.Add(TxtB_File);    ///13
            //lsCtrl_AutoSave.Add(TxtB_FileList);///14
            //lsCtrl_AutoSave.Add(OpenList);     ///15
            //lsCtrl_AutoSave.Add(ribbon);       ///16
            //lsCtrl_AutoSave.Add(PanelA);       ///17
            //lsCtrl_AutoSave.Add(PanelB);       ///18
            //lsCtrl_AutoSave.Add(PanelC);       ///19

            /*所有背景控件*/
            //lsCtrl_BackGroud.Add(ListPanel);      ///0
            //lsCtrl_BackGroud.Add(ribbon);      ///1
            lsCtrl_BackGroud.Add(PanelA);      ///2
            lsCtrl_BackGroud.Add(PanelB);      ///3
            lsCtrl_BackGroud.Add(PanelC);      ///4
            //lsCtrl_AutoSave.Add(ribbonStatusBar);      ///5

            /*表头集合*/
            lb_HeaderName.Add("序号" + Environment.NewLine + "No");            ///1
            lb_HeaderName.Add("创建日期" + Environment.NewLine + "Date");        ///2
            lb_HeaderName.Add("来源" + Environment.NewLine + "From");            ///3
            lb_HeaderName.Add("类型" + Environment.NewLine + "Type");            ///4
            lb_HeaderName.Add("问题描述" + Environment.NewLine + "Descrip");        ///5
            lb_HeaderName.Add("负责人" + Environment.NewLine + "Resp.");          ///6
            lb_HeaderName.Add("行动措施" + Environment.NewLine + "Action");        ///7
            lb_HeaderName.Add("原定计划" + Environment.NewLine + "Orig. Due");    ///8
            lb_HeaderName.Add("现定计划" + Environment.NewLine + "Cur. Due");    ///9
            lb_HeaderName.Add("状态" + Environment.NewLine + "Sta.");            ///10   
            lb_HeaderName.Add("到期" + Environment.NewLine + "Rem");            ///11  
            lb_HeaderName.Add("完成日期" + Environment.NewLine + "Close Date");            ///12

            /*表头集合*/
            lb_HeaderName2.Add("序号" + Environment.NewLine + "No");            ///1
            lb_HeaderName2.Add("创建日期" + Environment.NewLine + "Date");        ///2
            lb_HeaderName2.Add("来源" + Environment.NewLine + "From");            ///3
            lb_HeaderName2.Add("类型" + Environment.NewLine + "Type");            ///4
            lb_HeaderName2.Add("问题描述" + Environment.NewLine + "Descrip");        ///5
            lb_HeaderName2.Add("负责人" + Environment.NewLine + "Resp.");          ///6
            lb_HeaderName2.Add("行动措施" + Environment.NewLine + "Action");        ///7
            lb_HeaderName2.Add("状态" + Environment.NewLine + "Sta.");            ///10 
            lb_HeaderName2.Add("原定计划" + Environment.NewLine + "Orig. Due");    ///8
            lb_HeaderName2.Add("现定计划" + Environment.NewLine + "Cur. Due");    ///9
            lb_HeaderName2.Add("完成日期" + Environment.NewLine + "Close Date");            ///11
            #endregion

            ///设置双缓冲
            this.DoubleBuffered = true;//设置本窗体
            SetStyle(ControlStyles.UserPaint, true);
            SetStyle(ControlStyles.AllPaintingInWmPaint, true); // 禁止擦除背景.
            SetStyle(ControlStyles.DoubleBuffer, true);

            //初始化PowerMainForm
            PowerMainForm_Initial();
            mainGrid_Initial();
            
            ///PowerMainForm_RefreshLoad();已包含在GetDetailIni

            //初始化OpenListView
            OpenListView_ColumnsHeader_Initial();
            OpenListView_GetDetailIni();
            OpenListView_SetCallback();
        }
        #endregion

        #region 0-0界面函数 PowerMainForm_Initial()
        /// <summary>
        /// 修改Openlist界面参数
        /// </summary>
        private void PowerMainForm_Initial()
        {
            /// Font
            DevExpress.Utils.AppearanceObject.DefaultFont = new System.Drawing.Font("STXihei", 11);
            foreach (Control ctrl in lsCtrl_Visible)
            {
                ctrl.Font = new System.Drawing.Font("STXihei", 11);
            }
            this.TxtB_No_Fake.Font = new System.Drawing.Font("STXihei", 11);

            /// Configuration
            if (showLionHeadOrNot == 0)
            {
                this.Logo.Image.Dispose();
            }

            /// ToolTip
                ImageList list = new ImageList();
            //list.Images.Add(Image.FromFile(@"d:\Information.png"), Color.Transparent);
            HyqCtrl.NewToolTip(this.Lab_No, "问题序号", "自动生成，不可编辑", 5000, list, 0);
            HyqCtrl.NewToolTip(this.Lab_Source, "问题来源", "填写问题提出者，例如“客户”、“组长”", 5000, list, 0);
            HyqCtrl.NewToolTip(this.Lab_Type, "问题类型", "填写问题类型，例如“软件”、“硬件”", 5000, list, 0);
            HyqCtrl.NewToolTip(this.Lab_EDate, "原定关闭日期", "只可以填写一次，建议在问题创建时填写", 5000, list, 0);
            HyqCtrl.NewToolTip(this.Lab_EDate2, "现定关闭日期", "可以多次修改，建议在问题延期时填写", 5000, list, 0);
            
            ///Control
            this.TxtB_No.ReadOnly = true;
            this.TxtB_Date.ReadOnly = true;
            this.UpdateTimer.Enabled = true;
            this.UpdateTimer.Stop(); 
            this.UpdateTimer.Interval = 10;
            this.ListPanel.SelectedPageIndex = 0;
            this.Opacity = 0.95;
            this.AllowDrop = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.ControlBox = true;
            VisibleIconReset();
        }
        #endregion

        #region 0-1界面函数 mainGrid_Initial()
        /// <summary>
        /// 修改Openlist界面参数
        /// </summary>
        private void mainGrid_Initial()
        {
            this.mainGrid.Font = new System.Drawing.Font("STXihei", 11, FontStyle.Regular);

            ///DateGridView
            this.mainGridView.PopulateColumns(); //显示gridCOntrol数据
            //this.mainGridView.BestFitColumns();
            this.mainGridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.mainGridView.Appearance.HeaderPanel.Font = new System.Drawing.Font("STXihei", 11, FontStyle.Regular);
            this.mainGridView.OptionsView.RowAutoHeight = true; //自动设置行高
            this.mainGridView.OptionsView.ColumnHeaderAutoHeight = DevExpress.Utils.DefaultBoolean.True;
            //this.mainGridView.OptionsView.RowsDefaultCellStyle = true;
            //this.mainGrid.UseEmbeddedNavigator = false;  //隐藏导航栏
            //this.mainGridView.OptionsView.AllowCellMerge = true; //允许自动合并单元格
            //this.mainGridView.OptionsBehavior.Editable = false; //允许用户修改
            this.mainGridView.OptionsCustomization.AllowSort = false;
            this.mainGridView.OptionsView.ColumnAutoWidth = false;

            this.mainGridView.CustomDrawRowIndicator +=new DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventHandler(mainGridView_CustomDrawRowIndicator);
            this.mainGridView.RowCountChanged +=new EventHandler(mainGridView_RowCountChanged);

            this.mainGridView.Appearance.EvenRow.BackColor = Color.FromArgb(150, 237, 243, 254);
            this.mainGridView.Appearance.OddRow.BackColor = Color.White;
            this.mainGridView.Appearance.OddRow.ForeColor = Color.Black;
            this.mainGridView.Appearance.EvenRow.ForeColor = Color.Black;
            this.mainGridView.Appearance.Row.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.mainGridView.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.mainGridView.RowSeparatorHeight = 2; //行间距            
        }
        #endregion

        #region 0-2界面函数 PowerMainForm_RefreshLoad()
        /// <summary>
        /// 修改Openlist界面参数
        /// </summary>
        private void PowerMainForm_RefreshLoad()
        {
            this.FileLoad.ClearLinks();
            RecallStack.ClearStack();
            string[] filespath;
            try { filespath = Directory.GetFiles(@".\Save\"); }
            catch { return; }
            HYQFileInfoList fileList = new HYQFileInfoList(filespath);
            foreach (FileInfoWithIcon file in fileList.list)
            {
                //if (file.fileInfo.Name.EndsWith("_(Archieve).ini"))
                //{
                //    continue;
                //}
                DevExpress.XtraBars.BarButtonItem bt = new BarButtonItem();
                bt.Caption = System.IO.Path.GetFileNameWithoutExtension(file.fileInfo.Name);//Split('.')[0];
                if (saveFilePath_N == bt.Caption)
                {
                    bt.ItemAppearance.SetFont(new System.Drawing.Font("STXihei", 11, FontStyle.Bold));
                    bt.LargeGlyph = DevExpress.Images.ImageResourceCache.Default.
                    GetImageById("BOReport2", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
                    this.TxtB_Title.Text = saveFilePath_N;
                    this.TxtB_Title.Left = (this.PanelA.Width - this.Lab_No.Left -this.TxtB_Title.Width - 25) / 2 + this.Lab_No.Left;
                }
                else
                {
                    bt.ItemAppearance.SetFont(new System.Drawing.Font("STXihei", 11, FontStyle.Regular));
                    bt.LargeGlyph = DevExpress.Images.ImageResourceCache.Default.
                    GetImageById("BOReport", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored); 
                }
                bt.ItemClick += new ItemClickEventHandler(FormatLoad_Other_ItemClick);
                FileLoad.AddItem(bt);
            }
        }
        #endregion

        #region 0-3界面函数 OpenListView_ColumnsHeader_Initial()
        /// <summary>
        /// 修改Openlist界面的“列”头
        /// </summary>
        private void OpenListView_ColumnsHeader_Initial()
        {
           
            DataTable dt = new DataTable();
            if (ListPanel.SelectedPageIndex == 0)
            {
                ///标准OPL
                foreach (string str in lb_HeaderName)
                {
                    dt.Columns.Add(str);  ///增加各个列头
                }
                this.mainGrid.DataSource = dt;
            }
            else 
            {
                ///已归档OPL
                foreach (string str in lb_HeaderName2)
                {
                    dt.Columns.Add(str);  ///增加各个列头
                }
                this.mainGrid.DataSource = dt;
            }
            IniFlag = true;
            OpenListView_ColumnsWidthEdit();

        }

        private void OpenListView_ColumnsWidthEdit()
        {
            this.mainGridView.BeginUpdate();
            if (ListPanel.SelectedPageIndex == 0)
            {
                if (this.WindowState == FormWindowState.Maximized)
                {
                    ///最大化 + OPL
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[0], lb_HeaderName[0], false, 80);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[1], lb_HeaderName[1], false, 120);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[2], lb_HeaderName[2], false, 100);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[3], lb_HeaderName[3], false, 100);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[4], lb_HeaderName[4], true, 200);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[5], lb_HeaderName[5], true, 120);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[6], lb_HeaderName[6], true, this.OpenList.Width - 1130);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[7], lb_HeaderName[7], false, 120);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[8], lb_HeaderName[8], false, 120);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[9], lb_HeaderName[9], false, 100);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[10], lb_HeaderName[10], false, 100);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[11], lb_HeaderName[11], false, 100);
                }
                else
                {
                    ///正常界面 + OPL
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[0], lb_HeaderName[0], false, 80);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[1], lb_HeaderName[1], false, 100);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[2], lb_HeaderName[2], false, 100);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[3], lb_HeaderName[3], false, 100);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[4], lb_HeaderName[4], true, 280);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[5], lb_HeaderName[5], true, 120);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[6], lb_HeaderName[6], true, 520);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[7], lb_HeaderName[7], false, 100);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[8], lb_HeaderName[8], false, 100);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[9], lb_HeaderName[9], false, 80);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[10], lb_HeaderName[10], false, 100);
                    HyqDG.Dev_EditCol(this.OpenListView.Columns[11], lb_HeaderName[11], false, 100);
                }
            }
            else
            {
                if (this.WindowState == FormWindowState.Maximized)
                {
                    ///最大化 + CPL
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[0], lb_HeaderName2[0], false, 80);  //修改列头具体参数
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[1], lb_HeaderName2[1], false, 120);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[2], lb_HeaderName2[2], false, 100);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[3], lb_HeaderName2[3], false, 100);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[4], lb_HeaderName2[4], true, 170);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[5], lb_HeaderName2[5], true, 120);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[6], lb_HeaderName2[6], true, this.ArchieveList.Width - 1160);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[7], lb_HeaderName2[7], false, 120);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[8], lb_HeaderName2[8], false, 120);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[9], lb_HeaderName2[9], false, 120);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[10], lb_HeaderName2[10], false, 120);
                }
                else
                {
                    ///正常界面 + CPL
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[0], lb_HeaderName2[0], false, 80);  //修改列头具体参数
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[1], lb_HeaderName2[1], false, 100);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[2], lb_HeaderName2[2], false, 100);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[3], lb_HeaderName2[3], false, 100);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[4], lb_HeaderName2[4], true, 280);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[5], lb_HeaderName2[5], true, 100);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[6], lb_HeaderName2[6], true, 440);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[7], lb_HeaderName2[7], false, 100);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[8], lb_HeaderName2[8], false, 120);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[9], lb_HeaderName2[9], false, 120);
                    HyqDG.Dev_EditCol(this.ArchieveListView.Columns[10], lb_HeaderName2[10], false, 120);
                }
                
            }

            //取消第一行的显示
            this.mainGridView.Columns[0].Visible = false;
            this.OpenListView.Columns[11].Visible = false;

            //设置控件宽度
            OpenListView_CollectionWidthEdit();

            //增加排序
            this.OpenListView.CustomColumnSort += new DevExpress.XtraGrid.Views.Base.CustomColumnSortEventHandler(OpenListView_CustomColumnSort);
            this.ArchieveListView.CustomColumnSort += new DevExpress.XtraGrid.Views.Base.CustomColumnSortEventHandler(OpenListView_CustomColumnSort);


            this.mainGridView.EndUpdate();
        }
        #endregion

        #region 0-4界面函数 OpenListView_CollectionWidthEdit()
        private void OpenListView_CollectionWidthEdit()
        {
            #region TxtB_WidthEdit
            if (this.WindowState != FormWindowState.Maximized)
            {
                this.Width = 1760;
                this.Height = 990;
                this.PanelA.Width = 585;
                this.PanelB.Width = this.Width - this.PanelA.Width - 270;//this.PanelC.Width;

                this.Logo.Height = 150;
                this.Logo.Width = 125;
                this.Logo.Left = 0;               

                this.Lab_No.Left = this.Logo.Width + 25;
                this.Lab_Type.Left = this.Logo.Width + 25;
                this.Lab_Source.Left = this.Logo.Width + 25;
                this.Lab_Due.Left = this.Logo.Width + 215;
                this.Lab_EDate.Left = this.Logo.Width + 215;
                this.Lab_EDate2.Left = this.Logo.Width + 215;

                this.TxtB_Title.Left = (this.PanelA.Width - this.Lab_No.Left - this.TxtB_Title.Width - 25) / 2 + this.Lab_No.Left;

                this.TxtB_No.Visible = false;

                this.TxtB_No.Left = this.Lab_Source.Left + this.Lab_Source.Width + 5;
                this.TxtB_No_Fake.Left = this.Lab_Source.Left + this.Lab_Source.Width + 5;
                this.TxtB_Type.Left = this.Lab_Source.Left + this.Lab_Source.Width + 5;
                this.TxtB_Source.Left = this.Lab_Source.Left + this.Lab_Source.Width + 5;
                this.TxtB_Due.Left = this.Lab_EDate.Left + this.Lab_EDate.Width + 5;
                this.MC_EDate.Left = this.Lab_EDate.Left + this.Lab_EDate.Width + 5;
                this.MC_EDate2.Left = this.Lab_EDate.Left + this.Lab_EDate.Width + 5;

                this.TxtB_No.Width = 130;
                this.TxtB_No_Fake.Width = 130;
                this.TxtB_Type.Width = 130;
                this.MC_EDate.Width = 130;
                this.TxtB_Due.Width = 130;
                this.TxtB_Source.Width = 130;
                this.MC_EDate2.Width = 130;

                this.TxtB_No.Top = 53;
                this.TxtB_No_Fake.Top = 53;
                this.TxtB_Due.Top = 53;
                this.TxtB_Type.Top = this.TxtB_No.Top + this.TxtB_No.Height + 7;
                this.MC_EDate.Top = this.TxtB_No.Top + this.TxtB_No.Height + 7;
                this.TxtB_Source.Top = this.TxtB_Type.Top + this.TxtB_Type.Height + 7;
                this.MC_EDate2.Top = this.TxtB_Type.Top + this.TxtB_Type.Height + 7;

                this.Lab_No.Top = this.TxtB_No.Top + 5;
                this.Lab_Due.Top = this.TxtB_Due.Top + 5;
                this.Lab_Type.Top = this.TxtB_Type.Top + 5;
                this.Lab_EDate.Top = this.MC_EDate.Top + 5;
                this.Lab_Source.Top = this.TxtB_Source.Top + 5;
                this.Lab_EDate2.Top = this.MC_EDate2.Top + 5;

                this.Lab_Descrip.Left = 20;
                this.Lab_Track.Left = 340;
                this.Lab_Plan.Left = 480;

                this.TxtB_Descrip.Left = 20;
                this.TxtB_Track.Left = 340;
                this.TxtB_Plan.Left = 480;

                this.TxtB_Descrip.Width = 300;
                this.TxtB_Track.Width = 140;
                this.TxtB_Plan.Width = 400;

                this.Lab_FileList.Left = 0;
                this.TxtB_FileList.Left = 0;

                this.TxtB_FileList.Width = 215;
            }
            else
            {
                this.PanelA.Width = 300 + 125 + 130 + 35;
                this.PanelB.Width = this.Width - this.PanelA.Width - 300;//this.PanelC.Width;

                this.Logo.Height = 150;
                this.Logo.Width = 125;
                this.Logo.Left = 0;

                this.Lab_No.Left = this.Logo.Width + 25;
                this.Lab_Type.Left = this.Logo.Width + 25;
                this.Lab_Source.Left = this.Logo.Width + 25;
                this.Lab_Due.Left = this.Logo.Width + 215;
                this.Lab_EDate.Left = this.Logo.Width + 215;
                this.Lab_EDate2.Left = this.Logo.Width + 215;

                this.TxtB_Title.Left = (this.PanelA.Width - this.Lab_No.Left - this.TxtB_Title.Width - 25) / 2 + this.Lab_No.Left;

                this.TxtB_No.Visible = false;

                this.TxtB_No.Left = this.Lab_Source.Left + this.Lab_Source.Width + 5;
                this.TxtB_No_Fake.Left = this.Lab_Source.Left + this.Lab_Source.Width + 5;
                this.TxtB_Type.Left = this.Lab_Source.Left + this.Lab_Source.Width + 5;
                this.TxtB_Source.Left = this.Lab_Source.Left + this.Lab_Source.Width + 5;
                this.TxtB_Due.Left = this.Lab_EDate.Left + this.Lab_EDate.Width + 5;
                this.MC_EDate.Left = this.Lab_EDate.Left + this.Lab_EDate.Width + 5;
                this.MC_EDate2.Left = this.Lab_EDate.Left + this.Lab_EDate.Width + 5;

                this.TxtB_No.Width = 130;
                this.TxtB_No_Fake.Width = 130;
                this.TxtB_Type.Width = 130;
                this.MC_EDate.Width = 130;
                this.TxtB_Due.Width = 130;
                this.TxtB_Source.Width = 130;
                this.MC_EDate2.Width = 130;

                this.TxtB_No.Top = 53;
                this.TxtB_No_Fake.Top = 53;
                this.TxtB_Due.Top = 53;
                this.TxtB_Type.Top = this.TxtB_No.Top + this.TxtB_No.Height + 7;
                this.MC_EDate.Top = this.TxtB_No.Top + this.TxtB_No.Height + 7;
                this.TxtB_Source.Top = this.TxtB_Type.Top + this.TxtB_Type.Height + 7;
                this.MC_EDate2.Top = this.TxtB_Type.Top + this.TxtB_Type.Height + 7;

                //this.TxtB_Descrip.Left = 0;
                //this.TxtB_Track.Left = 335;
                //this.TxtB_Plan.Left = 495;

                this.TxtB_Track.Width = 140;
                this.TxtB_Descrip.Width = (this.PanelB.Width - this.TxtB_Track.Width - 40) * 4/10;
                this.TxtB_Plan.Width = (this.PanelB.Width - this.TxtB_Track.Width - 40) * 6/10;

                this.Lab_Descrip.Left = 0;
                this.Lab_Track.Left = this.TxtB_Descrip.Width + 20;
                this.Lab_Plan.Left = this.Lab_Track.Left + this.TxtB_Track.Width;
                this.TxtB_Descrip.Left = this.Lab_Descrip.Left;
                this.TxtB_Track.Left = this.Lab_Track.Left;
                this.TxtB_Plan.Left = this.Lab_Plan.Left;

                this.Lab_FileList.Left = 0;
                this.TxtB_FileList.Left = 0;

                this.TxtB_FileList.Width = 270;
            }
            #endregion
        }
        #endregion

        #region 0-5界面函数 OpenListView_SetCallback()
        private void OpenListView_SetCallback()
        {
            this.OpenListView.RowCellStyle += new DevExpress.XtraGrid.Views.Grid.RowCellStyleEventHandler(OpenListView_RowCellStyle);
            this.ArchieveListView.RowCellStyle += new DevExpress.XtraGrid.Views.Grid.RowCellStyleEventHandler(OpenListView_RowCellStyle);

            this.OpenListView.MouseDown += new MouseEventHandler(mainGridListView_MouseDown);
            this.OpenListView.MouseMove += new MouseEventHandler(mainGridListView_MouseMove);
            this.OpenListView.MouseUp += new MouseEventHandler(mainGridListView_MouseUp);
            this.ArchieveListView.MouseDown += new MouseEventHandler(mainGridListView_MouseDown);
            this.ArchieveListView.MouseMove += new MouseEventHandler(mainGridListView_MouseMove);
            this.ArchieveListView.MouseUp += new MouseEventHandler(mainGridListView_MouseUp);
            foreach (Control ctrl in lsCtrl_BackGroud)
            {
                ctrl.MouseDown += new MouseEventHandler(Form_MouseDown);
                ctrl.MouseMove += new MouseEventHandler(Form_MouseMove);
                ctrl.MouseUp += new MouseEventHandler(Form_MouseUp);
            }

            foreach (Control ctrl in lsCtrl_AutoSave)
            {
                ctrl.KeyUp += new KeyEventHandler(lsCtrl_AutoSave_KeyUp);
            }
            this.MC_EDate.EditValueChanged += new EventHandler(MC_EDate_EditValueChanged);
            this.MC_EDate2.EditValueChanged += new EventHandler(MC_EDate2_EditValueChanged);

            ///设置DoImport的图标函数
            DevExpress.XtraBars.BarButtonItem bt = new BarButtonItem();
            bt.Caption = "从UAES经典OPL导入";
            bt.ItemAppearance.SetFont(new System.Drawing.Font("STXihei", 11, FontStyle.Regular));
            bt.LargeGlyph = DevExpress.Images.ImageResourceCache.Default.
            GetImageById("newtablestyle", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
            bt.ItemClick += new ItemClickEventHandler(DoImport1);
            DoImport.AddItem(bt);

            ///设置DoExport1的图标函数
            //bt = new BarButtonItem();
            //bt.Caption = "导出到Excel文件（推荐）";
            //bt.ItemAppearance.SetFont(new System.Drawing.Font("STXihei", 11, FontStyle.Regular));
            //bt.LargeGlyph = DevExpress.Images.ImageResourceCache.Default.
            //GetImageById("freezepanes", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
            //bt.ItemClick += new ItemClickEventHandler(DoExport1);
            //DoExport.AddItem(bt);

            ///设置DoExport1的图标函数
            bt = new BarButtonItem();
            bt.Caption = "导出到UAES模板文件";
            bt.ItemAppearance.SetFont(new System.Drawing.Font("STXihei", 11, FontStyle.Regular));
            bt.LargeGlyph = DevExpress.Images.ImageResourceCache.Default.
            GetImageById("freezepanes", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
            bt.ItemClick += new ItemClickEventHandler(DoExport2);
            DoExport.AddItem(bt);

            ///设置DoFileSave1的图标函数
            bt = new BarButtonItem();
            bt.Caption = "另存为";
            bt.ItemAppearance.SetFont(new System.Drawing.Font("STXihei", 11, FontStyle.Regular));
            bt.LargeGlyph = DevExpress.Images.ImageResourceCache.Default.
            GetImageById("open", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
            bt.ItemClick += new ItemClickEventHandler(DoFileSave1);
            FileSave.AddItem(bt);

            ///设置DoFileSave2的图标函数
            bt = new BarButtonItem();
            bt.Caption = "重命名";
            bt.ItemAppearance.SetFont(new System.Drawing.Font("STXihei", 11, FontStyle.Regular));
            bt.LargeGlyph = DevExpress.Images.ImageResourceCache.Default.
            GetImageById("loadtheme", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
            bt.ItemClick += new ItemClickEventHandler(DoFileSave2);
            FileSave.AddItem(bt);

            ///设置FileSave的图标函数
            bt = new BarButtonItem();
            bt.Caption = "新建";
            bt.ItemAppearance.SetFont(new System.Drawing.Font("STXihei", 11, FontStyle.Regular));
            bt.LargeGlyph = DevExpress.Images.ImageResourceCache.Default.
            GetImageById("insert", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
            bt.ItemClick += new ItemClickEventHandler(FileNew_ItemClick);
            FileManage.AddItem(bt);

            ///设置DoDeleteFile的图标函数
            bt = new BarButtonItem();
            bt.Caption = "删除";
            bt.ItemAppearance.SetFont(new System.Drawing.Font("STXihei", 11, FontStyle.Regular));
            bt.LargeGlyph = DevExpress.Images.ImageResourceCache.Default.
            GetImageById("deletelist2", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
            bt.ItemClick += new ItemClickEventHandler(DoDeleteFile);
            FileManage.AddItem(bt);

            //设置DoSorting的图标函数
            DoSortingCheck1.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked;
        }

        #endregion

        #region 1界面函数 OpenListView_GetDetailIni()
        /// <summary>
        /// 初始化界面，读存档信息
        /// </summary>
        private void OpenListView_GetDetailIni()
        {
            string[] Items,subItems;
            string substr;
            List<string> _list = new List<string>();
            DataTable dt = new DataTable();
            ///显示等待界面
            //SubForm2.ProcessCommand();
            try
            {
                WaitForProgressing.ShowWaitForm();
            }
            catch { }
            

            ///重新读取savefile
            PowerMainForm_RefreshLoad();

            ///清空Datagridview
            dt = (DataTable)this.mainGrid.DataSource;
            if (dt != null)
            {
                dt.Rows.Clear();
            }
            else
            {
                dt = new DataTable();
            }
            
            //OpenList.DataSource = dt;   move to under
            VisibleIconReset();

            if (ListPanel.SelectedPageIndex == 0)
            {
                ///标准OPL
                ///重新读取item
                Items = HyqIni.GetItems("OPL", saveFilePath_D);
            }
            else 
            {
                ///已归档OPL
                Items = HyqIni.GetItems("CPL", saveFilePath_D);
            }

            if (Items != null)
            {
                for (int i = 0; i < Items.Length; i++)
                {
                    _list.Clear();
                    DataRow dr = dt.NewRow();
                    int ind = -1;
                    try
                    {
                        ind = Func.CalItemsIndex(Items, i);
                        substr = Func.CalItemsTail(Items, i);
                        if ((substr == "") || (substr == null))
                        {
                            continue;
                        }
                        subItems = Func.ConvertItems(substr, new string[] { "<SPLIT>" });
                    }
                    catch { continue; }
                    /*添加表格内容*/

                    if (ListPanel.SelectedPageIndex == 0)
                    {
                        ///标准OPL
                        dr[0] = (ind).ToString();                     //"序号" lb_HeaderName[0] ;
                        dr[1] = ItemReplaceString2Enter(subItems[1]); //"创建日期" lb_HeaderName[1] ;
                        dr[2] = ItemReplaceString2Enter(subItems[2]); //"来源" lb_HeaderName[2] ;
                        dr[3] = ItemReplaceString2Enter(subItems[3]); //"类型" lb_HeaderName[3] ;
                        dr[4] = ItemReplaceString2Enter(subItems[4]); //"问题描述" lb_HeaderName[4] ;
                        dr[5] = ItemReplaceString2Enter(subItems[5]); //"负责人" lb_HeaderName[5] ;
                        dr[6] = ItemReplaceString2Enter(subItems[6]); //"行动措施" lb_HeaderName[6] ;
                        dr[7] = ItemReplaceString2Enter(subItems[7]); //"原定计划日期" lb_HeaderName[7] ;
                        dr[8] = ItemReplaceString2Enter(subItems[8]); //"现定计划日期" lb_HeaderName[8] ;
                        dr[9] = ItemReplaceString2Enter(subItems[9]); //"状态" lb_HeaderName[9] ; 
                        //dr[10] 自动生成，预留项 - 还剩多少天到期
                        dr[11] = ItemReplaceString2Enter(subItems[10]); //"关闭日期" lb_HeaderName[11] ; 
                    }
                    else 
                    {
                        ///已归档OPL
                        dr[0] = (ind).ToString();                     //"序号" lb_HeaderName[0] ;
                        dr[1] = ItemReplaceString2Enter(subItems[1]); //"创建日期" lb_HeaderName[1] ;
                        dr[2] = ItemReplaceString2Enter(subItems[2]); //"来源" lb_HeaderName[2] ;
                        dr[3] = ItemReplaceString2Enter(subItems[3]); //"类型" lb_HeaderName[3] ;
                        dr[4] = ItemReplaceString2Enter(subItems[4]); //"问题描述" lb_HeaderName[4] ;
                        dr[5] = ItemReplaceString2Enter(subItems[5]); //"负责人" lb_HeaderName[5] ;
                        dr[6] = ItemReplaceString2Enter(subItems[6]); //"行动措施" lb_HeaderName[6] ;
                        dr[7] = ItemReplaceString2Enter(subItems[9]); //"状态" lb_HeaderName[9] ; 
                        dr[8] = ItemReplaceString2Enter(subItems[7]); //"原定计划日期" lb_HeaderName[7] ;
                        dr[9] = ItemReplaceString2Enter(subItems[8]); //"现定计划日期" lb_HeaderName[8] ;
                        dr[10] = ItemReplaceString2Enter(subItems[10]); //"关闭日期" lb_HeaderName[8] ;
                    }

                    dt.Rows.Add(dr);//HyqDG.Dev_AddNewRow(this.mainGridView, _list);
                    
                }
                ///标准OPL
                this.mainGrid.DataSource = dt;
                //选中对应行
                try
                {
                    this.mainGridView.UnselectRow(0);
                    this.mainGridView.SelectRow(curSelectionOPL);
                    this.mainGridView.FocusedRowHandle = curSelectionOPL; 
                    //this.mainGridView.MoveLast();
                    OpenList_Click(null, null);
                }
                catch { }
                
            }

            //调整列宽
            //this.mainGridView.BestFitColumns();

            //更新表格颜色和状态
            for (int i = 0; i < this.mainGridView.DataRowCount; i++)
            {
                OpenListView_CheckDetail(i);

                //删除冗余行
                try
                {
                    if (this.mainGridView.GetRowCellValue(i, this.mainGridView.Columns[0]).ToString().Trim() == "")
                    {
                        this.mainGridView.DeleteRow(i);
                    }
                }
                catch {  }
                
            }

            //排序
            DoSorting1_Carryout(sort1Flag);
            mainGridView.BeginSort();
            mainGridView.EndSort();

            try
            {
                WaitForProgressing.CloseWaitForm();
            }
            catch
            { }
                
        }

        private string ItemReplaceString2Enter(string subItem)
        {
            try
            {
                string str = subItem.Replace("<ENTER>", Environment.NewLine);
                str = str.Replace("<TRACKSPACE>", Environment.NewLine + Environment.NewLine);
                return str;
            }
            catch
            {
                return null;
            }           
        }
        #endregion

        #region 2界面函数 OpenListView_CheckDetail(ind)：检查界面信息
        private void OpenListView_CheckDetail(int ind)
        {
            if (ListPanel.SelectedPageIndex != 0) return;
            int row = this.mainGridView.RowCount;
            int col = 8;//"现定计划日期" + Environment.NewLine + "Current Due";
            int col_s = 9;//"状态" + Environment.NewLine + "Sta.";
            int col_d = 10;//"到期" + Environment.NewLine + "Remain";
            string status;
            
            #region 自动计算日历状态
            try
            {
                if ((this.mainGridView.GetRowCellValue(ind, this.mainGridView.Columns[col_s]).ToString() == "Close"))
                {
                    return;
                }
                DateTime dt1 = Convert.ToDateTime(this.mainGridView.GetRowCellValue(ind, this.mainGridView.Columns[col]).ToString());
                DateTime dt2 = Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd"));
                TimeSpan d = dt1 - dt2;
                if (DateTime.Compare(dt2, dt1) > 0)
                {
                    status = "Delay";
                    this.mainGridView.SetRowCellValue(ind, this.mainGridView.Columns[col_s], status);
                    this.mainGridView.SetRowCellValue(ind, this.mainGridView.Columns[col_d], d.TotalDays.ToString());
                }
                else
                {
                    status = "Open";
                    this.mainGridView.SetRowCellValue(ind, this.mainGridView.Columns[9], status);
                    string tmp = d.TotalDays.ToString();
                    this.mainGridView.SetRowCellValue(ind, this.mainGridView.Columns[10], tmp);
                }
            }
            catch
            {
                return;//MessageUtil.ShowWarning("日期计算错误");
            }
            #endregion
        }
        private void OpenListView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            System.Drawing.Color red = new Color();
            System.Drawing.Color green = new Color();
            System.Drawing.Color yellow = new Color();
            red = System.Drawing.Color.Red;
            green = System.Drawing.Color.FromArgb(0, 255, 0);
            yellow = System.Drawing.Color.Yellow;

            string cols = "状态" + Environment.NewLine + "Sta.";
            if (e.RowHandle < this.mainGridView.RowCount)
            {
                try
                {
                    string val = this.mainGridView.GetRowCellValue(e.RowHandle, cols).ToString();
                    if (e.Column.FieldName == cols)
                    {
                        if (val == "Close")
                        {
                            e.Appearance.BackColor = green;
                            e.Appearance.ForeColor = System.Drawing.Color.Black;
                        }
                        if (val == "Delay")
                        {
                            e.Appearance.BackColor = red;
                            e.Appearance.ForeColor = System.Drawing.Color.White;
                        }
                        else if (val == "Open")
                        {
                            e.Appearance.BackColor = yellow;
                            e.Appearance.ForeColor = System.Drawing.Color.Black;
                        }
                    }
                }
                catch
                {
                    string a = e.Column.FieldName;
                }
            }
            

        }
         #endregion

        #region 3界面函数 OpenList_SelectionChanged：点击OpenList
        private void OpenList_SelectionChanged(object sender, EventArgs e)
        {
            if (this.mainGridView.SelectedRowsCount == 0) return;
            //EditItemAndSave(TxtB_No.Text);
            IniFlag = true;
            try
            {
                curSelectionOPL = this.mainGridView.GetSelectedRows()[0];
                HyqIni.PutINI("Main", "curSelectionOPL", curSelectionOPL, saveFilePath_M);
            }
            catch { }
            DetailUpdate(curSelectionOPL);
            try
            {
                TxtB_Track.SetSelected(TxtB_Track.Items.Count - 1, true);
            }
            catch { }
            IniFlag = false;
        }

        private void DetailUpdate(int SelectionOPL)
        {
            int row = SelectionOPL, ind_Item;
            string Txt_No, substr;
            string[] Items;
            newTrackFlag = false;
            //initial
            VisibleIconReset();
            try
            {
                Txt_No = this.mainGridView.GetRowCellValue(row, this.mainGridView.Columns[0]).ToString();
                ///重新读取item
                if (ListPanel.SelectedPageIndex == 0)
                {
                    ///标准OPL
                    Items = HyqIni.GetItems("OPL", saveFilePath_D);
                }
                else
                {
                    ///已归档OPL
                    Items = HyqIni.GetItems("CPL", saveFilePath_D);
                }
            }
            catch
            {
                return;
            }

            ind_Item = Func.FindItemInIni(Txt_No, Items);
            if (ind_Item == -1) return;
            substr = Func.CalItemsTail(Items, ind_Item);
            
            //载入到控件
            this.TxtB_No_Fake.Text = (row + 1).ToString();
            ControlUpdate(substr);
        }

        private void ControlUpdate(string valueStr)
        {
            string[] subItems;
            subItems = Func.ConvertItems(valueStr, new string[] { "<SPLIT>" });  
            /*添加表格内容*/

            ///更新OpenList中内容
            //if (ListPanel.SelectedPageIndex == 0)
            //{
            setControlText(TxtB_No,subItems,0);
            setControlText(TxtB_Date,subItems,1);
            setControlText(TxtB_Source,subItems,2);
            setControlText(TxtB_Type,subItems,3);
            setControlText(TxtB_Descrip,subItems,4);
            setControlText(TxtB_Due,subItems,5);
            setControlText(TxtB_Plan,subItems,6);
            setControlText(TxtB_EDate,subItems,7);
            setControlText(TxtB_EDate2,subItems,8);
            setControlText(TxtB_Status,subItems,9);
            setControlText(TxtB_ADate, subItems,10);
            //}

            try
            {
                string[] sTrack = Func.ConvertItems(subItems[6], new string[] { "<TRACKSPACE>" });
                for (int k = 0; k < sTrack.Length; k++)
                {
                    TxtB_Track.Items.Add(sTrack[k].Replace("<ENTER>", Environment.NewLine));
                }
            }
            catch { }
            

            #region MCDate的处理
            try
            {
                MC_EDate.DateTime = Convert.ToDateTime(TxtB_EDate.Text);
                MC_EDate2.DateTime = Convert.ToDateTime(TxtB_EDate2.Text);
            }
            catch { }
            //OpenListView_CheckDetail();//需要优化
            #endregion

            #region File的特殊处理
            #region 将File的文件信息存储至TxtBox
            string curPath = Directory.GetCurrentDirectory();
            string targetPath = ".\\Data\\" + saveFilePath_N + "\\" + TxtB_No.Text;
            TxtB_File.Text = "";
            try
            {
                Directory.CreateDirectory(".\\Data\\" + saveFilePath_N);
                Directory.CreateDirectory(targetPath);
            }
            catch { };
            #endregion

            Func.ListViewShowIcon(TxtB_FileList, targetPath);

            string[] targetFilePath = new string[0];
            try
            {
                targetFilePath = Directory.GetFiles(targetPath);
            }
            catch { }
            foreach (string f in targetFilePath)
            {
                TxtB_File.Text = TxtB_File.Text + f + "<ENTER>";//!!界面函数存储ini数据
            }
            #endregion
        }

        
        private void setControlText(Control Ctrl, string[] subItem, int index)
        {
            DevExpress.XtraEditors.TextEdit textEdit = Ctrl as DevExpress.XtraEditors.TextEdit;
            if (subItem.Length > index) 
            {
                textEdit.Text = subItem[index].Replace("<ENTER>", Environment.NewLine);
                textEdit.SelectionStart = textEdit.Text.Length;
            }
            else 
            {
                textEdit.Text = "";
            }
        }

        private void VisibleIconReset()
        {
            foreach (Control ctrl in ls_Edit)
            {
                try 
                { 
                    ctrl.Text = "";
                }
                catch { }
            }
            TxtB_No_Fake.Text = "";
            TxtB_Track.Items.Clear();
            TxtB_FileList.Items.Clear();
        }
        #endregion

        #region 5界面函数 TxtB_Track_SelectedIndexChanged：点击TxtB_Track
        private void TxtB_Track_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ///IniFlag = true;
                curSelectionTRK = TxtB_Track.SelectedIndex;
                TracklistUpdate(curSelectionTRK);
                ///IniFlag = false;
            }
            catch
            {
            }

        }

        private void TracklistUpdate(int SelectionTRK)
        {
            string content;
            string trackfirstsplit = " "; ///[XXX]后面跟的第一个字符，可以是空格，或回车
            try { content = TxtB_Track.Items[SelectionTRK].ToString(); }
            catch { return; }
            
            try
            {
                int i = content.IndexOf("]" + trackfirstsplit);
                int j = -1;
                DateTime dt;
                if (i == -1)
                {
                    j = content.IndexOf("]");
                    if ((DateTime.TryParse(content.Substring(1, j - 1), out dt) && (j != -1)) && ((j+1) < content.Length))
                    {
                        TxtB_Plan.Text = content.Substring(j + 1).Replace("<ENTER>", Environment.NewLine);//加上空格
                    }
                    else
                    {
                        TxtB_Plan.Text = "";
                    }
                }
                else
                {
                    TxtB_Plan.Text = content.Substring(i + trackfirstsplit.Length + 1).Replace("<ENTER>", Environment.NewLine);//加上空格
                }
                
            }
            catch { }
            //TxtB_Plan.Select(this.TxtB_Plan.TextLength, 0);
        }
        #endregion

        #region 6界面函数 EditItemAndSave
        private void EditItemAndSave(string id, string sec)
        {
            
            if (IniFlag) return;
            if (recallFlag)
            {
                recallFlag = false;
                return;
            }
            if (id.StartsWith("(序号)")) return ;
            string keyStr = id;
            string valueStr = null;
            int curIndex = curSelectionOPL;
            foreach (Control ctrl in lsCtrl_AutoSave)
            {
                string tmpStr = null;
                string date = System.DateTime.Now.ToString("yyyy-MM-dd");
                if (ctrl.Name == "TxtB_Plan")
                {
                    if (curIndex != -1)
                    {
                        try
                        {
                            tmpStr = this.mainGridView.GetRowCellValue(curIndex,
                                this.mainGridView.Columns["行动措施" + Environment.NewLine + "Action"]).ToString();
                            tmpStr = tmpStr.Replace(Environment.NewLine + Environment.NewLine + "[20", "<TRACKSPACE>[20");
                            tmpStr = tmpStr.Replace(Environment.NewLine, "<ENTER>");
                            tmpStr = tmpStr.Replace("\n", "<ENTER>");
                            valueStr = valueStr + tmpStr + "<SPLIT>";
                        }
                        catch
                        {
                            valueStr = valueStr + "<SPLIT>";
                        }
                    }
                }
                else
                {
                    tmpStr = ctrl.Text.Replace(Environment.NewLine, "<ENTER>");
                    tmpStr = tmpStr.Replace("\n", "<ENTER>");
                    valueStr = valueStr + tmpStr + "<SPLIT>";
                }

            }
            saveIntoFile(sec, keyStr, valueStr, saveFilePath_D);
        }

        public void saveIntoFile(string sec, string keyStr, string valueStr, string savepath)
        {
            string oldvalue = HyqIni.GetINI(sec, keyStr, null, savepath);
            if (oldvalue != valueStr)
            {
                itemRecall recall = new itemRecall();
                HyqIni.PutINI(sec, keyStr, valueStr, savepath);  ///保存
                recall.saveRecall(oldvalue, valueStr, keyStr);  ///存储撤销信息
                RecallStack.PushToUndoStack(recall);            ///更新撤销栈
            }
            else
            { }
        }

        /// <summary>
        /// 控件按下的时候开始计时，再次按下重计时
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lsCtrl_AutoSave_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if ((e.KeyData == Keys.Control) || (e.KeyData == Keys.Alt) || (e.KeyData == Keys.CapsLock) ||
                    (e.KeyData == Keys.Tab) || (e.KeyData == Keys.Shift) || (e.KeyData == Keys.Up) ||
                    (e.KeyData == Keys.Down) || (e.KeyData == Keys.Left) || (e.KeyData == Keys.Right))
                {
                    return;
                }
            }
            catch { }            
            if (ListPanel.SelectedPageIndex == 1) { return; } //在Archieve状态下不能更改
            UpdateTimer_Ticks = 0;
            this.UpdateTimer.Start();
        }

        /// <summary>
        /// 如果大于100ms，则保存当前，执行EditSave
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            if (ListPanel.SelectedPageIndex == 1) { return; } //在Archieve状态下不能更改
            if (UpdateTimer_Ticks > 35) //大于100毫秒
            {
                UpdateItem(TxtB_Source, null);
                UpdateItem(TxtB_Type, null);
                UpdateItem(TxtB_Descrip, null);
                UpdateItem(TxtB_Due, null);
                UpdateItem(TxtB_Plan, null);
                UpdateItem(TxtB_EDate, null);
                UpdateItem(TxtB_EDate2, null);
                EditItemAndSave(TxtB_No.Text, "OPL");
                UpdateTimer_Ticks = 0;
                this.UpdateTimer.Stop();
            }
            else
            {
                UpdateTimer_Ticks++;
            }
        }
        #endregion

        #region 7界面函数 UpdateItem
        /// <summary>
        /// 将control内容更新至Openlist，不能更新至Achievelist
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateItem(object sender, KeyEventArgs e)
        {
            int row = curSelectionOPL, row1 = curSelectionTRK;
            int col = 0;
            int cnt = -1;
            var tmp = new DevExpress.XtraEditors.TextEdit();
            string date = System.DateTime.Now.ToString("yyyy-MM-dd");
            string result = null;
            string content;
            List<string> slist = new List<string>();
            List<string> slist_real = new List<string>();
            string trackfirstsplit = " "; ///[XXX]后面跟的第一个字符，可以是空格，或回车
            if (sender.GetType().Equals(typeof(DevExpress.XtraEditors.TextEdit)))
            {
                tmp = sender as DevExpress.XtraEditors.TextEdit;
            }
            else if (sender.GetType().Equals(typeof(DevExpress.XtraEditors.MemoEdit)))
            {
                tmp = sender as DevExpress.XtraEditors.MemoEdit;
            }
            else
            {
                
                return;
            }
            for (int i = 1; i < lsCtrl_AutoSave.Count; i++) 
            {
                if (tmp.Name == lsCtrl_AutoSave[i].Name)
                {
                    col = i;
                    if (tmp.Name == "TxtB_Plan")
                    {
                        #region 处理Track List
                        try { content = this.mainGridView.GetRowCellValue(row, this.mainGridView.Columns[col]).ToString(); }//OPL里面具体的单元格显示的信息}
                        catch { content = null; }
                        if ((content != "") && (content != null))
                        {
                            ///先取出OPL中的原始数据（将"行"转换为"表"）             
                            slist = Func.ConvertLine2List(content);

                            foreach (string str in slist) //sreal代表实际TxtB_Track中的项目数
                            {
                                if (str.StartsWith("[20") )
                                {
                                    slist_real.Add(str + Environment.NewLine);
                                    cnt++;
                                }
                                else if ((str.Trim() != "") && (cnt != -1))
                                {
                                    slist_real[cnt] += str + Environment.NewLine;
                                }
                                else if ((str.Trim() != "") && (cnt == -1))
                                {
                                    slist_real.Add(str + Environment.NewLine);
                                    cnt++;
                                }
                                else
                                { 
                                    ///do not add
                                }
                            }

                            ///判断是否为新增项
                            
                            if (newTrackFlag == true)
                            {
                                if (IniFlag) { return; }
                                ///处理第row1项
                                slist_real.Add("[" + date + "]" + trackfirstsplit + tmp.Text + Environment.NewLine);
                                TxtB_Track.Items.Add("[" + date + "]" + trackfirstsplit + tmp.Text);
                                TxtB_Track.SetSelected(TxtB_Track.Items.Count - 1, true);
                            }
                            else
                            {
                                //非新增项
                                string sTmp = "";
                                if(row1 >= 0)
                                {
                                    sTmp = slist_real[row1].ToString();
                                }
                                if (sTmp.StartsWith("["))
                                {
                                    int posTmp = sTmp.IndexOf("]" + trackfirstsplit);
                                    if (tmp.Text == "<DELETE>")
                                    {
                                        slist_real.RemoveAt(row1);
                                        this.TxtB_Track.Items.RemoveAt(curSelectionTRK);
                                    }
                                    else 
                                    {
                                        slist_real[row1] = sTmp.Substring(0, posTmp + trackfirstsplit.Length + 1) + tmp.Text +Environment.NewLine; //表格中的日期跟踪
                                        this.TxtB_Track.Items[row1] = slist_real[row1];
                                    }

                                }
                            }

                            foreach (string str in slist_real)
                            {
                                result += str + Environment.NewLine; //Tracklist间间隔回车符
                            }
                            result = Func.DeleteTail(result, Environment.NewLine);
                            result = Func.DeleteTail(result, Environment.NewLine);
                            this.mainGridView.SetRowCellValue(row,this.mainGridView.Columns[col],result);
                        }
                        else
                        {
                            if (IniFlag) { return; }
                            if ((TxtB_No.Text.Trim() != "") && (tmp.Text !=""))
                            {
                                this.mainGridView.SetRowCellValue(row, this.mainGridView.Columns[col], "[" + date + "]" + trackfirstsplit + tmp.Text);

                                try
                                {
                                    TxtB_Track.Items[0] = "[" + date + "]" + trackfirstsplit + tmp.Text;
                                }
                                catch
                                {
                                    TxtB_Track.Items.Add("[" + date + "]" + trackfirstsplit + tmp.Text);
                                }

                                TxtB_Track.SetSelected(0, true);
                            }

                        }
                        #endregion
                    }
                    else
                    {
                        try
                        {
                            this.mainGridView.SetRowCellValue(row,this.mainGridView.Columns[col],tmp.Text);
                        }
                        catch { }
                    }
                    
                }
            }

            if ((tmp.Name == "TxtB_EDate") || (tmp.Name == "TxtB_EDate2"))
            {
                OpenListView_CheckDetail(row);
            }
        }
        #endregion

        #region Extra界面函数 主界面双击
        private void Form1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            switch (winformSize)
            {
                case 0:
                    {
                        ///normal
                        winformSize = 1;
                        this.WindowState = FormWindowState.Normal;
                        foreach (Control ctrl in this.Controls)
                        {
                            ctrl.Visible = false;
                        }
                        break;
                    }
                case 1:
                    {
                        ///max
                        winformSize = 2;
                        this.WindowState = FormWindowState.Maximized;
                        foreach (Control ctrl in this.Controls)
                        {
                            ctrl.Visible = false;
                        }
                        break;
                    }
                case 2:
                    {
                        ///mini
                        winformSize = 0;
                        this.WindowState = FormWindowState.Normal;
                        foreach (Control ctrl in this.Controls)
                        {
                            ctrl.Visible = false;
                        }
                        break;
                    }
                default:
                    {
                        winformSize = 1;
                        break;
                    }
            }
            OpenListView_GetDetailIni();
            //if (isMiniSize)
            //{
            //    foreach (Control ctrl in this.Controls)
            //    {
            //        ctrl.Visible = false;
            //        foreach (Control ctrl2 in lsCtrl_Visible)
            //        {
            //            ctrl2.Visible = true;
            //        }
            //    }
            //}

        }

        private void FormMax_Click(object sender, EventArgs e)
        {
            if (this.WindowState != FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Maximized;
                winformSize = 1;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
                winformSize = 0;
            }
            foreach (Control ctrl in this.Controls)
            {
                ctrl.Visible = true;
            }
            OpenListView_GetDetailIni(); ;
        }
        #endregion

        #region Extra界面函数 OpenList_CellValueChanged：OpenList参数更新
        private void OpenList_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //CheckDetail();
        }
        #endregion

        #region Extra界面函数 小窗口快捷键触发OpenList_KeyUp和TxtB_Track_KeyUp
        /// <summary>
        /// 问题清单Openlist中按下快捷键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void OpenList_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Modifiers.CompareTo(Keys.Control) == 0 && e.KeyCode == Keys.C)
            {
                OpenList_Copy(null,null);
            }

            if (e.Modifiers.CompareTo(Keys.Control) == 0 && e.KeyCode == Keys.V)
            {
                OpenList_Paste(null, null);
            }
            if (e.KeyData == Keys.Enter)
            {
                //MessageBox.Show("回车键");
            }
            else if (e.KeyData == Keys.Delete)
            {
                try
                {
                    //MessageBox.Show("删除键");
                    DoDelOPL_ItemClick(null, null);
                }
                catch { }
            }
            else if (e.KeyData == Keys.F)
            {
                if (this.mainGridView.OptionsFind.AlwaysVisible == true)
                {
                    this.mainGridView.OptionsFind.AlwaysVisible = false;
                }
                else
                {
                    this.mainGridView.OptionsFind.AlwaysVisible = true;
                }
            }
            else if (e.KeyData == Keys.Down || e.KeyData == Keys.Up)
            {
                OpenList_Click(null, null);
            }
        }

        /// <summary>
        /// 跟踪列表Tracklist中按下快捷键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtB_Track_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                //MessageBox.Show("回车键");
            }
            else if (e.KeyData == Keys.Delete)
            {
                try
                {
                    //MessageBox.Show("删除键");
                    DoDelPlan_ItemClick(null,null);
                }
                catch { }
            }
            
        }
        #endregion

        #region Extra界面函数 拖动文件
        /// <summary>
        /// 将文件拖入TxtB_File框时触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtB_File_DragDrop(object sender, DragEventArgs e)
        {

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                string curPath = Directory.GetCurrentDirectory();
                string targetPath = curPath + ".\\Data\\" + saveFilePath_N + "\\" + TxtB_No.Text;
                try
                {
                    Directory.CreateDirectory(".\\Data\\" + saveFilePath_N);
                    Directory.CreateDirectory(targetPath);
                }
                catch { };

                foreach (string f in files)
                {
                    string filename = System.IO.Path.GetFileName(f);
                    //string dir = System.IO.Path.GetDirectoryName(f);                    
                    filename = HYQReNameHelper.FileReName(targetPath + "\\" + filename); //如果文件目录下已存在，则重命名，否则返回原命名
                    File.Copy(f, targetPath + "\\" + filename, true);
                }

                #region 将File的文件信息存储至TxtBox
                string[] targetFilePath = Directory.GetFiles(targetPath);
                TxtB_File.Text = "";
                foreach (string f in targetFilePath)
                {
                    TxtB_File.Text = TxtB_File.Text + f + "<ENTER>";//!!界面函数存储ini数据
                }
                #endregion

                Func.ListViewShowIcon(TxtB_FileList, targetPath);
            }
        }

        private void TxtB_File_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;
            else
                e.Effect = DragDropEffects.None;
        }

        private void TxtB_FileList_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //ListViewHitTestInfo info = TxtB_FileList.HitTest(e.X, e.Y);
            try
            {
                int ind = this.TxtB_FileList.SelectedIndex;
                string[] targetFilePath = Func.ConvertItems(TxtB_File.Text, new string[] { "<ENTER>" });
                System.Diagnostics.Process.Start(targetFilePath[ind]);
            }
            catch { }
        }

        private void Lab_File_DoubleClick(object sender, EventArgs e)
        {
            /*打开File存储的文件夹*/
            System.Diagnostics.Process.Start(Directory.GetCurrentDirectory() + "\\" + TxtB_No.Text);
        }
        #endregion

        #region Extra界面函数 主窗口点击移动
        /// <summary>
        /// 鼠标按下主窗口时触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (beginMove == true) beginMove = false;
                else beginMove = true;
                curX = MousePosition.X;
                curY = MousePosition.Y;
            }
        }

        /// <summary>
        /// 鼠标移动主窗口时触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form_MouseMove(object sender, MouseEventArgs e)
        {
            if (beginMove)
            {
                this.Left += MousePosition.X - curX;
                this.Top += MousePosition.Y - curY;
                curX = MousePosition.X;
                curY = MousePosition.Y;
            }
        }

        /// <summary>
        /// 鼠标松开主窗口时触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                curY = 0;
                curX = 0;
                beginMove = false;
            }
        }

        /// <summary>
        /// 主窗口关闭按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 主窗口最小化按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMin_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        #endregion

        #region Extra界面函数 操作日历
        /// <summary>
        /// 操作日历EDate，修改日期时触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MC_EDate_EditValueChanged(object sender, EventArgs e)
        {
            //bool editflg = true;

            //try
            //{
            //    if (this.TxtB_EDate.Text != Convert.ToDateTime(TxtB_Date.Text).AddDays(5).ToString("d"))
            //    {
            //        DialogResult result = DevExpress.XtraEditors.XtraMessageBox.Show
            //            ("是否确定需要修改“原定”计划日期？建议修改“现定”计划日期", "Warning", 
            //            MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            //        if (result == DialogResult.OK) { editflg = true; } else { editflg = false; }
            //    }
            //}
            //catch { }
            if (IniFlag == true) return;

            if (true)
            {
                if (EDateCheck(TxtB_Date.Text, MC_EDate.DateTime.ToShortDateString()))
                {
                    this.TxtB_EDate.Text = MC_EDate.DateTime.ToString("yyyy-MM-dd");//ToShortDateString();
                    this.MC_EDate2.DateTime = MC_EDate.DateTime;
                    this.TxtB_EDate2.Text = TxtB_EDate.Text;   
                    lsCtrl_AutoSave_KeyUp(null, null);
                    //this.MC_EDatea.Visible = false;
                }
                else
                {
                    MC_EDate.DateTime = Convert.ToDateTime(TxtB_EDate.Text);
                    MessageUtil.ShowTips("“原定计划日期”不能早于“问题创建日期”");
                }
            }
        }

        /// <summary>
        /// 操作日历EDate2，修改日期时触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MC_EDate2_EditValueChanged(object sender, EventArgs e)
        {
            if (IniFlag == true) return;

            if (EDateCheck(TxtB_Date.Text, MC_EDate2.DateTime.ToShortDateString()))
            {
                this.TxtB_EDate2.Text = MC_EDate2.DateTime.ToString("yyyy-MM-dd");
                lsCtrl_AutoSave_KeyUp(null, null);
                //this.MC_EDate2a.Visible = false;
            }
            else
            {
                MC_EDate2.DateTime = Convert.ToDateTime(TxtB_EDate2.Text);
                MessageUtil.ShowTips("“现定计划日期”不能早于“问题创建日期”");
            }
        }

        /// <summary>
        /// 本地函数：检查选中计划日期是否早于问题创建日期
        /// </summary>
        /// <param name="tobeChecked"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        private bool EDateCheck(string tobeChecked, string item)
        {
            try
            {
                if (Convert.ToDateTime(item) >= Convert.ToDateTime(tobeChecked))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                return true;
            }

        }

        #endregion

        #region Extra界面函数 点击Openlist切换Item
        private void OpenList_Click(object sender, EventArgs e)
        {
            IniFlag = true;
            OpenList_SelectionChanged(sender, e);
            IniFlag = false;
            if (TxtB_Status.Text == "Close")
            {
                DoClose.Caption = "Open";
                DoClose.ImageOptions.Image = DevExpress.Images.ImageResourceCache.Default.
            GetImageById("warning", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
                DoClose.ImageOptions.LargeImage = DevExpress.Images.ImageResourceCache.Default.
            GetImageById("warning", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);

            }
            else
            {
                DoClose.Caption = "Close";
                DoClose.ImageOptions.Image = DevExpress.Images.ImageResourceCache.Default.
            GetImageById("apply", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
                DoClose.ImageOptions.LargeImage = DevExpress.Images.ImageResourceCache.Default.
            GetImageById("apply", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
            }
        }
        #endregion

        #region Ribbon界面函数 复制Item
        private void OpenList_Copy(object sender, ItemClickEventArgs e)
        {
            if ((TxtB_No.Text != "") && (this.mainGridView.SelectedRowsCount > 0) && (ListPanel.SelectedPageIndex == 0))
            {
                //WaitForProgressing.ShowWaitForm();
                ///重新读取item
                string[] Items = HyqIni.GetItems("OPL", saveFilePath_D);
                string[] Items_s = HyqIni.GetItems("CPL", saveFilePath_D);
                string col_s = "状态" + Environment.NewLine + "Sta.";

                ///复制选中项
                int[] rows = this.mainGridView.GetSelectedRows();
                for (int i = 0; i < rows.Count(); i++)
                {
                    curSelectionOPL = rows[i];
                    string tmp = Func_LookDetailListInItem(curSelectionOPL);
                    copyStack.Add(tmp);
                    //保存在归档文件中
                }
                //WaitForProgressing.CloseWaitForm();
            }
        }
        #endregion

        #region Ribbon界面函数 粘贴Item
        private void OpenList_Paste(object sender, ItemClickEventArgs e)
        {
            if ((TxtB_No.Text != "") && (this.mainGridView.SelectedRowsCount > 0) && (ListPanel.SelectedPageIndex == 0))
            {                
                int ind;
                string[] Items, Items2, newItemPart;
                string tmp;
                List<int> indlist = new List<int>();
                DataTable dt = (DataTable)OpenList.DataSource;
                string newItem;

                //WaitForProgressing.ShowWaitForm();

                ///复制选中项
                int[] rows = this.mainGridView.GetSelectedRows();
                for (int i = 0; i < copyStack.Count; i++)
                {
                    newItem = copyStack[i];
                    newItemPart = Func.ConvertItems(newItem, new string[] { "<SPLIT>" });

                    #region 下述内容改自新增OPL
                    ///新建内容

                    try
                    {
                        ///判断新增序号大小
                        Items = HyqIni.GetItems("OPL", saveFilePath_D);
                        Items2 = HyqIni.GetItems("CPL", saveFilePath_D);


                        //int[] arr = new int[Items.Length + Items2.Length];
                        if (Items != null)
                        {
                            for (int j = 0; j < Items.Length; j++)
                            {
                                indlist.Add(Func.CalItemsIndex(Items, j));
                            }
                        }
                        if (Items2 != null)
                        {
                            for (int j = 0; j < Items2.Length; j++)
                            {
                                indlist.Add(Func.CalItemsIndex(Items2, j));
                            }
                        }

                        indlist.Sort();
                        ind = indlist[indlist.Count - 1];
                        curSelectionOPL = Items.Length;
                    }
                    catch
                    {
                        ind = 0;
                        curSelectionOPL = 0;
                    }

                    List<string> _list = new List<string>();

                    /*新增部分*/
                    VisibleIconReset();
                    if (dt == null)
                    {
                        return;
                        //for (int j = 0; j < 11; j++)
                        //{
                        //    dt.Columns.Add();
                        //}
                        //dt = new DataTable();
                    }
                    DataRow dr = dt.NewRow();
                    dr[0] = (ind + 1).ToString();
                    this.TxtB_No.Text = saveFilePath_N.Split('_')[0] + "_" + (ind + 1).ToString();//("0000");
                    this.TxtB_No_Fake.Text = curSelectionOPL.ToString();
                    this.TxtB_Date.Text = System.DateTime.Now.ToString("yyyy-MM-dd");
                    dr[1] = TxtB_Date.Text;
                    try { tmp = newItemPart[2].Replace("<ENTER>", Environment.NewLine);  }
                    catch { tmp = "(来源)"; }
                    dr[2] = tmp;
                    this.TxtB_Source.Text = tmp;
                    try { tmp = newItemPart[3].Replace("<ENTER>", Environment.NewLine);  }
                    catch { tmp = "(归类)"; }
                    dr[3] = tmp;
                    this.TxtB_Type.Text = tmp;
                    dr[4] = "";
                    this.TxtB_Descrip.Text = "";
                    try { tmp = newItemPart[5].Replace("<ENTER>", Environment.NewLine);  }
                    catch { tmp = System.Environment.UserName; }
                    dr[5] = tmp;
                    this.TxtB_Due.Text = tmp;
                    dr[6] = "";
                    this.TxtB_Track.Items.Clear();
                    this.TxtB_Plan.Text = "";
                    dr[7] = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
                    this.TxtB_EDate.Text = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
                    dr[8] = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
                    this.TxtB_EDate2.Text = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
                    dr[9] = "Open";
                    this.TxtB_Status.Text = "Open";
                    dr[10] = 5;
                    this.TxtB_ADate.Text = "";
                    dr[11] = "";
                    //TxtB_ADate.Text = System.DateTime.Now.AddDays(20).ToString("d");
                    dt.Rows.Add(dr);
                    this.mainGrid.DataSource = dt;

                    EditItemAndSave(TxtB_No.Text, "OPL");
                    try
                    {
                        this.mainGridView.MoveLast();
                        OpenList_Click(null, null);
                    }
                    catch
                    {
                        VisibleIconReset();
                    }


                    #region 创建临时文件夹
                    string curPath = Directory.GetCurrentDirectory();
                    string targetPath = ".\\Data\\" + saveFilePath_N + "\\" + TxtB_No.Text;
                    try { Func.ClearFolder(targetPath); }
                    catch { }
                    try
                    {
                        Directory.CreateDirectory(".\\Data\\" + saveFilePath_N);
                        Directory.CreateDirectory(targetPath);
                    }
                    catch { }
                    #endregion
                    #endregion
                }
                copyStack.Clear();
                //WaitForProgressing.CloseWaitForm();
            }
        }
        #endregion

        #region listPanel界面函数 换页
        private void ListPanel_SelectedPageChanged(object sender, DevExpress.XtraBars.Navigation.SelectedPageChangedEventArgs e)
        {
            ///重新读取item
            if (ListPanel.SelectedPageIndex == 0)
            {
                ///标准OPL
                this.mainGrid = this.OpenList;
                this.mainGridView = this.OpenListView;

                DoAddOPL.Enabled = true;
                DoAddPlan.Enabled = true;
                DoDelPlan.Enabled = true;
                DoClose.Enabled = true;
                DoAchieve.Enabled = true;
                DoAchieve.Caption = "Achieve";

                //Initial
                mainGrid_Initial();
                OpenListView_ColumnsHeader_Initial();
                OpenListView_GetDetailIni();
            }
            else if (ListPanel.SelectedPageIndex == 1)
            {
                ///已归档OPL
                this.mainGrid = this.ArchieveList;
                this.mainGridView = this.ArchieveListView;

                DoAddOPL.Enabled = false;
                DoAddPlan.Enabled = false;
                DoDelPlan.Enabled = false;
                DoClose.Enabled = false;
                DoAchieve.Enabled = true;
                DoAchieve.Caption = "Reopen";

                //Initial
                mainGrid_Initial();
                OpenListView_ColumnsHeader_Initial();
                OpenListView_GetDetailIni();
            }
            else
            {
                DoAddOPL.Enabled = false;
                DoAddPlan.Enabled = false;
                DoDelPlan.Enabled = false;
                DoClose.Enabled = false;
                DoAchieve.Enabled = false;
            }

            

            //OpenListView.ClearSorting();
            //排序

        }
        #endregion

        #region Ribbon界面函数 点击界面，选择需要的File
        private void FormatLoad_Other_ItemClick(object sender, ItemClickEventArgs e)
        {

            saveFilePath_D = @".\Save\" + e.Item.Caption + ".ini";
            saveFilePath_N = e.Item.Caption;
            RecallStack.ClearStack();
            //saveFilePath_S = @".\Save\" + e.Item.Caption + "_(Archieve).ini";
            
            ///重新载入
            ListPanel.SelectedPageIndex = 0;
            OpenListView_GetDetailIni();
            HyqIni.PutINI("Main", "saveFilePath_D", saveFilePath_D, saveFilePath_M);
            HyqIni.PutINI("Main", "saveFilePath_N", saveFilePath_N, saveFilePath_M);
        }
        #endregion

        #region Ribbon界面函数 点击新建File
        private void FileNew_ItemClick(object sender, ItemClickEventArgs e)
        {
            FileNewClick();
        }

        private bool FileNewClick()
        {
            if (SubForm1 != null)
            {
                SubForm1.Activate();
                SubForm1.WindowState = FormWindowState.Normal;
            }
            if (SubForm1.ShowDialog() == DialogResult.OK)
            {             
                DataTable dt = (DataTable)this.mainGrid.DataSource;
                dt.Rows.Clear();
                this.mainGrid.DataSource = dt;
                DoAddOPL_ItemClick(null, null);
                PowerMainForm_RefreshLoad();
                OpenListView_GetDetailIni();
                VisibleIconReset();
                HyqIni.PutINI("Main", "saveFilePath_D", saveFilePath_D, saveFilePath_M);
                HyqIni.PutINI("Main", "saveFilePath_N", saveFilePath_N, saveFilePath_M);
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region Ribbon界面函数 点击增加OPL
        private void DoAddOPL_ItemClick(object sender, ItemClickEventArgs e)
        {
            
            int ind;
            string[] Items, Items2;
            string tmp;
            List<int> indlist = new List<int>();
            ///获取数据
            DataTable dt = (DataTable)OpenList.DataSource;          
       
            ///新建内容
            try
            {
                ///判断新增序号大小
                Items = HyqIni.GetItems("OPL", saveFilePath_D);
                Items2 = HyqIni.GetItems("CPL", saveFilePath_D);


                //int[] arr = new int[Items.Length + Items2.Length];
                if (Items != null)
                {
                    for (int i = 0; i < Items.Length; i++)
                    {
                        indlist.Add(Func.CalItemsIndex(Items, i));
                    }
                }
                if (Items2 != null)
                {
                    for (int i = 0; i < Items2.Length; i++)
                    {
                        indlist.Add(Func.CalItemsIndex(Items2, i));
                    }
                }

                indlist.Sort();
                ind = indlist[indlist.Count - 1];
                curSelectionOPL = Items.Length;
            }
            catch 
            { 
                ind = 0;
                curSelectionOPL = 0;
            }

            List<string> _list = new List<string>();

            /*新增部分*/
            VisibleIconReset();
            if ((dt == null) || (dt.Columns.Count == 0))
            {
                return;
                //dt = new DataTable();
                //for (int i = 0; i < 12; i++)
                //{
                //    dt.Columns.Add();
                //}                
            }
            DataRow dr = dt.NewRow();
            dr[0] = (ind + 1).ToString();
            this.TxtB_No.Text = saveFilePath_N.Split('_')[0] + "_" + (ind + 1).ToString();//("0000");
            this.TxtB_No_Fake.Text = curSelectionOPL.ToString();
            this.TxtB_Date.Text = System.DateTime.Now.ToString("yyyy-MM-dd");
            dr[1] = TxtB_Date.Text;
            try { tmp = this.mainGridView.GetRowCellValue(curSelectionOPL - 1, this.mainGridView.Columns[2]).ToString(); }
            catch { tmp = "(来源)"; }
            dr[2] = tmp;
            this.TxtB_Source.Text = tmp;
            try { tmp = this.mainGridView.GetRowCellValue(curSelectionOPL - 1, this.mainGridView.Columns[3]).ToString(); }
            catch { tmp = "(归类)"; }
            dr[3] = tmp;
            this.TxtB_Type.Text = tmp;
            dr[4] = "";
            this.TxtB_Descrip.Text = "";
            try { tmp = this.mainGridView.GetRowCellValue(curSelectionOPL - 1, this.mainGridView.Columns[5]).ToString(); }
            catch { tmp = System.Environment.UserName; }
            dr[5] = tmp;
            this.TxtB_Due.Text = tmp;
            dr[6] = "";
            this.TxtB_Track.Items.Clear();
            this.TxtB_Plan.Text = "";
            dr[7] = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
            this.TxtB_EDate.Text = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
            dr[8] = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
            this.TxtB_EDate2.Text = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
            dr[9] = "Open";
            this.TxtB_Status.Text = "Open";
            dr[10] = 5;
            this.TxtB_ADate.Text = "";
            dr[11] = "";
            //TxtB_ADate.Text = System.DateTime.Now.AddDays(20).ToString("d");
            dt.Rows.Add(dr);
            this.mainGrid.DataSource = dt;
            
            EditItemAndSave(TxtB_No.Text, "OPL");
            try
            {
                this.mainGridView.MoveLast();
                OpenList_Click(null, null);
            }
            catch
            {
                VisibleIconReset();
            }
            

            #region 创建临时文件夹
            string curPath = Directory.GetCurrentDirectory();
            string targetPath = ".\\Data\\" + saveFilePath_N + "\\" + TxtB_No.Text;
            try { Func.ClearFolder(targetPath); }
            catch { }
            try
            {
                Directory.CreateDirectory(".\\Data\\" + saveFilePath_N);
                Directory.CreateDirectory(targetPath);
            }
            catch { }
            #endregion
        }
        #endregion

        #region Ribbon界面函数 点击删除OPL
        private void DoDelOPL_ItemClick(object sender, ItemClickEventArgs e)
        {
            if ((TxtB_No.Text != "")&&(this.mainGridView.SelectedRowsCount > 0))
            {
                DialogResult result = DevExpress.XtraEditors.XtraMessageBox.Show("是否删除选中项？", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (result == DialogResult.OK)
                {
                    ///重新读取item
                    string[] Items;
                    if (ListPanel.SelectedPageIndex == 0)
                    {
                        ///标准OPL
                        Items = HyqIni.GetItems("OPL", saveFilePath_D);
                    }
                    else
                    {
                        ///已归档OPL
                        Items = HyqIni.GetItems("CPL", saveFilePath_D);
                    }

                    ///删除已选中项
                    int[] rows = this.mainGridView.GetSelectedRows();
                    for (int i = 0; i < rows.Count(); i++)
                    {
                        string Txt_No = this.mainGridView.GetRowCellValue(rows[i], this.mainGridView.Columns[0]).ToString();
                        int ind_Item = Func.FindItemInIni(Txt_No, Items);
                        string Txt_No_Exact = Func.CalItemsHead(Items, ind_Item);
                        if (ListPanel.SelectedPageIndex == 0)
                        {
                            ///标准OPL
                            saveIntoFile("OPL", Txt_No_Exact, null, saveFilePath_D);//删除对应项
                        }
                        else
                        {
                            ///已归档OPL
                            saveIntoFile("CPL", Txt_No_Exact, null, saveFilePath_D);//删除对应项
                        }
                    }
                    this.mainGridView.DeleteSelectedRows();

                    ///焦点移向下一行
                    try
                    {
                        this.mainGridView.MoveNext();
                        OpenList_Click(null, null);
                    }
                    catch
                    {
                        VisibleIconReset();
                    }
                }
            }
        }
        #endregion

        #region Ribbon界面函数 点击增加Plan
        private void DoAddPlan_ItemClick(object sender, ItemClickEventArgs e)
        {
            string date = System.DateTime.Now.ToString("yyyy-MM-dd");
            //TxtB_Track.Items.Add("[" + date + "] " );
            //TxtB_Track.SetSelected(TxtB_Track.Items.Count - 1,true);
            TxtB_Plan.Text = "";
            newTrackFlag = true; //标记新增
            UpdateItem(TxtB_Plan, null);
            newTrackFlag = false;
            EditItemAndSave(TxtB_No.Text, "OPL");
        }
        #endregion

        #region Ribbon界面函数 点击删除Plan
        private void DoDelPlan_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (ListPanel.SelectedPageIndex == 1) { return; }
            try
            {
                if ((TxtB_Track.Items.Count == 1) && (TxtB_Track.Items[0].ToString().Trim() == ""))
                {
                    return;
                }
            }
            catch
            { }
            
            if (TxtB_No.Text != "")
            {
                DialogResult result = DevExpress.XtraEditors.XtraMessageBox.Show("是否删除选中项？", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (result == DialogResult.OK)
                {
                    this.TxtB_Plan.Text = "<DELETE>";
                    UpdateItem(this.TxtB_Plan, null);
                    IniFlag = true;
                    this.TxtB_Plan.Text = "";
                    IniFlag = false;
                    EditItemAndSave(TxtB_No.Text, "OPL");
                    OpenList_Click(null, null);
                }
            }
        }
        #endregion

        #region Ribbon界面函数 点击归档OPL
        private void DoAchieve_ItemClick(object sender, ItemClickEventArgs e)
        {
            string qstr;
            if (ListPanel.SelectedPageIndex == 0)
            {
                qstr = "是否'归档'选中项？";
            }
            else
            {
                qstr = "是否重新'重载'选中项进Open List？";
            }
            DialogResult result = DevExpress.XtraEditors.XtraMessageBox.Show(qstr, "Info", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (result == DialogResult.OK)
            {
                WaitForProgressing.ShowWaitForm();
                
                ///重新读取item
                string[] Items = HyqIni.GetItems("OPL", saveFilePath_D);
                string col_s = "状态" + Environment.NewLine + "Sta.";
                List<int> int_d = new List<int>();
                ///归档已选中项
                int[] rows = this.mainGridView.GetSelectedRows();
                for (int i = rows.Count() - 1; i >= 0; i--)
                {
                ///for (int i = this.mainGridView.DataRowCount - 1; i >= 0; i--)
                ///{ //全部
                    curSelectionOPL = rows[i];
                    DetailUpdate(curSelectionOPL);
                    if (TxtB_Status.Text == "Close")
                    {
                        if (ListPanel.SelectedPageIndex == 0)
                        {
                            ///OPL界面
                            //保存在归档文件中
                            EditItemAndSave(TxtB_No.Text, "CPL");

                            //删除原OPL文件中的对应项
                            string Txt_No = this.mainGridView.GetRowCellValue(curSelectionOPL, this.mainGridView.Columns[0]).ToString();
                            int ind_Item = Func.FindItemInIni(Txt_No, Items);
                            string keystr = Func.CalItemsHead(Items, ind_Item);
                            string valuestr = null;
                            saveIntoFile("OPL", keystr, valuestr, saveFilePath_D);
                            int_d.Add(curSelectionOPL);
                        }
                        else
                        {
                            ///CPL界面
                            //保存在归档文件中
                            EditItemAndSave(TxtB_No.Text, "OPL");

                            //删除原OPL文件中的对应项
                            string Txt_No = this.mainGridView.GetRowCellValue(curSelectionOPL, this.mainGridView.Columns[0]).ToString();
                            int ind_Item = Func.FindItemInIni(Txt_No, Items);
                            string keystr = Func.CalItemsHead(Items, ind_Item);
                            string valuestr = null;
                            saveIntoFile("CPL", keystr, valuestr, saveFilePath_D);
                            int_d.Add(curSelectionOPL);
                        }
                    }
                }

                VisibleIconReset();

                foreach (int i2 in int_d)
                {
                    this.mainGridView.DeleteRow(i2);
                }
                this.mainGridView.MoveLast();
                curSelectionOPL = this.mainGridView.Columns.Count - 1;
                DetailUpdate(curSelectionOPL);
                WaitForProgressing.CloseWaitForm();
            }
        }
        #endregion

        #region Ribbon界面函数 点击关闭按钮
        private void DoClose_ItemClick(object sender, ItemClickEventArgs e)
        {
            if ((TxtB_No.Text != "") && (this.mainGridView.SelectedRowsCount > 0))
            {
                string qstr;
                string firstFlag;
                if (TxtB_Status.Text == "Close")
                {
                    qstr = "是否将选中项重新设置为''Open''？";
                    firstFlag = "Open";
                }
                else
                {
                    qstr = "是否'关闭'选中项？";
                    firstFlag = "Close";
                }
                DialogResult result = DevExpress.XtraEditors.XtraMessageBox.Show(qstr, "Info", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (result == DialogResult.OK)
                {
                    WaitForProgressing.ShowWaitForm();
                    ///重新读取item
                    string[] Items = HyqIni.GetItems("OPL", saveFilePath_D);
                    string[] Items_s = HyqIni.GetItems("CPL", saveFilePath_D);
                    string col_s = "状态" + Environment.NewLine + "Sta.";
                    
                    ///关闭已选中项
                    int[] rows = this.mainGridView.GetSelectedRows();
                    for (int i = 0; i < rows.Count(); i++)
                    {
                        curSelectionOPL = rows[i];
                        DetailUpdate(rows[i]);
                        TxtB_ADate.Text = System.DateTime.Now.ToString("yyyy-MM-dd");
                        if (firstFlag == "Open")
                        {
                            TxtB_Status.Text = "Open";
                        }
                        else
                        {
                            TxtB_Status.Text = "Close";
                        }
                        UpdateItem(TxtB_Status, null);

                        //保存在归档文件中
                        EditItemAndSave(TxtB_No.Text, "OPL");
                    }
                    WaitForProgressing.CloseWaitForm();
                }
            }

                
        }
        #endregion

        #region Extra界面函数 根据行号寻找Item
        private string Func_LookDetailListInItem(int rowInd)
        {
            string Txt_No;
            string substr = "";
            string[] Items;
            int ind_Item;
            try
            {
                Txt_No = this.mainGridView.GetRowCellValue(rowInd, this.mainGridView.Columns[0]).ToString();
                ///重新读取item
                if (ListPanel.SelectedPageIndex == 0)
                {
                    ///标准OPL
                    Items = HyqIni.GetItems("OPL", saveFilePath_D);
                }
                else
                {
                    ///已归档OPL
                    Items = HyqIni.GetItems("CPL", saveFilePath_D);
                }
            }
            catch
            {
                    return substr;
            }

            ind_Item = Func.FindItemInIni(Txt_No, Items);
            if (ind_Item == -1) return substr;
            substr = Func.CalItemsTail(Items, ind_Item);
            //subItems = Func.ConvertItems(substr, new string[] { "<SPLIT>" });
            return substr;
        }
        #endregion

        #region Ribbon界面函数 点击选择“导出到UAES EXCEL表”
        private void DoExport2(object sender, ItemClickEventArgs e)
        {
            SaveFileDialog exportFile = new SaveFileDialog();
            exportFile.Title = "Export to EXCEL file";
            exportFile.FileName = saveFilePath_N;
            exportFile.InitialDirectory = @"D:\";   //@是取消转义字符的意思
            exportFile.Filter = "All Files(*.*)|*.*|Office Excel(*.xlsx)|*.xlsx|2007 Excel(*.xls)|*.xls ";
            exportFile.FilterIndex = 2;
            exportFile.RestoreDirectory = true;

            if (exportFile.ShowDialog() == DialogResult.OK)
            {
                ///开始
                WaitForProgressing.ShowWaitForm();

                string filename = System.IO.Path.GetFileNameWithoutExtension(exportFile.FileName);
                string sourcepath = System.IO.Path.GetFullPath(exportFile.FileName);
                string sourcepath_d = System.IO.Path.GetDirectoryName(exportFile.FileName) + "//" + filename;
                string DemoPath = @".\User\UAES_Demo.xlsx";
                try
                {
                    //if (!File.Exists(sourcepath)) File.Create(sourcepath);                    

                    if (!File.Exists(DemoPath))
                    {
                        if ((DevExpress.XtraEditors.XtraMessageBox.Show("无法找到Demo表格，请联系管理员",
                            "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK)
                        {
                            WaitForProgressing.CloseWaitForm();
                            return;
                        }
                    }
                }
                catch
                {
                    if ((DevExpress.XtraEditors.XtraMessageBox.Show("无法新建表格（路径和文件名有误），请重试",
                            "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK)
                    {
                        WaitForProgressing.CloseWaitForm();
                        return;
                    }
                }
                ///load excel's sheets
                try
                {
                    File.Copy(DemoPath, sourcepath, true);
                }
                catch 
                {
                    if ((DevExpress.XtraEditors.XtraMessageBox.Show("无法打开目标路径表格，表格是否已被打开，请关闭后重试",
                            "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK)
                    {
                        WaitForProgressing.CloseWaitForm();
                        return;
                    }
                }

                #region “任务”表单
                Workbook wb;
                Worksheet ws;
                Cells cell = null;
                try
                {
                    wb = new Workbook(sourcepath);
                    ws = wb.Worksheets["OPL"];
                }
                catch
                {
                    wb = new Workbook();
                    ws = wb.Worksheets[0];
                    if ((DevExpress.XtraEditors.XtraMessageBox.Show("Demo表格格式检查错误，请表Demo是否被篡改（需要存在“任务”sheet)",
                        "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK)
                    {
                        WaitForProgressing.CloseWaitForm();
                        return;
                    }
                }

                #region 创建样式
                //创建样式 for textB
                Aspose.Cells.Style txtStyle = HyqCtrl.creatStyle(wb, System.Drawing.Color.Black, System.Drawing.Color.White, TextAlignmentType.Center);
                Aspose.Cells.Style memStyle = HyqCtrl.creatStyle(wb, System.Drawing.Color.Black, System.Drawing.Color.White, TextAlignmentType.Left);
                Aspose.Cells.Style txtStyle_d = HyqCtrl.creatStyle(wb, System.Drawing.Color.White, System.Drawing.Color.Red, TextAlignmentType.Center);
                Aspose.Cells.Style txtStyle_c = HyqCtrl.creatStyle(wb, System.Drawing.Color.Black, System.Drawing.Color.FromArgb(0,255,0), TextAlignmentType.Center);
                Aspose.Cells.Style txtStyle_o = HyqCtrl.creatStyle(wb, System.Drawing.Color.Black, System.Drawing.Color.Yellow, TextAlignmentType.Center);
                #endregion
                ///write sheet's cells
                try
                {
                    cell = ws.Cells;
                }
                catch
                {
                    if ((DevExpress.XtraEditors.XtraMessageBox.Show("导出失败",
                        "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK)
                    {
                        WaitForProgressing.CloseWaitForm();
                        return;
                    }
                }
                
                //Title set
                cell[0, 0].PutValue(this.TxtB_Title.Text);
                cell.SetRowHeight(0, 40);//设置行高

                #region 创建附件文件夹
                bool CrtSrcFlg = true;
                try
                {
                    Directory.CreateDirectory(sourcepath_d);
                }
                catch
                {
                    CrtSrcFlg = false;
                }
                #endregion


                int startrow = 2; //起始行为第3行
                string fileCombi = "";
                string[] targetFilePath1 = new string[0];
                if (ListPanel.SelectedPageIndex == 1) ListPanel.SelectedPageIndex = 0;
                for (int i = startrow; i < startrow + this.OpenListView.DataRowCount; i++)
                {
                    #region 读取附件文件名信息                    
                    try
                    {
                        string[] tmp = Func.ConvertItems(Func_LookDetailListInItem(i - startrow), new string[] { "<SPLIT>" });
                        string ItemName = tmp[0]; //读取选中第几行对应的Item尾缀

                        #region File的特殊处理
                        string targetPath = ".\\Data\\" + saveFilePath_N + "\\" + ItemName;
                        fileCombi = "";
                        #endregion

                        try
                        {
                            targetFilePath1 = Directory.GetFiles(targetPath);
                        }
                        catch { }
                        foreach (string f in targetFilePath1)
                        {
                            fileCombi = fileCombi + System.IO.Path.GetFileName(f) + Environment.NewLine;
                            filename = System.IO.Path.GetFileName(f);
                            filename = HYQReNameHelper.FileReName(sourcepath_d + "\\" + filename); //如果文件目录下已存在，则重命名，否则返回原命名
                            if (CrtSrcFlg == true) { File.Copy(System.IO.Path.GetFullPath(f), sourcepath_d + "\\" + filename); } //复制到目标目录
                        }
                        #endregion
                    }
                    catch
                    {

                    }   

                    cell[i, 0].PutValue(i - startrow + 1);
                    cell[i, 0].SetStyle(txtStyle); //添加样式 

                    cell[i, 1].PutValue(this.OpenListView.GetRowCellValue(i - startrow, this.OpenListView.Columns[1]));
                    cell[i, 1].SetStyle(txtStyle); //添加样式 

                    cell[i, 2].PutValue(this.OpenListView.GetRowCellValue(i - startrow, this.OpenListView.Columns[2]));
                    cell[i, 2].SetStyle(txtStyle); //添加样式 

                    cell[i, 3].PutValue(this.OpenListView.GetRowCellValue(i - startrow, this.OpenListView.Columns[3]));
                    cell[i, 3].SetStyle(txtStyle); //添加样式 

                    cell[i, 4].PutValue(this.OpenListView.GetRowCellValue(i - startrow, this.OpenListView.Columns[4]));
                    cell[i, 4].SetStyle(memStyle); //添加样式 

                    cell[i, 5].PutValue(this.OpenListView.GetRowCellValue(i - startrow, this.OpenListView.Columns[5]));
                    cell[i, 5].SetStyle(txtStyle); //添加样式 

                    cell[i, 6].PutValue(this.OpenListView.GetRowCellValue(i - startrow, this.OpenListView.Columns[6]).ToString().Replace(Environment.NewLine + "[2", "[2"));
                    cell[i, 6].SetStyle(memStyle); //添加样式 

                    cell[i, 7].PutValue(this.OpenListView.GetRowCellValue(i - startrow, this.OpenListView.Columns[9]));
                    if (cell[i, 7].StringValue.Trim() == "Delay") cell[i, 7].SetStyle(txtStyle_d); //添加样式 
                    if (cell[i, 7].StringValue.Trim() == "Close") cell[i, 7].SetStyle(txtStyle_c); //添加样式 
                    if (cell[i, 7].StringValue.Trim() == "Open") cell[i, 7].SetStyle(txtStyle_o); //添加样式 

                    cell[i, 8].PutValue(this.OpenListView.GetRowCellValue(i - startrow, this.OpenListView.Columns[7]));
                    cell[i, 8].SetStyle(txtStyle); //添加样式 

                    cell[i, 9].PutValue(this.OpenListView.GetRowCellValue(i - startrow, this.OpenListView.Columns[8]));
                    cell[i, 9].SetStyle(txtStyle); //添加样式 

                    cell[i, 10].PutValue(this.OpenListView.GetRowCellValue(i - startrow, this.OpenListView.Columns[11]));
                    cell[i, 10].SetStyle(txtStyle); //添加样式 

                    cell[i, 11].PutValue(fileCombi);
                    cell[i, 11].SetStyle(txtStyle); //添加样式
                }

                ws.AutoFitRows();
                #endregion

                #region “封面”表单
                try
                {
                    ws = wb.Worksheets["Cover"];
                    cell = ws.Cells;
                }
                catch
                {
                    wb = new Workbook();
                    ws = wb.Worksheets[0];
                    cell = ws.Cells;
                    if ((DevExpress.XtraEditors.XtraMessageBox.Show("Demo表格格式检查错误，请表Demo是否被篡改（需要存在“封面”sheet)",
                        "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK)
                    {
                        WaitForProgressing.CloseWaitForm();
                        return;
                    }
                }

                //Title set
                cell[1, 0].PutValue(this.TxtB_Title.Text);
                #endregion

                #region “存档”表单
                try
                {
                    ws = wb.Worksheets["Achieved OPL"];
                    cell = ws.Cells;
                }
                catch
                {
                    wb = new Workbook();
                    ws = wb.Worksheets[0];
                    cell = ws.Cells;
                    if ((DevExpress.XtraEditors.XtraMessageBox.Show("Demo表格格式检查错误，请表Demo是否被篡改（需要存在“封面”sheet)",
                        "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK)
                    {
                        WaitForProgressing.CloseWaitForm();
                        return;
                    }
                }

                //Title set
                startrow = 2; //起始行为第3行
                string[] targetFilePath2 = new string[0];
                ListPanel.SelectedPageIndex = 1;
                for (int i = startrow; i < startrow + this.ArchieveListView.DataRowCount; i++)
                {
                    #region 读取附件文件名信息                    
                    try
                    {
                        string[] tmp = Func.ConvertItems(Func_LookDetailListInItem(i - startrow), new string[] { "<SPLIT>" });
                        string ItemName = tmp[0]; //读取选中第几行对应的Item尾缀
                        fileCombi = "";
                        #region File的特殊处理
                        string targetPath = ".\\Data\\" + saveFilePath_N + "\\" + ItemName;
                        #endregion

                        try
                        {
                            targetFilePath2 = Directory.GetFiles(targetPath);
                        }
                        catch { }
                        foreach (string f in targetFilePath2)
                        {
                            fileCombi = fileCombi + System.IO.Path.GetFileName(f) + Environment.NewLine;
                            filename = System.IO.Path.GetFileName(f);
                            filename = HYQReNameHelper.FileReName(sourcepath_d + "\\" + filename); //如果文件目录下已存在，则重命名，否则返回原命名
                            if (CrtSrcFlg == true) { File.Copy(System.IO.Path.GetFullPath(f), sourcepath_d + "\\" + filename); } //复制到目标目录
                        }
                        #endregion
                    }
                    catch
                    {

                    }


                    cell[i, 0].PutValue(i - startrow + 1);
                    cell[i, 0].SetStyle(txtStyle); //添加样式 

                    cell[i, 1].PutValue(this.ArchieveListView.GetRowCellValue(i - startrow, this.ArchieveListView.Columns[1]));
                    cell[i, 1].SetStyle(txtStyle); //添加样式 

                    cell[i, 2].PutValue(this.ArchieveListView.GetRowCellValue(i - startrow, this.ArchieveListView.Columns[2]));
                    cell[i, 2].SetStyle(txtStyle); //添加样式 

                    cell[i, 3].PutValue(this.ArchieveListView.GetRowCellValue(i - startrow, this.ArchieveListView.Columns[3]));
                    cell[i, 3].SetStyle(txtStyle); //添加样式 

                    cell[i, 4].PutValue(this.ArchieveListView.GetRowCellValue(i - startrow, this.ArchieveListView.Columns[4]));
                    cell[i, 4].SetStyle(memStyle); //添加样式 

                    cell[i, 5].PutValue(this.ArchieveListView.GetRowCellValue(i - startrow, this.ArchieveListView.Columns[5]));
                    cell[i, 5].SetStyle(txtStyle); //添加样式 

                    cell[i, 6].PutValue(this.ArchieveListView.GetRowCellValue(i - startrow, this.ArchieveListView.Columns[6]).ToString().Replace(Environment.NewLine + "[2", "[2"));
                    cell[i, 6].SetStyle(memStyle); //添加样式 

                    cell[i, 7].PutValue(this.ArchieveListView.GetRowCellValue(i - startrow, this.ArchieveListView.Columns[7]));
                    if (cell[i, 7].StringValue.Trim() == "Delay") cell[i, 7].SetStyle(txtStyle_d); //添加样式 
                    if (cell[i, 7].StringValue.Trim() == "Close") cell[i, 7].SetStyle(txtStyle_c); //添加样式
                    if (cell[i, 7].StringValue.Trim() == "Open") cell[i, 7].SetStyle(txtStyle_o); //添加样式 

                    cell[i, 8].PutValue(this.ArchieveListView.GetRowCellValue(i - startrow, this.ArchieveListView.Columns[8]));
                    cell[i, 8].SetStyle(txtStyle); //添加样式 

                    cell[i, 9].PutValue(this.ArchieveListView.GetRowCellValue(i - startrow, this.ArchieveListView.Columns[9]));
                    cell[i, 9].SetStyle(txtStyle); //添加样式 

                    cell[i, 10].PutValue(this.ArchieveListView.GetRowCellValue(i - startrow, this.ArchieveListView.Columns[10]));
                    cell[i, 10].SetStyle(txtStyle); //添加样式 
                                           
                }

                ws.AutoFitRows();
                #endregion

                wb.Save(sourcepath);
                if (ListPanel.SelectedPageIndex == 1) ListPanel.SelectedPageIndex = 0;
                GC.Collect();
            
                try
                {
                    WaitForProgressing.CloseWaitForm();
                }
                catch
                { }
                    
            }


        }
        #endregion

        #region Ribbon界面函数 点击选择“导出到EXCEL表”
        private void DoExport1(object sender, ItemClickEventArgs e)
        {
            SaveFileDialog exportFile = new SaveFileDialog();
            exportFile.Title = "Export to EXCEL file";
            exportFile.FileName = saveFilePath_N;
            exportFile.InitialDirectory = @"D:\";   //@是取消转义字符的意思
            exportFile.Filter = "All Files(*.*)|*.*|Office Excel(*.xlsx)|*.xlsx|2007 Excel(*.xls)|*.xls ";
            exportFile.FilterIndex = 2;
            exportFile.RestoreDirectory = true;

            if (exportFile.ShowDialog() == DialogResult.OK)
            {
                ///开始
                WaitForProgressing.ShowWaitForm();

                string filename = System.IO.Path.GetFileNameWithoutExtension(exportFile.FileName);
                string sourcepath = System.IO.Path.GetFullPath(exportFile.FileName);
                string DemoPath = @".\User\Export_Demo.xlsx";
                try
                {
                    //if (!File.Exists(sourcepath)) File.Create(sourcepath);
                    if (!File.Exists(DemoPath))
                    {
                        if ((DevExpress.XtraEditors.XtraMessageBox.Show("无法找到Demo表格，请联系管理员",
                            "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK)
                        { 
                            WaitForProgressing.CloseWaitForm();
                            return;
                        } 
                    } 
                }
                catch
                {
                    if ((DevExpress.XtraEditors.XtraMessageBox.Show("无法新建表格（路径和文件名有误），请重试",
                            "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK)
                    {
                        WaitForProgressing.CloseWaitForm();
                        return;
                    } 
                }
                ///load excel's sheets
                File.Copy(DemoPath, sourcepath, true);
                Workbook wb;
                Worksheet ws;
                Cells cell = null;
                try
                {
                    wb = new Workbook(sourcepath);
                    ws = wb.Worksheets["Open Point List"];
                }
                catch
                {
                    wb = new Workbook();
                    ws = wb.Worksheets[0];
                    if ((DevExpress.XtraEditors.XtraMessageBox.Show("Demo表格格式检查错误，请表Demo是否被篡改（需要存在“OpenPointList”和“Achieve”sheet)",
                        "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK)
                    {
                        WaitForProgressing.CloseWaitForm();
                        return;
                    } 
                }
                
                #region 创建样式
                //创建样式 for textB
                Aspose.Cells.Style txtStyle = HyqCtrl.creatStyle(wb, System.Drawing.Color.Black, System.Drawing.Color.White, TextAlignmentType.Center);
                Aspose.Cells.Style memStyle = HyqCtrl.creatStyle(wb, System.Drawing.Color.Black, System.Drawing.Color.White, TextAlignmentType.Left);
                Aspose.Cells.Style txtStyle_d = HyqCtrl.creatStyle(wb, System.Drawing.Color.White, System.Drawing.Color.Red, TextAlignmentType.Center);
                Aspose.Cells.Style txtStyle_c = HyqCtrl.creatStyle(wb, System.Drawing.Color.White, System.Drawing.Color.Green, TextAlignmentType.Center);
                Aspose.Cells.Style txtStyle_o = HyqCtrl.creatStyle(wb, System.Drawing.Color.Black, System.Drawing.Color.Yellow, TextAlignmentType.Center);
                #endregion  
                ///write sheet's cells
                try
                {
                    cell = ws.Cells;
                }
                catch
                {
                    if ((DevExpress.XtraEditors.XtraMessageBox.Show("导出失败",
                        "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK)
                    {
                        WaitForProgressing.CloseWaitForm();
                        return;
                    } 
                }
                
                int startrow = 1; //起始行为第2行
                for (int i = startrow; i < startrow + this.mainGridView.DataRowCount; i++)
                {
                    for (int j = 1; j < lb_HeaderName.Count; j++)
                    {
                        cell[i, j].PutValue(this.mainGridView.GetRowCellValue(i - startrow, this.mainGridView.Columns[j]));
                        if ((j == 4) || (j == 6))
                        {
                            cell[i, j].SetStyle(memStyle); //添加样式 
                        }
                        else if (j == 9)
                        {
                            if (cell[i, j].StringValue.Trim() == "Delay") cell[i, j].SetStyle(txtStyle_d); //添加样式 
                            if (cell[i, j].StringValue.Trim() == "Close") cell[i, j].SetStyle(txtStyle_c); //添加样式 
                            if (cell[i, j].StringValue.Trim() == "Open") cell[i, j].SetStyle(txtStyle_o); //添加样式 
                        }
                        else
                        {
                            cell[i, j].SetStyle(txtStyle); //添加样式 
                        }
                    }
                    cell[i, 0].PutValue(i);
                }
                
                ws.AutoFitRows();
                wb.Save(sourcepath);
                GC.Collect();

                WaitForProgressing.CloseWaitForm();
            }

            
        }
        #endregion

        #region Ribbon界面函数 点击选择“从UAES经典OPL导入”
        private void DoImport1(object sender, ItemClickEventArgs e)
        {
            OpenFileDialog importFile = new OpenFileDialog();
            importFile.Title = "Open UAES classical OPL form";
            importFile.InitialDirectory = @"D:\";   //@是取消转义字符的意思
            importFile.Filter = "All Files(*.*)|*.*|Office Excel(*.xlsx)|*.xlsx|2007 Excel(*.xls)|*.xls ";
            importFile.FilterIndex = 2;
            importFile.RestoreDirectory = true;

            if (importFile.ShowDialog() == DialogResult.OK)
            {
                ///开始
                WaitForProgressing.ShowWaitForm();
                string filename = System.IO.Path.GetFileNameWithoutExtension(importFile.FileName);
                string sourcepath = System.IO.Path.GetFullPath(importFile.FileName);
                Workbook wb;
                Worksheet ws;

                string[] filespath;

                ///set new file name
                ///filename = Regex.Replace(filename, @"\s", "");
                string oldpath_N = saveFilePath_N;
                string oldpath_D = saveFilePath_D;
                saveFilePath_D = @".\Save\" + filename + ".ini";
                saveFilePath_N = filename;

                ///check wether filename already exist or not
                try { filespath = Directory.GetFiles(@".\Save\"); }
                catch { return; }
                HYQFileInfoList fileList = new HYQFileInfoList(filespath);
                foreach (FileInfoWithIcon file in fileList.list)
                {
                    if (System.IO.Path.GetFileNameWithoutExtension(file.fileInfo.Name) == filename)
                    {
                        WaitForProgressing.CloseWaitForm();
                        DialogResult result = DevExpress.XtraEditors.XtraMessageBox.Show(
                            "已经存在重复的OPL文件名，请重新输入！", "Info", MessageBoxButtons.OK, MessageBoxIcon.Question);
                        if (result == DialogResult.OK)
                        {
                            if (FileNewClick())//新建文件
                            {
                                WaitForProgressing.ShowWaitForm();
                            }
                            else
                            {
                                saveFilePath_N = oldpath_N;
                                saveFilePath_D = oldpath_D;
                                return;
                            }
                        }
                        else
                        {
                            saveFilePath_N = oldpath_N;
                            saveFilePath_D = oldpath_D;
                            return;
                        }
                    }
                }
                HyqIni.PutINI("Main", "saveFilePath_D", saveFilePath_D, saveFilePath_M);
                HyqIni.PutINI("Main", "saveFilePath_N", saveFilePath_N, saveFilePath_M);

                ///load excel's sheets
                try
                {
                    wb = new Workbook(sourcepath);
                    ws = wb.Worksheets["任务"];
                }
                catch
                {
                    wb = new Workbook();
                    ws = wb.Worksheets[0];
                    if ((DevExpress.XtraEditors.XtraMessageBox.Show("格式检查错误，请确认是否为UAES经典格式！(需要存在“任务”sheet)", "Info", MessageBoxButtons.OK, MessageBoxIcon.Question)) == DialogResult.OK) return;
                }

                ///read sheet's cells
                Cells cell = ws.Cells;
                int rowcount = cell.MaxDataRow;
                int startrow = 2; //根据Chelsea提供的表格，起始行为第三行
                string valueStr = "";
                string keyStr = "";
                string content;
                string content_last;
                int ind;
                DateTime content_date;
                string[] content_arry;
                DateTime dt;
                List<string> content_list = new List<string>();
                for (int i = startrow; i < rowcount + 1; i++) // i = 0和1 都是表头
                {
                    ///处理No
                    #region 处理No
                    content = cell[i, 0].StringValue.Trim();
                    content_last = cell[i - 1, 0].StringValue.Trim();

                    if (content == ""
                        && cell[i, 1].StringValue.Trim() == ""
                        && cell[i, 2].StringValue.Trim() == "")
                    {
                        continue;
                    }

                    if ((content.Trim() == "")
                        || (content == null)
                        || (Regex.IsMatch(content, @"^\d*[.]?\d*$") == false)) //为空、或为非uint数字
                    {
                        if ((i > startrow) && (Regex.IsMatch(content_last, @"^\d*[.]?\d*$") == true))
                        {
                            int last;
                            int.TryParse(content_last, out last);
                            valueStr = saveFilePath_N.Split('_')[0] + "_" + (last + 1).ToString(); ///如果未输入，但是上一行有值，则为上一行index+1
                        }
                        else if (i > startrow)
                        {
                            valueStr = saveFilePath_N.Split('_')[0] + "_" + i.ToString();  ///如果上一行为非数字，则为行数
                        }
                        else
                        {
                            valueStr = saveFilePath_N.Split('_')[0] + "_1";  ///如果为第一行，则为1
                        }
                    }
                    else
                    {
                        valueStr = saveFilePath_N.Split('_')[0] + "_" + content;///如果已输入数字，则为当前数字
                    }
                    keyStr = valueStr;
                    valueStr = valueStr + "<SPLIT>";
                    #endregion

                    ///处理日期
                    if (DateTime.TryParse(cell[i, 1].StringValue.Trim(), out content_date))
                    {
                        content = content_date.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        content = System.DateTime.Now.ToString("yyyy-MM-dd");
                    }
                    valueStr = valueStr + content + "<SPLIT>";

                    ///处理来源和类型
                    content = cell[i, 2].StringValue.Trim().Replace(Environment.NewLine, "; ");
                    valueStr = valueStr + content + "<SPLIT>";

                    content = cell[i, 3].StringValue.Trim().Replace(Environment.NewLine, "; ");
                    valueStr = valueStr + content + "<SPLIT>";

                    ///处理问题描述
                    content = cell[i, 4].StringValue.Trim().Replace(Environment.NewLine, "<ENTER>");
                    valueStr = valueStr + content + "<SPLIT>";

                    ///处理责任人
                    content = cell[i, 5].StringValue.Trim().Replace(Environment.NewLine, "<ENTER>");
                    valueStr = valueStr + content + "<SPLIT>";

                    ///处理行动措施
                    #region 处理行动措施
                    content = cell[i, 6].StringValue.Trim();
                    content_arry = Func.ConvertItems(content, new string[] { Environment.NewLine });

                    for (int j = 0; j < content_arry.Length; j++)
                    {
                        if (content_arry[j].ToString().IndexOf(":") == -1)
                        {
                            ind = content_arry[j].ToString().IndexOf("：");//中文字符
                        }
                        else
                        {
                            ind = content_arry[j].ToString().IndexOf(":");//英文字符
                        }

                        if (content_arry[j].ToString().Trim() == "")
                        {
                            continue;
                        }

                        if (ind != -1)
                        {
                            content = content_arry[j].Substring(0, ind);
                        }
                        else
                        {
                            content = content_arry[j].Trim();
                        }
                        string trackfirstsplit = " ";
                        if ((DateTime.TryParse(content, out dt)) && (j != 0) && (ind != -1))
                        {
                            ///为日期，不为第一项
                            content = "<TRACKSPACE>" + "[" + content + "]" + trackfirstsplit + content_arry[j].Substring(ind + 1);
                        }
                        else if (j != 0)
                        {
                            ///不为日期，不为第一项
                            content = "<ENTER>" + content_arry[j];
                        }
                        else if ((DateTime.TryParse(content, out dt) == true) && (ind != -1))
                        {
                            ///为日期，为第一项
                            content = "[" + content + "]" + trackfirstsplit + content_arry[j].Substring(ind + 1);
                        }
                        else
                        {
                            ///不为日期，为第一项
                            content = "[" + System.DateTime.Now.ToString("2018-01-01") + "]" + trackfirstsplit + content;//缺省日期为2018-01-01
                        }
                        valueStr = valueStr + content;
                        
                        
                    }
                    valueStr = valueStr + "<SPLIT>";

                    #endregion

                    ///处理原定日期
                    if (DateTime.TryParse(cell[i, 8].StringValue.Trim(), out content_date))
                    {
                        content = content_date.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        content = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
                    }
                    valueStr = valueStr + content + "<SPLIT>";

                    ///处理现定日期
                    if (DateTime.TryParse(cell[i, 9].StringValue.Trim(), out content_date))
                    {
                        content = content_date.ToString("yyyy-MM-dd");
                    }
                    else 
                    {
                        //keey last content(original due)
                    }
                    valueStr = valueStr + content + "<SPLIT>";

                    ///处理状态
                    //content = cell[i, 7].StringValue.Trim();
                    //int status;
                    //if (int.TryParse(content, out status))
                    //{
                    //    if (status == 1)
                    //    {
                    //        content = "Close";
                    //    }
                    //    else
                    //    {
                    //        content = "Open";
                    //    }
                    //}
                    //else
                    //{
                    //    content = "Open";
                    //}
                    if (DateTime.TryParse(cell[i, 10].StringValue.Trim(), out content_date))
                    {
                        content = "Close";
                    }
                    else
                    {
                        content = "Open";
                    }
                    valueStr = valueStr + content + "<SPLIT>";

                    ///处理完成时间
                    if (DateTime.TryParse(cell[i, 10].StringValue.Trim(), out content_date))
                    {
                        content = content_date.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        content = "";
                    }
                    valueStr = valueStr + content + "<SPLIT>";
                    HyqIni.PutINI("OPL", keyStr, valueStr, saveFilePath_D);
                }
                WaitForProgressing.CloseWaitForm();
            }
            else
            {
                return;
            }
            ///结束
            
            PowerMainForm_RefreshLoad();            
            OpenListView_GetDetailIni();
        }
        #endregion

        #region Ribbon界面函数 保存之“另存为”
        private void DoFileSave1(object sender, ItemClickEventArgs e)
        {
            string oldPath = saveFilePath_D;
            string oldName = saveFilePath_N;
            FileInfo fi = new FileInfo(oldPath);
            ///下述内容雷同DoAddFile
            if (SubForm1 != null)
            {
                SubForm1.Activate();
                SubForm1.WindowState = FormWindowState.Normal;
            }
            if (SubForm1.ShowDialog() == DialogResult.OK)
            {
                VisibleIconReset();
                ///DoAddOPL_ItemClick(null, null);
                fi.CopyTo(saveFilePath_D); //复制

                #region 更换重命名对象下存储的文件夹
                string[] tmpPath = new string[0];
                try
                {
                    FileInfo fi0 = new FileInfo(@".\Data\" + oldName);
                    fi0.MoveTo(@".\Data\" + saveFilePath_N);
                    //tmpPath = Directory.GetDirectories(@".\Data\" + saveFilePath_N);
                }
                catch { }
                #endregion

                PowerMainForm_RefreshLoad();
                OpenListView_GetDetailIni();
                HyqIni.PutINI("Main", "saveFilePath_D", saveFilePath_D, saveFilePath_M);
                HyqIni.PutINI("Main", "saveFilePath_N", saveFilePath_N, saveFilePath_M);

            }
        }

        #endregion

        #region Ribbon界面函数 保存之“重命名”
        private void DoFileSave2(object sender, ItemClickEventArgs e)
        {
            string oldPath = saveFilePath_D;
            string oldName = saveFilePath_N;
            FileInfo fi = new FileInfo(oldPath);
            ///下述内容雷同DoAddFile
            if (SubForm1 != null)
            {
                SubForm1.Activate();
                SubForm1.WindowState = FormWindowState.Normal;
            }
            if (SubForm1.ShowDialog() == DialogResult.OK)
            {
                VisibleIconReset();
                ///DoAddOPL_ItemClick(null, null);
                fi.MoveTo(saveFilePath_D);//剪切

                #region 更换重命名对象下存储的文件夹
                string[] tmpPath = new string[0];
                try
                {
                    FileInfo fi0 = new FileInfo(@".\Data\" + oldName);
                    fi0.MoveTo(@".\Data\" + saveFilePath_N);
                    //tmpPath = Directory.GetDirectories(@".\Data\" + saveFilePath_N);
                }
                catch { }
                //foreach (string path in tmpPath)
                //{
                //    FileInfo fi1 = new FileInfo(path);
                //    try
                //    {
                //        string newName = fi1.Name.Split('_')[0].Replace(oldName, saveFilePath_N) + "_" + fi1.Name.Split('_')[1];
                //        fi1.MoveTo(@".\Data\" + saveFilePath_N + "\\" + newName);
                //    }
                //    catch { }
                //}
                #endregion

                PowerMainForm_RefreshLoad();
                OpenListView_GetDetailIni();
                HyqIni.PutINI("Main", "saveFilePath_D", saveFilePath_D, saveFilePath_M);
                HyqIni.PutINI("Main", "saveFilePath_N", saveFilePath_N, saveFilePath_M);

                
            }
        }

        #endregion

        #region Ribbon界面函数 保存之“删除”
        private void DoDeleteFile(object sender, ItemClickEventArgs e)
        {
            string oldPath = saveFilePath_D;
            FileInfo fi = new FileInfo(oldPath);
            ///下述内容雷同DoAddFile
            DialogResult result = DevExpress.XtraEditors.XtraMessageBox.Show("是否删除Lion Project [" + saveFilePath_N + "]?", "Info", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (result == DialogResult.OK)
            {
                fi.Delete();//(saveFilePath_D);//删除
                saveFilePath_N = "日常记录";
                saveFilePath_D = @".\Save\日常记录.ini";
                this.FileLoad.ClearLinks();
                RecallStack.ClearStack();
                string[] filespath;
                try { filespath = Directory.GetFiles(@".\Save\"); }
                catch { return; }
                HYQFileInfoList fileList = new HYQFileInfoList(filespath);
                foreach (FileInfoWithIcon file in fileList.list)
                {
                    saveFilePath_N = System.IO.Path.GetFileNameWithoutExtension(file.fileInfo.Name);
                    saveFilePath_D = @".\Save\" + saveFilePath_N + ".ini";
                }
                VisibleIconReset();
                ///DoAddOPL_ItemClick(null, null);
                
                PowerMainForm_RefreshLoad();
                OpenListView_GetDetailIni();
                HyqIni.PutINI("Main", "saveFilePath_D", saveFilePath_D, saveFilePath_M);
                HyqIni.PutINI("Main", "saveFilePath_N", saveFilePath_N, saveFilePath_M);
            }
        }

        #endregion

        #region Ribbon界面函数 撤销案件
        private void DoUndo_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (RecallStack.CanUndo() == false) return;
            string str_old = RecallStack.Undo();
            //string str_new = RecallStack.Redo();
            ///现判断修改前内容
            if (str_old != null)
            {
                VisibleIconReset();
                int pos = str_old.IndexOf("=");
                string strTail = str_old.Substring(pos + 1);
                string strHead = str_old.Substring(0, pos);

                ///更新被撤销的内容  item -> 
                ControlUpdate(strTail);

                ///更新被撤销的内容  control -> dgv
                //foreach (Control ctrl in lsCtrl_AutoSave)
                //{
                //    UpdateItem(ctrl, null);//UpdateItem(TxtB_Plan, null);
                //}

                ///更新被撤销的内容  control -> file
                string sec;
                if (ListPanel.SelectedPageIndex == 0)
                {
                    sec = "OPL";
                }
                else
                {
                    sec = "CPL";
                }
                if (strTail == "") strTail = null;
                HyqIni.PutINI(sec, strHead, strTail, saveFilePath_D);
                recallFlag = true;
                this.mainGridView.FocusedRowHandle = -1;
            }

            OpenListView_GetDetailIni();
        }
        #endregion

        #region Ribbon界面函数 分类 DoSorting
        private void OpenListView_CustomColumnSort(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnSortEventArgs e)
        {
            DevExpress.XtraGrid.Columns.GridColumn targetCol;

            if (sort1Flag == true)
            {
                targetCol = this.mainGridView.Columns[10];
            }
            else
            {
                targetCol = this.mainGridView.Columns[0];
            }

            if (e.Column == targetCol)
            {
                e.Handled = true;
                int i1 = 0, i2 = 0;
                if ((e.Value1 != null) & (e.Value2 != null))
                {
                    int.TryParse(e.Value1.ToString(), out i1);
                    int.TryParse(e.Value2.ToString(), out i2);
                }
                else if ((e.Value1 != null) & (e.Value2 == null))
                {
                    e.Result = -1;
                    return;
                }
                else if ((e.Value1 == null) & (e.Value2 != null))
                {
                    e.Result = 1;
                    return;
                }
                else
                {
                    e.Result = 1;
                    return;
                }
                if (i1 >= i2)
                {
                    e.Result = 1;
                }
                else
                {
                    e.Result = -1;
                    //e.Result = System.Collections.Comparer.Default.Compare(i1, i2);
                }
            }
        }

        private void DoSorting1Check_CheckedChanged(object sender, EventArgs e)
        {
            bool sortByPrio;
            if (((DevExpress.XtraEditors.CheckEdit)(sender)).Checked)
            {
                sortByPrio = true;
            }
            else
            {
                sortByPrio = false;
            }
            DoSorting1_Carryout(sortByPrio);
        }

        private void DoSorting1_Carryout(bool sortByPrio)
        {
            ///清空Filter
            this.mainGridView.ActiveFilterCriteria = null;
            this.mainGridView.ClearSorting();

            if (sortByPrio)
            {
                sort1Flag = true;
                this.mainGridView.Columns[10].SortMode = DevExpress.XtraGrid.ColumnSortMode.Custom;
                this.mainGridView.Columns[10].SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
            }
            else
            {
                sort1Flag = false;
                this.mainGridView.Columns[0].SortMode = DevExpress.XtraGrid.ColumnSortMode.Custom;
                this.mainGridView.Columns[0].SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
            }
            this.mainGridView.BeginSort();
            this.mainGridView.EndSort();
        }
        #endregion

        #region OpenList界面函数 鼠标拖动多选Openlist
        private void mainGridListView_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isMouseDown = true;
            }
        }

        private void mainGridListView_MouseMove(object sender, MouseEventArgs e)
        {
            if (isMouseDown)
            {
                GridHitInfo info = mainGridView.CalcHitInfo(e.X, e.Y);
                //如果鼠标落在单元格里
                if (info.InRow)
                {
                    if (!isSetStartRow)
                    {
                        StartRowHandle = info.RowHandle;
                        isSetStartRow = true;
                    }
                    else
                    {
                        //获得当前的单元格
                        int newRowHandle = info.RowHandle;
                        if (CurrentRowHandle != newRowHandle)
                        {
                            CurrentRowHandle = newRowHandle;
                            //选定 区域 单元格
                            SelectRows(StartRowHandle, CurrentRowHandle);
                        }
                    }
                }
            }
        }

        private void mainGridListView_MouseUp(object sender, MouseEventArgs e)
        {
            StartRowHandle = -1;
            CurrentRowHandle = -1;
            isMouseDown = false;
            isSetStartRow = false;
        }

        private void SelectRows(int startRow, int endRow)
        {
            if (startRow > -1 && endRow > -1)
            {
                mainGridView.BeginSelection();
                mainGridView.ClearSelection();
                mainGridView.SelectRange(startRow, endRow);
                mainGridView.EndSelection();
            }
        }
        #endregion

        #region OpenList界面函数 自动记录序号
        private void mainGridView_RowCountChanged(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gv = (DevExpress.XtraGrid.Views.Grid.GridView)sender;
            if (gv.RowCount < 10)
                gv.IndicatorWidth = 30;
            else if (gv.RowCount < 100)
                gv.IndicatorWidth = 50;
            else if (gv.RowCount < 1000)
                gv.IndicatorWidth = 70;
            else 
                gv.IndicatorWidth = 90;

        }
        protected void mainGridView_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
           
        }
        #endregion

        #region 界面函数 禁止修改TITLE和NO
        private void TxtB_Title_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = false;
        }

        private void TxtB_No_Fake_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = false;
        }
        #endregion

        #region 界面函数 右键CMS菜单 - 选择FileList并打开浏览器路径
        private void CMS_FileList_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if ((e.ClickedItem).Name == "CMS_FileList_OpenFile")
            {
                string curPath = Directory.GetCurrentDirectory();
                string targetPath = curPath + ".\\Data\\" + saveFilePath_N + "\\" + TxtB_No.Text;
                System.Diagnostics.Process.Start("explorer.exe", targetPath);
            }
        }
        #endregion

        #region 界面函数 最大化过程中修改Item
        private void PowerMainForm_SizeChanged(object sender, EventArgs e)
        {
            if (!IniFlag)
            {
                OpenListView_ColumnsWidthEdit();
                OpenListView_CollectionWidthEdit();
            }

        }

        #endregion

    }
}