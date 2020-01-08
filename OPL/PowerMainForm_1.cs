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
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraBars.Ribbon;
using DevExpress.Images;
using DevExpress.Data.Filtering;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

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
        HYQUndoStack RecallStack = new HYQUndoStack();
        bool beginMove = false;
        bool newTrackFlag = false;
        bool IniFlag = false;
        bool recallFlag = false;
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

        #region 主函数 Load
        private void PowerMainForm_Load(object sender, EventArgs e)
        {

            //Initializing

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
            RegisterHotKey(Handle, 104, KeyModifiers.Ctrl, Keys.Z);
            #endregion

            ///主程序
            InitialAtOpening(false);

            IniFlag = false;

            #region DevExpress Ribon参考代码
            ////添加Page 
            //DevExpress.XtraBars.Ribbon.RibbonPage ribbonPage = new RibbonPage(); 
            //ribbonControl.Pages.Add(ribbonPage); 
            ////添加Group 
            //DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup = new RibbonPageGroup(); 
            //ribbonPage.Groups.Add(ribbonPageGroup); 
            ////添加Button
            //DevExpress.XtraBars.BarButtonItem barButtonItem = new BarButtonItem(); 
            //ribbonPageGroup.ItemLinks.Add(barButtonItem); 
            ////添加barSubItem
            //DevExpress.XtraBars.BarSubItem barSubItem = new BarSubItem();
            //ribbonPageGroup.ItemLinks.Add(barSubItem); 
            ////barSubItem下添加Button 
            //barSubItem.AddItem(barButtonItem); 
            ////奇异的删除Page问题( DevExpress使用技巧)
            //while (this.ribbonControl.Pages.Count > 0) 
            //{ ribbonControl.Pages.Remove(ribbonControl.Pages[0]); 
            //    //调试正常，运转报异常 
            //} 
            //while (this.ribbonControl.Pages.Count > 0) 
            //{ ribbonControl.SelectedPage = ribbonControl.Pages[0]; 
            //    ribbonControl.Pages.Remove(ribbonControl.SelectedPage); //运转正常
            //} 
            ////遏止F10键Tips (DevExpress使用技巧) 
            //ribbonControl.Manager.UseF10KeyForMenu = false; 
            ////DX按钮 ApplicationIcon属性改动图标右键 Add ApplicationMenu 添加
            //DevExpress.XtraBars.Ribbon.ApplicationMenu;
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
                        case 104: //按下的是Ctrl + Z
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
                                HyqIni.PutINI(sec, strHead, strTail, saveFilePath_D); 
                                recallFlag = true;
                                this.mainGridView.FocusedRowHandle = -1;
                            }
                            
                            OpenListView_GetDetailIni();
                            break;
                    }
                    break;
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
            int defWidth = 1200;
            int defHeight = 751;
            this.Width = defWidth;
            this.Height = defHeight;
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
            lb_HeaderName.Add("创建日期" + Environment.NewLine + "Creat Date");        ///2
            lb_HeaderName.Add("来源" + Environment.NewLine + "Source");            ///3
            lb_HeaderName.Add("类型" + Environment.NewLine + "Type");            ///4
            lb_HeaderName.Add("问题描述" + Environment.NewLine + "Description");        ///5
            lb_HeaderName.Add("负责人" + Environment.NewLine + "Responsible");          ///6
            lb_HeaderName.Add("行动措施" + Environment.NewLine + "Action");        ///7
            lb_HeaderName.Add("原定计划日期" + Environment.NewLine + "Original Due");    ///8
            lb_HeaderName.Add("现定计划日期" + Environment.NewLine + "Current Due");    ///9
            lb_HeaderName.Add("状态" + Environment.NewLine + "Status");            ///10   
            lb_HeaderName.Add("到期" + Environment.NewLine + "Remain");            ///11   

            /*表头集合*/
            lb_HeaderName2.Add("序号" + Environment.NewLine + "No");            ///1
            lb_HeaderName2.Add("创建日期" + Environment.NewLine + "Creat Date");        ///2
            lb_HeaderName2.Add("来源" + Environment.NewLine + "Source");            ///3
            lb_HeaderName2.Add("类型" + Environment.NewLine + "Type");            ///4
            lb_HeaderName2.Add("问题描述" + Environment.NewLine + "Description");        ///5
            lb_HeaderName2.Add("负责人" + Environment.NewLine + "Responsible");          ///6
            lb_HeaderName2.Add("行动措施" + Environment.NewLine + "Action");        ///7
            lb_HeaderName2.Add("状态" + Environment.NewLine + "Status");            ///10 
            lb_HeaderName2.Add("原定计划日期" + Environment.NewLine + "Original Due");    ///8
            lb_HeaderName2.Add("现定计划日期" + Environment.NewLine + "Current Due");    ///9
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
            ///Font
            DevExpress.Utils.AppearanceObject.DefaultFont = new Font("STXihei", 12);
            foreach (Control ctrl in lsCtrl_Visible)
            {
                ctrl.Font = new Font("STXihei", 12);
            }

            ///ToolTip
            //HyqCtrl.NewToolTip(this.Lab_No, "问题序号，自动生成，不可编辑");
            //HyqCtrl.NewToolTip(this.Lab_Source, "问题来源，例如“客户”、“售后”");
            //HyqCtrl.NewToolTip(this.Lab_Type, "问题类型，例如“软件”、“硬件”");
            ////HyqCtrl.NewToolTip(this.TxtB_Date, "问题创建日期，自动生成，不可编辑");
            //HyqCtrl.NewToolTip(this.Lab_EDate, "原定关闭日期，只可以填写一次");
            //ToolTip Ctrl = HyqCtrl.NewToolTip(this.Lab_EDate2, "现定关闭日期，可以多次填写");
            //Ctrl.Draw += new DrawToolTipEventHandler(Ctrl_Draw);

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
            
            this.OpenListView.RowCellStyle += new DevExpress.XtraGrid.Views.Grid.RowCellStyleEventHandler(OpenListView_RowCellStyle);
            this.ArchieveListView.RowCellStyle += new DevExpress.XtraGrid.Views.Grid.RowCellStyleEventHandler(OpenListView_RowCellStyle);
        }


        #endregion

        #region 0-1界面函数 mainGrid_Initial()
        /// <summary>
        /// 修改Openlist界面参数
        /// </summary>
        private void mainGrid_Initial()
        {
            this.mainGrid.Font = new Font("STXihei", 11, FontStyle.Regular);
            

            ///DateGridView
            this.mainGridView.PopulateColumns(); //显示gridCOntrol　数据
            this.mainGridView.BestFitColumns();
            this.mainGridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.mainGridView.Appearance.HeaderPanel.Font = new Font("STXihei", 11, FontStyle.Regular);
            this.mainGridView.OptionsView.RowAutoHeight = true; //自动设置行高
            this.mainGridView.OptionsView.ColumnHeaderAutoHeight = DevExpress.Utils.DefaultBoolean.True;
            //this.mainGridView.OptionsView.RowsDefaultCellStyle = true;
            //this.mainGrid.UseEmbeddedNavigator = false;  //隐藏导航栏
            //this.mainGridView.OptionsView.AllowCellMerge = true; //允许自动合并单元格
            //this.mainGridView.OptionsBehavior.Editable = false; //允许用户修改
            this.mainGridView.OptionsCustomization.AllowSort = false;

            this.mainGridView.Appearance.EvenRow.BackColor = Color.FromArgb(150, 237, 243, 254);
            this.mainGridView.Appearance.OddRow.BackColor = Color.White;
            this.mainGridView.Appearance.OddRow.ForeColor = Color.Black;
            this.mainGridView.Appearance.EvenRow.ForeColor = Color.Black;
            this.mainGridView.Appearance.Row.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.mainGridView.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
        }
        #endregion

        #region 0-2界面函数 PowerMainForm_RefreshLoad()
        /// <summary>
        /// 修改Openlist界面参数
        /// </summary>
        private void PowerMainForm_RefreshLoad()
        {
            FileLoad.ClearLinks();

            string[] filespath;
            try { filespath = Directory.GetFiles(saveFilePath); }
            catch { return; }
            HYQFileInfoList fileList = new HYQFileInfoList(filespath);
            foreach (FileInfoWithIcon file in fileList.list)
            {
                //if (file.fileInfo.Name.EndsWith("_(Archieve).ini"))
                //{
                //    continue;
                //}
                DevExpress.XtraBars.BarButtonItem bt = new BarButtonItem();
                bt.Caption = file.fileInfo.Name.Split('.')[0];
                if (saveFilePath_N == bt.Caption)
                {
                    bt.ItemAppearance.SetFont(new Font("STXihei", 11, FontStyle.Bold));
                    bt.LargeGlyph = DevExpress.Images.ImageResourceCache.Default.
                    GetImageById("BOReport2", DevExpress.Utils.Design.ImageSize.Size32x32, DevExpress.Utils.Design.ImageType.Colored);
                }
                else
                {
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
                HyqDG.Dev_EditCol(this.OpenListView.Columns[0], lb_HeaderName[0], false);  //修改列头具体参数
                HyqDG.Dev_EditCol(this.OpenListView.Columns[1], lb_HeaderName[1], false);
                HyqDG.Dev_EditCol(this.OpenListView.Columns[2], lb_HeaderName[2], false);
                HyqDG.Dev_EditCol(this.OpenListView.Columns[3], lb_HeaderName[3], false);
                HyqDG.Dev_EditCol(this.OpenListView.Columns[4], lb_HeaderName[4], true);
                HyqDG.Dev_EditCol(this.OpenListView.Columns[5], lb_HeaderName[5], false);
                HyqDG.Dev_EditCol(this.OpenListView.Columns[6], lb_HeaderName[6], true);
                HyqDG.Dev_EditCol(this.OpenListView.Columns[7], lb_HeaderName[7], false);
                HyqDG.Dev_EditCol(this.OpenListView.Columns[8], lb_HeaderName[8], false);
                HyqDG.Dev_EditCol(this.OpenListView.Columns[9], lb_HeaderName[9], false);
                HyqDG.Dev_EditCol(this.OpenListView.Columns[10], lb_HeaderName[10], false);
            }
            else 
            {
                ///已归档OPL
                foreach (string str in lb_HeaderName2)
                {
                    dt.Columns.Add(str);  ///增加各个列头
                }
                this.mainGrid.DataSource = dt;
                HyqDG.Dev_EditCol(this.ArchieveListView.Columns[0], lb_HeaderName2[0], false);  //修改列头具体参数
                HyqDG.Dev_EditCol(this.ArchieveListView.Columns[1], lb_HeaderName2[1], false);
                HyqDG.Dev_EditCol(this.ArchieveListView.Columns[2], lb_HeaderName2[2], false);
                HyqDG.Dev_EditCol(this.ArchieveListView.Columns[3], lb_HeaderName2[3], false);
                HyqDG.Dev_EditCol(this.ArchieveListView.Columns[4], lb_HeaderName2[4], true);
                HyqDG.Dev_EditCol(this.ArchieveListView.Columns[5], lb_HeaderName2[5], false);
                HyqDG.Dev_EditCol(this.ArchieveListView.Columns[6], lb_HeaderName2[6], true);
                HyqDG.Dev_EditCol(this.ArchieveListView.Columns[7], lb_HeaderName2[7], false);
                HyqDG.Dev_EditCol(this.ArchieveListView.Columns[8], lb_HeaderName2[8], false);
                HyqDG.Dev_EditCol(this.ArchieveListView.Columns[9], lb_HeaderName2[9], false);
                HyqDG.Dev_EditCol(this.ArchieveListView.Columns[10], lb_HeaderName2[10], false);
            }
            
        }
        #endregion

        #region 0-4界面函数 OpenListView_SetCallback()
        private void OpenListView_SetCallback()
        {
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
            WaitForProgressing.ShowWaitForm();

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

            ///清空Filter
            this.mainGridView.ActiveFilterCriteria = null;
               
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
            this.mainGridView.BestFitColumns();

            //更新表格颜色和状态
            for (int i = 0; i < this.mainGridView.DataRowCount; i++)
            {
                OpenListView_CheckDetail(i);

                //删除冗余行
                if (this.mainGridView.GetRowCellValue(i, this.mainGridView.Columns[0]).ToString().Trim() == "")
                {
                    this.mainGridView.DeleteRow(i);
                }
            }            
            WaitForProgressing.CloseWaitForm(); ;
        }

        private string ItemReplaceString2Enter(string subItem)
        {
            try
            {
                string str = subItem.Replace("<ENTER>", Environment.NewLine);
                str = str.Replace("<TRACKSPACE>", Environment.NewLine);
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
            int row = this.mainGridView.RowCount;
            int col = 8;//"现定计划日期" + Environment.NewLine + "Current Due";
            int col_s = 9;//"状态" + Environment.NewLine + "Status";
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
            try
            {
                string val = this.mainGridView.GetRowCellValue(e.RowHandle, e.Column).ToString();
                if (e.Column.FieldName == "状态" + Environment.NewLine + "Status")
                {
                    if (val == "Close")
                    {
                        e.Appearance.BackColor = System.Drawing.Color.Green;
                        e.Appearance.ForeColor = System.Drawing.Color.FloralWhite;
                    }
                    if (val == "Delay")
                    {
                        e.Appearance.BackColor = System.Drawing.Color.Red;
                        e.Appearance.ForeColor = System.Drawing.Color.FloralWhite;
                    }
                    else if (val == "Open")
                    {
                        e.Appearance.BackColor = System.Drawing.Color.Yellow;
                        e.Appearance.ForeColor = System.Drawing.Color.Black;
                    }
                }
            }
            catch { }

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
            ControlUpdate(substr);
        }

        private void ControlUpdate(string valueStr)
        {
            string[] subItems;
            subItems = Func.ConvertItems(valueStr, new string[] { "<SPLIT>" });  
            /*添加表格内容*/

            ///更新OpenList中内容
            if (ListPanel.SelectedPageIndex == 0)
            {
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
            }

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
            try { content = TxtB_Track.Items[SelectionTRK].ToString(); }
            catch { return; }
            
            try
            {
                int i = content.IndexOf("]") + 1;
                TxtB_Plan.Text = content.Substring(i + 1).Replace("<ENTER>", Environment.NewLine);//加上空格
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
                            tmpStr = tmpStr.Replace(Environment.NewLine + "[201", "<TRACKSPACE>[201");
                            valueStr = valueStr + tmpStr.Replace(Environment.NewLine, "<ENTER>") + "<SPLIT>";
                        }
                        catch
                        {
                            valueStr = valueStr + "<SPLIT>";
                        }
                    }
                }
                else
                {
                    valueStr = valueStr + ctrl.Text.Replace(Environment.NewLine, "<ENTER>") + "<SPLIT>";
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
            if (UpdateTimer_Ticks > 10) //大于100毫秒
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
                                if (str.StartsWith("[") )
                                {
                                    slist_real.Add(str + Environment.NewLine);
                                    cnt++;
                                }
                                else if (str.Trim()!="")
                                {
                                    slist_real[cnt] += str + Environment.NewLine;
                                }
                            }

                            ///判断是否为新增项
                            if (newTrackFlag == true)
                            {
                                if (IniFlag  ) { return; }
                                ///处理第row1项
                                slist_real.Add("[" + date + "] " + tmp.Text + Environment.NewLine);
                                TxtB_Track.Items.Add("[" + date + "] " + tmp.Text);
                                TxtB_Track.SetSelected(TxtB_Track.Items.Count - 1, true);
                            }
                            else
                            {
                                //非新增项
                                string sTmp = slist_real[row1].ToString();
                                if (sTmp.StartsWith("["))
                                {
                                    int posTmp = sTmp.IndexOf("]");
                                    if (tmp.Text == "<DELETE>")
                                    {
                                        slist_real.RemoveAt(row1);
                                        this.TxtB_Track.Items.RemoveAt(curSelectionTRK);
                                    }
                                    else 
                                    {
                                        slist_real[row1] = sTmp.Substring(0, posTmp + 1) + " " + tmp.Text + Environment.NewLine; //表格中的日期跟踪
                                        this.TxtB_Track.Items[row1] = slist_real[row1];
                                    }

                                }
                            }

                            foreach (string str in slist_real)
                            {
                                result += str;// +Environment.NewLine;
                            }
                            result = Func.DeleteTail(result, Environment.NewLine);
                            this.mainGridView.SetRowCellValue(row,this.mainGridView.Columns[col],result);
                        }
                        else
                        {
                            if (IniFlag) { return; }
                            if ((TxtB_No.Text.Trim() != "") && (tmp.Text !=""))
                            {
                                this.mainGridView.SetRowCellValue(row, this.mainGridView.Columns[col], "[" + date + "] " + tmp.Text);

                                try
                                {
                                    TxtB_Track.Items[0] = "[" + date + "] " + tmp.Text;
                                }
                                catch
                                {
                                    TxtB_Track.Items.Add("[" + date + "] " + tmp.Text);
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
        private void OpenList_KeyUp(object sender, KeyEventArgs e)
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
                    DoDelOPL_ItemClick(null, null);
                }
                catch { }
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
                string targetPath = curPath + ".\\Data\\" + saveFilePath_N + TxtB_No.Text;
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
        }
        #endregion

        #region Ribbon界面函数 点击界面，选择需要的File
        private void FormatLoad_Other_ItemClick(object sender, ItemClickEventArgs e)
        {

            saveFilePath_D = @".\Save\" + e.Item.Caption + ".ini";
            saveFilePath_N = e.Item.Caption;
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
            if (SubForm1 != null)
            {
                SubForm1.Activate();
                SubForm1.WindowState = FormWindowState.Normal;
            }
            if (SubForm1.ShowDialog() == DialogResult.OK)
            {
                DoAddOPL_ItemClick(null, null);
                PowerMainForm_RefreshLoad();
                OpenListView_GetDetailIni();
                HyqIni.PutINI("Main", "saveFilePath_D", saveFilePath_D, saveFilePath_M);
                HyqIni.PutINI("Main", "saveFilePath_N", saveFilePath_N, saveFilePath_M);
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
            if (dt == null)
            {
                for (int i = 0; i < 11; i++)
                {
                    dt.Columns.Add();
                }
                dt = new DataTable();
            }
            DataRow dr = dt.NewRow();
            dr[0] = (ind + 1).ToString();
            TxtB_No.Text = saveFilePath_N + "_" + (ind + 1).ToString();//("0000");
            TxtB_Date.Text = System.DateTime.Now.ToString("yyyy-MM-dd");
            dr[1] = TxtB_Date.Text;
            try { tmp = this.mainGridView.GetRowCellValue(curSelectionOPL - 2, this.mainGridView.Columns[2]).ToString(); }
            catch { tmp = "(来源)"; }
            dr[2] = tmp;
            TxtB_Source.Text = tmp;
            try { tmp = this.mainGridView.GetRowCellValue(curSelectionOPL - 2, this.mainGridView.Columns[3]).ToString(); }
            catch { tmp = "(归类)"; }
            dr[3] = tmp;
            TxtB_Type.Text = tmp;
            dr[4] = "";
            TxtB_Descrip.Text = "";
            try { tmp = this.mainGridView.GetRowCellValue(curSelectionOPL - 2, this.mainGridView.Columns[5]).ToString(); }
            catch { tmp = "(责任人)"; }
            dr[5] = tmp;
            TxtB_Due.Text = tmp;
            dr[6] = "";
            TxtB_Track.Items.Clear();
            TxtB_Plan.Text = "";
            dr[7] = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
            TxtB_EDate.Text = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
            dr[8] = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
            TxtB_EDate2.Text = System.DateTime.Now.AddDays(5).ToString("yyyy-MM-dd");
            dr[9] = "Open";
            TxtB_Status.Text = "Open";
            dr[10] = 5;
            //TxtB_ADate.Text = System.DateTime.Now.AddDays(20).ToString("d");
            dt.Rows.Add(dr);
            this.mainGrid.DataSource = dt;
            
            EditItemAndSave(TxtB_No.Text, "OPL");
            try
            {
                this.mainGridView.MoveNext();
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
                }
            }
        }
        #endregion

        #region Ribbon界面函数 点击归档OPL
        private void DoAchieve_ItemClick(object sender, ItemClickEventArgs e)
        {
            DialogResult result = DevExpress.XtraEditors.XtraMessageBox.Show("是否'归档'close项？", "Info", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (result == DialogResult.OK)
            {
                ///重新读取item
                string[] Items = HyqIni.GetItems("OPL", saveFilePath_D);
                string col_s = "状态" + Environment.NewLine + "Status";
                List<int> int_d = new List<int>();
                ///归档已选中项
                for (int i = this.mainGridView.DataRowCount - 1; i >= 0 ; i--)
                {
                    curSelectionOPL = i;
                    DetailUpdate(curSelectionOPL);
                    if(TxtB_Status.Text == "Close")
                    {
                        //保存在归档文件中
                        EditItemAndSave(TxtB_No.Text, "CPL");

                        //删除原OPL文件中的对应项
                        string Txt_No = this.mainGridView.GetRowCellValue(i, this.mainGridView.Columns[0]).ToString();
                        int ind_Item = Func.FindItemInIni(Txt_No, Items);
                        string keystr = Func.CalItemsHead(Items, ind_Item);
                        string valuestr = null;
                        saveIntoFile("OPL", keystr, valuestr, saveFilePath_D);
                        int_d.Add(i);
                    }
                }

                foreach (int i2 in int_d)
                {
                    this.mainGridView.DeleteRow(i2);
                }
            }
        }
        #endregion

        #region Ribbon界面函数 点击关闭按钮
        private void DoClose_ItemClick(object sender, ItemClickEventArgs e)
        {
            if ((TxtB_No.Text != "") && (this.mainGridView.SelectedRowsCount > 0))
            {
                DialogResult result = DevExpress.XtraEditors.XtraMessageBox.Show("是否'关闭'选中项？", "Info", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (result == DialogResult.OK)
                {
                    ///重新读取item
                    string[] Items = HyqIni.GetItems("OPL", saveFilePath_D);
                    string[] Items_s = HyqIni.GetItems("CPL", saveFilePath_D);
                    string col_s = "状态" + Environment.NewLine + "Status";

                    ///关闭已选中项
                    int[] rows = this.mainGridView.GetSelectedRows();
                    for (int i = 0; i < rows.Count(); i++)
                    {
                        curSelectionOPL = rows[i];
                        DetailUpdate(rows[i]);
                        TxtB_ADate.Text = System.DateTime.Now.ToString("yyyy-MM-dd");
                        TxtB_Status.Text = "Close";
                        UpdateItem(TxtB_Status, null);

                        //保存在归档文件中
                        EditItemAndSave(TxtB_No.Text, "OPL");
                    }
                }
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
            }
            else
            {
                ///已归档OPL
                this.mainGrid = this.ArchieveList;
                this.mainGridView = this.ArchieveListView;

                DoAddOPL.Enabled = false;
                DoAddPlan.Enabled = false;
                DoDelPlan.Enabled = false;
                DoClose.Enabled = false;
                DoAchieve.Enabled = false;
                mainGrid_Initial();
            }

            OpenListView_ColumnsHeader_Initial();
            OpenListView_GetDetailIni();
        }
        #endregion

        private void barButtonItem1_ItemClick(object sender, ItemClickEventArgs e)
        {
            this.Close();
        }

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


    }
}