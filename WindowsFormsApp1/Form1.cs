using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Word =  Microsoft.Office.Interop.Word;

using NAudio.Wave;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Linq;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Text;

using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Media;
using Timer = System.Windows.Forms.Timer;

namespace WindowsFormsApp1
{



    public partial class Form1 : Form
    {
        private bool isAlarmActive = false; // Flag to track if the alarm has already sounded
        
        string fontFilePath;
        string filePathData;
        private object lastSelectedItem;
        string pressedChar;
        string wordName;
        string wordPath;
        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
        int projecktNummber;
        string DocumentsPath;
        static int newVoice = 0;
        //static int newFile = 0;
        private WaveInEvent waveIn;
        private WaveFileWriter writer;
        string soundMappingFiles;
        private IWavePlayer waveOutDevice;
        private AudioFileReader audioFileReader;
        private Dictionary<char, string> soundMappings;
        private Button saveBtn;
        string textBoxValue;
        //string projektName;
        //string projektPath;
        private string outputFilePath;
        string mainFolderPath;
        string wordFilePath;
        private string mainFolder = "subDatas";
        private Timer beepTimer;
        private string[] subfolder = { "voices", "datafiles", "sound mapping", "pdf folder" };
        List<string> allDocxFiles = new List<string>();
        private List<string> dataList = new List<string>();
        private List<string> alphaList = new List<string>();
        private List<string> nameList = new List<string>();
        private List<string> voiceTagList = new List<string>();
        private List<string> dateList = new List<string>();
        private List<string> fontList = new List<string>();
        //02D1
        private AudioFileReader audioFile;
        private WaveOutEvent outputDevice;

        public Form1()
        {
            this.KeyPreview = true;
            InitializeComponent();
            PopulateFontComboBox();
            InitializeTimer();
            InitializeSoundMappings();
            this.textBox1.PreviewKeyDown += new PreviewKeyDownEventHandler(textBox1_PreviewKeyDown);
           
        }

        private void InitializeTimer() //this make a timer for alarm group
        {
            beepTimer = new Timer();
            beepTimer.Interval = 2000; // 2 seconds
            beepTimer.Tick += BeepTimer_Tick;
            
        }
        


        private void InitializeComponent()
        {
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.saveBtn = new System.Windows.Forms.Button();
            this.searchBtn = new System.Windows.Forms.Button();
            this.startparBtn = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.endparBtn = new System.Windows.Forms.Button();
            this.startsentBtn = new System.Windows.Forms.Button();
            this.bookMarkBtn = new System.Windows.Forms.Button();
            this.endsentBtn = new System.Windows.Forms.Button();
            this.starttextBtn = new System.Windows.Forms.Button();
            this.exitBtn = new System.Windows.Forms.Button();
            this.endtextBtn = new System.Windows.Forms.Button();
            this.gotoBtn = new System.Windows.Forms.Button();
            this.gotoBmark = new System.Windows.Forms.Button();
            this.insertBtn = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.fontSizeComboBox = new System.Windows.Forms.ComboBox();
            this.fontComboBox = new System.Windows.Forms.ComboBox();
            this.savePdfBtn = new System.Windows.Forms.Button();
            this.printBtn = new System.Windows.Forms.Button();
            this.saveAsBtn = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.statusBtn = new System.Windows.Forms.Button();
            this.fileManagment = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.PBObtn = new System.Windows.Forms.Button();
            this.changeColor = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.BIUbtn = new System.Windows.Forms.Button();
            this.fontGroupBtn = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.stopBtn = new System.Windows.Forms.Button();
            this.playBtn = new System.Windows.Forms.Button();
            this.skipBtn = new System.Windows.Forms.Button();
            this.settingBtn = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Arial Narrow", 19.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.ForeColor = System.Drawing.SystemColors.Desktop;
            this.textBox1.Location = new System.Drawing.Point(250, 150);
            this.textBox1.Margin = new System.Windows.Forms.Padding(100);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.textBox1.Size = new System.Drawing.Size(676, 450);
            this.textBox1.TabIndex = 0;
            this.textBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // saveBtn
            // 
            this.saveBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.saveBtn.ImageAlign = System.Drawing.ContentAlignment.TopRight;
            this.saveBtn.Location = new System.Drawing.Point(195, 0);
            this.saveBtn.Name = "saveBtn";
            this.saveBtn.Size = new System.Drawing.Size(100, 50);
            this.saveBtn.TabIndex = 1;
            this.saveBtn.Text = "ذخیره";
            this.saveBtn.UseVisualStyleBackColor = true;
            this.saveBtn.Click += new System.EventHandler(this.saveBtn_Click);
            // 
            // searchBtn
            // 
            this.searchBtn.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.searchBtn.Location = new System.Drawing.Point(176, 56);
            this.searchBtn.Name = "searchBtn";
            this.searchBtn.Size = new System.Drawing.Size(101, 57);
            this.searchBtn.TabIndex = 2;
            this.searchBtn.Text = "جستجو";
            this.searchBtn.UseVisualStyleBackColor = true;
            // 
            // startparBtn
            // 
            this.startparBtn.Location = new System.Drawing.Point(229, 7);
            this.startparBtn.Name = "startparBtn";
            this.startparBtn.Size = new System.Drawing.Size(73, 88);
            this.startparBtn.TabIndex = 3;
            this.startparBtn.Text = "شروع پاراگراف";
            this.startparBtn.UseVisualStyleBackColor = true;
            this.startparBtn.Visible = false;
            this.startparBtn.Click += new System.EventHandler(this.startparBtn_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(227, 101);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 58);
            this.button4.TabIndex = 4;
            this.button4.Text = "button4";
            this.button4.UseVisualStyleBackColor = true;
            // 
            // endparBtn
            // 
            this.endparBtn.Location = new System.Drawing.Point(102, 111);
            this.endparBtn.Name = "endparBtn";
            this.endparBtn.Size = new System.Drawing.Size(97, 48);
            this.endparBtn.TabIndex = 5;
            this.endparBtn.Text = "پایان پاراگراف";
            this.endparBtn.UseVisualStyleBackColor = true;
            this.endparBtn.Visible = false;
            this.endparBtn.Click += new System.EventHandler(this.endparBtn_Click);
            // 
            // startsentBtn
            // 
            this.startsentBtn.Location = new System.Drawing.Point(135, 72);
            this.startsentBtn.Name = "startsentBtn";
            this.startsentBtn.Size = new System.Drawing.Size(75, 40);
            this.startsentBtn.TabIndex = 6;
            this.startsentBtn.Text = "شروع جمله";
            this.startsentBtn.UseVisualStyleBackColor = true;
            this.startsentBtn.Visible = false;
            this.startsentBtn.Click += new System.EventHandler(this.startsentBtn_Click);
            // 
            // bookMarkBtn
            // 
            this.bookMarkBtn.Location = new System.Drawing.Point(95, 42);
            this.bookMarkBtn.Name = "bookMarkBtn";
            this.bookMarkBtn.Size = new System.Drawing.Size(103, 55);
            this.bookMarkBtn.TabIndex = 7;
            this.bookMarkBtn.Text = "بوک مارک";
            this.bookMarkBtn.UseVisualStyleBackColor = true;
            this.bookMarkBtn.Visible = false;
            this.bookMarkBtn.Click += new System.EventHandler(this.button7_Click);
            // 
            // endsentBtn
            // 
            this.endsentBtn.Location = new System.Drawing.Point(6, 4);
            this.endsentBtn.Name = "endsentBtn";
            this.endsentBtn.Size = new System.Drawing.Size(55, 67);
            this.endsentBtn.TabIndex = 8;
            this.endsentBtn.Text = "پایان جمله";
            this.endsentBtn.UseVisualStyleBackColor = true;
            this.endsentBtn.Visible = false;
            this.endsentBtn.Click += new System.EventHandler(this.endsentBtn_Click);
            // 
            // starttextBtn
            // 
            this.starttextBtn.Location = new System.Drawing.Point(-4, 64);
            this.starttextBtn.Name = "starttextBtn";
            this.starttextBtn.Size = new System.Drawing.Size(68, 67);
            this.starttextBtn.TabIndex = 9;
            this.starttextBtn.Text = "شروع متن";
            this.starttextBtn.UseVisualStyleBackColor = true;
            this.starttextBtn.Visible = false;
            this.starttextBtn.Click += new System.EventHandler(this.starttextBtn_Click);
            // 
            // exitBtn
            // 
            this.exitBtn.Location = new System.Drawing.Point(0, 2);
            this.exitBtn.Name = "exitBtn";
            this.exitBtn.Size = new System.Drawing.Size(80, 65);
            this.exitBtn.TabIndex = 10;
            this.exitBtn.Text = "exit";
            this.exitBtn.UseVisualStyleBackColor = true;
            this.exitBtn.Click += new System.EventHandler(this.exitBtn_Click);
            // 
            // endtextBtn
            // 
            this.endtextBtn.Location = new System.Drawing.Point(91, 5);
            this.endtextBtn.Name = "endtextBtn";
            this.endtextBtn.Size = new System.Drawing.Size(98, 58);
            this.endtextBtn.TabIndex = 11;
            this.endtextBtn.Text = "پایان متن";
            this.endtextBtn.UseVisualStyleBackColor = true;
            this.endtextBtn.Visible = false;
            // 
            // gotoBtn
            // 
            this.gotoBtn.Location = new System.Drawing.Point(47, 0);
            this.gotoBtn.Name = "gotoBtn";
            this.gotoBtn.Size = new System.Drawing.Size(108, 50);
            this.gotoBtn.TabIndex = 12;
            this.gotoBtn.Text = "برو به ";
            this.gotoBtn.UseVisualStyleBackColor = true;
            this.gotoBtn.Click += new System.EventHandler(this.gotoBtn_Click);
            // 
            // gotoBmark
            // 
            this.gotoBmark.Location = new System.Drawing.Point(151, 78);
            this.gotoBmark.Name = "gotoBmark";
            this.gotoBmark.Size = new System.Drawing.Size(109, 56);
            this.gotoBmark.TabIndex = 13;
            this.gotoBmark.Text = "go bookmark";
            this.gotoBmark.UseVisualStyleBackColor = true;
            this.gotoBmark.Visible = false;
            this.gotoBmark.Click += new System.EventHandler(this.gotoBmark_Click);
            // 
            // insertBtn
            // 
            this.insertBtn.Location = new System.Drawing.Point(101, 72);
            this.insertBtn.Name = "insertBtn";
            this.insertBtn.Size = new System.Drawing.Size(60, 35);
            this.insertBtn.TabIndex = 17;
            this.insertBtn.Text = "درج";
            this.insertBtn.UseVisualStyleBackColor = true;
            this.insertBtn.Click += new System.EventHandler(this.button5_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.AutoSize = true;
            this.groupBox1.Controls.Add(this.fontSizeComboBox);
            this.groupBox1.Controls.Add(this.fontComboBox);
            this.groupBox1.Controls.Add(this.savePdfBtn);
            this.groupBox1.Controls.Add(this.printBtn);
            this.groupBox1.Controls.Add(this.saveBtn);
            this.groupBox1.Controls.Add(this.searchBtn);
            this.groupBox1.Controls.Add(this.bookMarkBtn);
            this.groupBox1.Controls.Add(this.gotoBtn);
            this.groupBox1.Controls.Add(this.gotoBmark);
            this.groupBox1.Location = new System.Drawing.Point(928, 9);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(295, 166);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            // 
            // fontSizeComboBox
            // 
            this.fontSizeComboBox.FormattingEnabled = true;
            this.fontSizeComboBox.Location = new System.Drawing.Point(151, 119);
            this.fontSizeComboBox.Name = "fontSizeComboBox";
            this.fontSizeComboBox.Size = new System.Drawing.Size(107, 24);
            this.fontSizeComboBox.TabIndex = 26;
            this.fontSizeComboBox.Visible = false;
            this.fontSizeComboBox.SelectedIndexChanged += new System.EventHandler(this.fontSizeComboBox_SelectedIndexChanged);
            this.fontSizeComboBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.fontSizeComboBox_keyPress);
            // 
            // fontComboBox
            // 
            this.fontComboBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.fontComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.fontComboBox.Location = new System.Drawing.Point(16, 86);
            this.fontComboBox.Name = "fontComboBox";
            this.fontComboBox.Size = new System.Drawing.Size(112, 23);
            this.fontComboBox.TabIndex = 0;
            this.fontComboBox.Visible = false;
            this.fontComboBox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.FontComboBox_DrawItem);
            this.fontComboBox.MeasureItem += new System.Windows.Forms.MeasureItemEventHandler(this.FontComboBox_MeasureItem);
            this.fontComboBox.SelectedIndexChanged += new System.EventHandler(this.FontComboBox_SelectedIndexChanged);
            this.fontComboBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ComboBox_KeyPress);
            // 
            // savePdfBtn
            // 
            this.savePdfBtn.Location = new System.Drawing.Point(101, 86);
            this.savePdfBtn.Name = "savePdfBtn";
            this.savePdfBtn.Size = new System.Drawing.Size(97, 48);
            this.savePdfBtn.TabIndex = 7;
            this.savePdfBtn.Text = "save pdf";
            this.savePdfBtn.UseVisualStyleBackColor = true;
            this.savePdfBtn.Visible = false;
            this.savePdfBtn.Click += new System.EventHandler(this.savePdfBtn_Click);
            // 
            // printBtn
            // 
            this.printBtn.Location = new System.Drawing.Point(19, 47);
            this.printBtn.Name = "printBtn";
            this.printBtn.Size = new System.Drawing.Size(76, 45);
            this.printBtn.TabIndex = 6;
            this.printBtn.Text = "print";
            this.printBtn.UseVisualStyleBackColor = true;
            this.printBtn.Visible = false;
            this.printBtn.Click += new System.EventHandler(this.printBtn_Click);
            // 
            // saveAsBtn
            // 
            this.saveAsBtn.Location = new System.Drawing.Point(85, 44);
            this.saveAsBtn.Name = "saveAsBtn";
            this.saveAsBtn.Size = new System.Drawing.Size(76, 45);
            this.saveAsBtn.TabIndex = 26;
            this.saveAsBtn.Text = "save as";
            this.saveAsBtn.UseVisualStyleBackColor = true;
            this.saveAsBtn.Visible = false;
            this.saveAsBtn.Click += new System.EventHandler(this.saveAsBtn_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.AutoSize = true;
            this.groupBox2.Controls.Add(this.statusBtn);
            this.groupBox2.Controls.Add(this.saveAsBtn);
            this.groupBox2.Controls.Add(this.fileManagment);
            this.groupBox2.Controls.Add(this.exitBtn);
            this.groupBox2.Controls.Add(this.starttextBtn);
            this.groupBox2.Controls.Add(this.endtextBtn);
            this.groupBox2.Location = new System.Drawing.Point(12, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(208, 205);
            this.groupBox2.TabIndex = 19;
            this.groupBox2.TabStop = false;
            // 
            // statusBtn
            // 
            this.statusBtn.Location = new System.Drawing.Point(-12, 129);
            this.statusBtn.Name = "statusBtn";
            this.statusBtn.Size = new System.Drawing.Size(73, 55);
            this.statusBtn.TabIndex = 27;
            this.statusBtn.Text = "status";
            this.statusBtn.UseVisualStyleBackColor = true;
            this.statusBtn.Visible = false;
            this.statusBtn.Click += new System.EventHandler(this.statusBtn_Click);
            // 
            // fileManagment
            // 
            this.fileManagment.Location = new System.Drawing.Point(40, 107);
            this.fileManagment.Name = "fileManagment";
            this.fileManagment.Size = new System.Drawing.Size(69, 47);
            this.fileManagment.TabIndex = 7;
            this.fileManagment.Text = "file magment";
            this.fileManagment.UseVisualStyleBackColor = true;
            this.fileManagment.Click += new System.EventHandler(this.fileMagament_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox3.AutoSize = true;
            this.groupBox3.Controls.Add(this.PBObtn);
            this.groupBox3.Controls.Add(this.changeColor);
            this.groupBox3.Controls.Add(this.endsentBtn);
            this.groupBox3.Controls.Add(this.startsentBtn);
            this.groupBox3.Controls.Add(this.insertBtn);
            this.groupBox3.Location = new System.Drawing.Point(12, 403);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(216, 133);
            this.groupBox3.TabIndex = 20;
            this.groupBox3.TabStop = false;
            this.groupBox3.Enter += new System.EventHandler(this.groupBox3_Enter);
            // 
            // PBObtn
            // 
            this.PBObtn.Location = new System.Drawing.Point(22, 52);
            this.PBObtn.Name = "PBObtn";
            this.PBObtn.Size = new System.Drawing.Size(73, 55);
            this.PBObtn.TabIndex = 20;
            this.PBObtn.Text = "PBO";
            this.PBObtn.UseVisualStyleBackColor = true;
            this.PBObtn.Visible = false;
            this.PBObtn.Click += new System.EventHandler(this.PBObtn_Click);
            // 
            // changeColor
            // 
            this.changeColor.Location = new System.Drawing.Point(131, 34);
            this.changeColor.Name = "changeColor";
            this.changeColor.Size = new System.Drawing.Size(73, 55);
            this.changeColor.TabIndex = 19;
            this.changeColor.Text = "color";
            this.changeColor.UseVisualStyleBackColor = true;
            this.changeColor.Visible = false;
            this.changeColor.Click += new System.EventHandler(this.changeColor_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox4.AutoSize = true;
            this.groupBox4.Controls.Add(this.BIUbtn);
            this.groupBox4.Controls.Add(this.fontGroupBtn);
            this.groupBox4.Controls.Add(this.startparBtn);
            this.groupBox4.Controls.Add(this.endparBtn);
            this.groupBox4.Controls.Add(this.button4);
            this.groupBox4.Location = new System.Drawing.Point(927, 365);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(308, 180);
            this.groupBox4.TabIndex = 21;
            this.groupBox4.TabStop = false;
            // 
            // BIUbtn
            // 
            this.BIUbtn.Location = new System.Drawing.Point(67, 79);
            this.BIUbtn.Name = "BIUbtn";
            this.BIUbtn.Size = new System.Drawing.Size(89, 48);
            this.BIUbtn.TabIndex = 7;
            this.BIUbtn.Text = "BIU";
            this.BIUbtn.UseVisualStyleBackColor = true;
            this.BIUbtn.Visible = false;
            this.BIUbtn.Click += new System.EventHandler(this.BIUbtn_Click);
            // 
            // fontGroupBtn
            // 
            this.fontGroupBtn.Location = new System.Drawing.Point(162, 24);
            this.fontGroupBtn.Name = "fontGroupBtn";
            this.fontGroupBtn.Size = new System.Drawing.Size(73, 55);
            this.fontGroupBtn.TabIndex = 6;
            this.fontGroupBtn.Text = "font";
            this.fontGroupBtn.UseVisualStyleBackColor = true;
            this.fontGroupBtn.Click += new System.EventHandler(this.fontGroupBtn_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 16;
            this.listBox1.Location = new System.Drawing.Point(944, 181);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(370, 148);
            this.listBox1.TabIndex = 22;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.ListBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(773, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 16);
            this.label1.TabIndex = 23;
            this.label1.Text = "record status";
            this.label1.Visible = false;
            // 
            // textBox2
            // 
            this.textBox2.Enabled = false;
            this.textBox2.Location = new System.Drawing.Point(250, 119);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(676, 33);
            this.textBox2.TabIndex = 25;
            this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            this.textBox2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox2_keyPress);
            // 
            // stopBtn
            // 
            this.stopBtn.Location = new System.Drawing.Point(0, 0);
            this.stopBtn.Name = "stopBtn";
            this.stopBtn.Size = new System.Drawing.Size(75, 23);
            this.stopBtn.TabIndex = 2;
            // 
            // playBtn
            // 
            this.playBtn.Location = new System.Drawing.Point(0, 0);
            this.playBtn.Name = "playBtn";
            this.playBtn.Size = new System.Drawing.Size(75, 23);
            this.playBtn.TabIndex = 1;
            // 
            // skipBtn
            // 
            this.skipBtn.Location = new System.Drawing.Point(0, 0);
            this.skipBtn.Name = "skipBtn";
            this.skipBtn.Size = new System.Drawing.Size(75, 23);
            this.skipBtn.TabIndex = 0;
            // 
            // settingBtn
            // 
            this.settingBtn.Location = new System.Drawing.Point(74, 261);
            this.settingBtn.Name = "settingBtn";
            this.settingBtn.Size = new System.Drawing.Size(73, 55);
            this.settingBtn.TabIndex = 28;
            this.settingBtn.Text = "status";
            this.settingBtn.UseVisualStyleBackColor = true;
            this.settingBtn.Click += new System.EventHandler(this.settingBtn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1231, 545);
            this.Controls.Add(this.settingBtn);
            this.Controls.Add(this.skipBtn);
            this.Controls.Add(this.playBtn);
            this.Controls.Add(this.stopBtn);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.textBox1);
            this.KeyPreview = true;
            this.Name = "Form1";
            this.Text = "gcw";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Resize += new System.EventHandler(this.Form1_Resize);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        //83899518
     

        private void InitializeSoundMappings()
        {

            soundMappings = new Dictionary<char, string>
        {
        {'ا', @"C:\Users\Artin\Documents\voice files project\alef.mp3"},
        {'ب', @"C:\Users\Artin\Documents\voice files project\be.mp3"},
        {'پ', @"C:\Users\Artin\Documents\voice files project\pe.mp3"},
        {'ت', @"C:\Users\Artin\Documents\voice files project\te.mp3"},
        {'ث', @"C:\Users\Artin\Documents\voice files project\se.mp3"},
        {'ج', @"C:\Users\Artin\Documents\voice files project\je.mp3"},
        {'چ', @"C:\Users\Artin\Documents\voice files project\che.mp3"},
        {'ح', @"C:\Users\Artin\Documents\voice files project\hhe.mp3"},
        {'خ', @"C:\Users\Artin\Documents\voice files project\khe.mp3"},
        {'د', @"C:\Users\Artin\Documents\voice files project\dal.mp3"},
        {'ذ', @"C:\Users\Artin\Documents\voice files project\dal_ze.mp3"},
        {'ر', @"C:\Users\Artin\Documents\voice files project\re.mp3"},
        {'ز', @"C:\Users\Artin\Documents\voice files project\ze.mp3"},
        {'ژ', @"C:\Users\Artin\Documents\voice files project\zhe.mp3"},
        {'س', @"C:\Users\Artin\Documents\voice files project\sse.mp3"},
        {'ش', @"C:\Users\Artin\Documents\voice files project\she.mp3"},
        {'ص', @"C:\Users\Artin\Documents\voice files project\sad.mp3"},
        {'ض', @"C:\Users\Artin\Documents\voice files project\zad.mp3"},
        {'ط', @"C:\Users\Artin\Documents\voice files project\ta.mp3"},
        {'ظ', @"C:\Users\Artin\Documents\voice files project\za.mp3"},
        {'ع', @"C:\Users\Artin\Documents\voice files project\ain.mp3"},
        {'غ', @"C:\Users\Artin\Documents\voice files project\ghain.mp3"},
        {'ف', @"C:\Users\Artin\Documents\voice files project\fe.mp3"},
        {'ق', @"C:\Users\Artin\Documents\voice files project\ghaf.mp3"},
        {'ک', @"C:\Users\Artin\Documents\voice files project\kaf.mp3"},
        {'گ', @"C:\Users\Artin\Documents\voice files project\gaf.mp3"},
        {'ل', @"C:\Users\Artin\Documents\voice files project\lam.mp3"},
        {'م', @"C:\Users\Artin\Documents\voice files project\mim.mp3"},
        {'ن', @"C:\Users\Artin\Documents\voice files project\non.mp3"},
        {'و', @"C:\Users\Artin\Documents\voice files project\ve.mp3"},
        {'ه', @"C:\Users\Artin\Documents\voice files project\he.mp3"},
        {'ی', @"C:\Users\Artin\Documents\voice files project\ye.mp3"},
        {'إ', @"C:\Users\Artin\Documents\voice files project\alf_hamze.mp3"},
        {'ؤ', @"C:\Users\Artin\Documents\voice files project\ve_hamze.mp3"},
        {'ئ', @"C:\Users\Artin\Documents\voice files project\ye_hamze.mp3"},


        {'a', @"C:\Users\Artin\Documents\voice files project\a.mp3"},
        {'b', @"C:\Users\Artin\Documents\voice files project\b.mp3"},
        {'c', @"C:\Users\Artin\Documents\voice files project\c.mp3"},
        {'d', @"C:\Users\Artin\Documents\voice files project\d.mp3"},
        {'e', @"C:\Users\Artin\Documents\voice files project\e.mp3"},
        {'f', @"C:\Users\Artin\Documents\voice files project\f.mp3"},
        {'g', @"C:\Users\Artin\Documents\voice files project\g.mp3"},
        {'h', @"C:\Users\Artin\Documents\voice files project\h.mp3"},
        {'i', @"C:\Users\Artin\Documents\voice files project\i.mp3"},
        {'j', @"C:\Users\Artin\Documents\voice files project\j.mp3"},
        {'k', @"C:\Users\Artin\Documents\voice files project\k.mp3"},
        {'l', @"C:\Users\Artin\Documents\voice files project\l.mp3"},
        {'m', @"C:\Users\Artin\Documents\voice files project\m.mp3"},
        {'n', @"C:\Users\Artin\Documents\voice files project\n.mp3"},
        {'o', @"C:\Users\Artin\Documents\voice files project\o.mp3"},
        {'p', @"C:\Users\Artin\Documents\voice files project\p.mp3"},
        {'q', @"C:\Users\Artin\Documents\voice files project\q.mp3"},
        {'r', @"C:\Users\Artin\Documents\voice files project\r.mp3"},
        {'s', @"C:\Users\Artin\Documents\voice files project\s.mp3"},
        {'t', @"C:\Users\Artin\Documents\voice files project\t.mp3"},
        {'u', @"C:\Users\Artin\Documents\voice files project\u.mp3"},
        {'v', @"C:\Users\Artin\Documents\voice files project\v.mp3"},
        {'w', @"C:\Users\Artin\Documents\voice files project\w.mp3"},
        {'x', @"C:\Users\Artin\Documents\voice files project\x.mp3"},
        {'y', @"C:\Users\Artin\Documents\voice files project\y.mp3"},
        {'z', @"C:\Users\Artin\Documents\voice files project\z.mp3"},

        {'1', @"C:\Users\Artin\Documents\voice files project\1.mp3"},
        {'2', @"C:\Users\Artin\Documents\voice files project\2.mp3"},
        {'3', @"C:\Users\Artin\Documents\voice files project\3.mp3"},
        {'4', @"C:\Users\Artin\Documents\voice files project\4.mp3"},
        {'5', @"C:\Users\Artin\Documents\voice files project\5.mp3"},
        {'6', @"C:\Users\Artin\Documents\voice files project\6.mp3"},
        {'7', @"C:\Users\Artin\Documents\voice files project\7.mp3"},
        {'8', @"C:\Users\Artin\Documents\voice files project\8.mp3"},
        {'9', @"C:\Users\Artin\Documents\voice files project\9.mp3"},
        // Add more mappings here
        };
        }







        private void Form1_Load(object sender, EventArgs e)
        {
            
            PositionGroupBoxes();

            DocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            mainFolderPath = Path.Combine(DocumentsPath, mainFolder);

            fontFilePath = Path.Combine(DocumentsPath, "subDatas", "dataFiles", "fontList.txt");
            filePathData = Path.Combine(DocumentsPath, "subDatas", "dataFiles", "data.txt");
            soundMappingFiles = Path.Combine(DocumentsPath, "subDatas", "voices");
            wordPath = Path.Combine(DocumentsPath, "wordFiles");
            // Create a new instance of the Word application

            try
            {

                ReadFileIntoList(filePathData, fontFilePath);
                projecktNummber = File.ReadAllLines(filePathData).Length;
                Console.WriteLine("****************" + projecktNummber);
                MessageBox.Show("data loaded seccessfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"an error occurred :{ex.Message}");
            }


            if (!Directory.Exists(mainFolder))
            {
                Directory.CreateDirectory(mainFolderPath);

                foreach (var folder in subfolder)
                {
                    Directory.CreateDirectory(Path.Combine(mainFolderPath, folder));
                }
                Console.WriteLine("the first run ! folders maided seccessfully");
            }
            else
            {
                Console.WriteLine("already exist!");
            }

            List<int> fontSize = new List<int> { 8, 9, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
            //size of fonts 

            foreach (int size in fontSize)
            {
                fontSizeComboBox.Items.Add(size.ToString());
            }



        }


        private void ReadFileIntoList(string filePath,string fontFilePath)
        {
            allDocxFiles.Clear();
            dataList.Clear();
            alphaList.Clear();
            nameList.Clear();
            voiceTagList.Clear();
            dateList.Clear();
            fontList.Clear();

            char delimiter = '|';

            try
            {

                if (Directory.Exists(wordPath))
                {
                    string[] docxFiles = Directory.GetFiles(wordPath, "*.docx", SearchOption.AllDirectories);
                    allDocxFiles = new List<string>(docxFiles); // Initialize and add found files
                }
                else
                {
                    Console.WriteLine($"Directory {wordPath} does not exist.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }



            

           

            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);
                dataList.AddRange(lines);
            }
            else
            {
                MessageBox.Show("file does not exist");
                return;
            }

            if (dataList.Count > 0)
            {
                foreach (string line in dataList)
                {
                    Console.WriteLine(line);
                    string[] substring = line.Split(delimiter);
                    if (substring.Length >= 4)
                    {
                        alphaList.Add(substring[0]);
                        nameList.Add(substring[1]);
                        voiceTagList.Add(substring[2]);
                        dateList.Add(substring[3]);
                    }
                }
            }
            else
            {
                MessageBox.Show("no data available ");
                return;
            }

            // Parse dateList to DateTime and sort indices based on the dates
            List<int> sortedIndices = dateList
                .Select((date, index) => new { Date = DateTime.Parse(date), Index = index })
                .OrderBy(x => x.Date)
                .Select(x => x.Index)
                .ToList();

            // Reorder all lists based on the sorted indices
            alphaList = sortedIndices.Select(index => alphaList[index]).ToList();
            nameList = sortedIndices.Select(index => nameList[index]).ToList();
            voiceTagList = sortedIndices.Select(index => voiceTagList[index]).ToList();
            dateList = sortedIndices.Select(index => dateList[index]).ToList();


            // Print sorted lists for verification
            for (int i = 0; i < alphaList.Count; i++)
            {
                Console.WriteLine($"{alphaList[i]} | {nameList[i]} | {voiceTagList[i]} | {dateList[i]}");
            }
        }




        public class Printer
        {
            private PrintDocument printDocument = new PrintDocument();
            private string documentText;

            public Printer(string text)
            {

                documentText = text;
                printDocument.PrintPage += new PrintPageEventHandler(PrintPage);

            }

            public void Print()
            {
                try
                {
                    printDocument.Print();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while printing: " + ex.Message);
                }
            }

            private void PrintPage(object sender, PrintPageEventArgs ev)
            {
                // Set the font and location for printing
                System.Drawing.Font printFont = new System.Drawing.Font("Arial", 12);
                float leftMargin = ev.MarginBounds.Left;
                float topMargin = ev.MarginBounds.Top;

                // Draw the string on the page
                ev.Graphics.DrawString(documentText, printFont, Brushes.Black, leftMargin, topMargin);
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            PositionGroupBoxes();
        }

        private void PositionGroupBoxes()
        {
            int margin = 10;

            // Top-left corner
            groupBox2.Location = new System.Drawing.Point(margin, margin);

            // Top-right corner
            groupBox1.Location = new System.Drawing.Point(ClientSize.Width - groupBox1.Width - margin, margin);

            // Bottom-left corner
            groupBox3.Location = new System.Drawing.Point(margin, ClientSize.Height - groupBox3.Height - margin);

            // Bottom-right corner
            groupBox4.Location = new System.Drawing.Point(ClientSize.Width - groupBox4.Width - margin, ClientSize.Height - groupBox4.Height - margin);
        }



        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (soundMappings.ContainsKey(e.KeyChar))
            {
                PlayMp3(soundMappings[e.KeyChar]);
            }
        }

        private void PlayMp3(string filePath)
        {
            // Dispose existing objects if they are not null
            if (waveOutDevice != null)
            {
                waveOutDevice.Dispose();
                waveOutDevice = null;
            }
            if (audioFileReader != null)
            {
                audioFileReader.Dispose();
                audioFileReader = null;
            }

            // Initialize the WaveOut and AudioFileReader
            waveOutDevice = new WaveOut();
            audioFileReader = new AudioFileReader(filePath);

            // Register event handler for playback stopped
            waveOutDevice.PlaybackStopped += OnPlaybackStopped;

            // Initialize and play  
            waveOutDevice.Init(audioFileReader);
            waveOutDevice.Play();
        }

        // Event handler for playback stopped
        private void OnPlaybackStopped(object sender, StoppedEventArgs e)
        {
            // Dispose WaveOut and AudioFileReader to release the file handle
            if (waveOutDevice != null)
            {
                waveOutDevice.Dispose();
                waveOutDevice = null;
            }
            if (audioFileReader != null)
            {
                audioFileReader.Dispose();
                audioFileReader = null;
            }
        }



        private TextBox textBox1;


        private void textBox1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            int currentPosition = textBox1.SelectionStart;

            if (e.KeyCode == Keys.Left && currentPosition > 0)
            {
                try
                {
                    Console.WriteLine("flag left");
                    char previousChar = textBox1.Text[currentPosition];
                    Console.WriteLine(previousChar);
                    if (soundMappings.ContainsKey(previousChar))
                    {
                        PlayMp3(soundMappings[previousChar]);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred: " + ex.Message);
                }
            }
            else if (e.KeyCode == Keys.Right && currentPosition < textBox1.Text.Length + 1)
            {
                try
                {
                    Console.WriteLine("flag right");
                    char nextChar = textBox1.Text[currentPosition - 1];
                    Console.WriteLine(nextChar);
                    if (soundMappings.ContainsKey(nextChar))
                    {
                        PlayMp3(soundMappings[nextChar]);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred: " + ex.Message);
                }
            }
        }


        private void startparBtn_Click(object sender, EventArgs e)
        {
            // Get the current cursor position
            int cursorPosition = textBox1.SelectionStart;

            // Find the start of the current paragraph (the first newline before the cursor)
            int currentParagraphStart = textBox1.Text.LastIndexOf(Environment.NewLine, cursorPosition - 1);

            if (currentParagraphStart > 0)
            {
                // Find the start of the previous paragraph (the newline before the current paragraph start)
                int previousParagraphStart = textBox1.Text.LastIndexOf(Environment.NewLine, currentParagraphStart - 1);

                // If a previous paragraph was found, move the cursor to the character after the previous paragraph's newline
                // If no previous newline, move to the start of the text
                textBox1.SelectionStart = previousParagraphStart >= 0 ? previousParagraphStart + Environment.NewLine.Length : 0;
            }
            else
            {
                // If the cursor is in the first paragraph, move to the start of the text
                textBox1.SelectionStart = 0;
            }

            // Scroll to the cursor position and focus the textbox
            textBox1.ScrollToCaret();
            textBox1.Focus();
        }





        private void endparBtn_Click(object sender, EventArgs e)
        {

            // Get the current cursor position
            int cursorPosition = textBox1.SelectionStart;

            // Find the start of the current paragraph (the first newline before the cursor)
            int currentParagraphStart = textBox1.Text.LastIndexOf(Environment.NewLine, cursorPosition - 1);

            if (currentParagraphStart > 0)
            {
                // Find the start of the previous paragraph (another newline before the current paragraph)
                int previousParagraphStart = textBox1.Text.LastIndexOf(Environment.NewLine, currentParagraphStart - 1);

                // Move the cursor to the start of the previous paragraph
                textBox1.SelectionStart = previousParagraphStart >= 0 ? previousParagraphStart + Environment.NewLine.Length : 0;

                // Scroll to the cursor position
                textBox1.ScrollToCaret();
                textBox1.Focus();
            }
            else
            {
                // If no previous paragraph, move to the very start of the text
                textBox1.SelectionStart = 0;
                textBox1.ScrollToCaret();
                textBox1.Focus();
            }

        }

        private void startsentBtn_Click(object sender, EventArgs e)
        {
            textBox1.Focus();
            // Get the current cursor position
            int cursorPosition = textBox1.SelectionStart;

            // Find the position of the previous period before the cursor
            int startOfSentence = textBox1.Text.LastIndexOf('.', cursorPosition - 1);

            // If the cursor is at the beginning of a sentence or there is no period found
            if (cursorPosition == startOfSentence + 1 || startOfSentence == -1)
            {
                // Find the period before the current sentence
                startOfSentence = textBox1.Text.LastIndexOf('.', startOfSentence - 1);

                // If another period is found, move to the character after it
                if (startOfSentence != -1)
                {
                    startOfSentence += 1; // Move to the first character after the period (and space)
                }
                else
                {
                    // If no previous period is found, move to the start of the text
                    startOfSentence = 0;
                }
            }
            else
            {
                // Move to the first character after the found period (and space)
                startOfSentence += 1;
            }

            // Set the cursor to the beginning of the sentence
            textBox1.SelectionStart = startOfSentence;
            textBox1.SelectionLength = 0;

            // Ensure the TextBox has focus so the cursor is visible
            textBox1.Focus();
        }



        public int[] savedCarentPosiotion = new int[10];
        public int countBmark = 1;





        private void Recording()
        {
            try
            {
                waveIn = new WaveInEvent();
                waveIn.WaveFormat = new WaveFormat(44100, 1);
                waveIn.DataAvailable += OnDataAvailable;
                waveIn.RecordingStopped += onRecordingStopped;


                writer = new WaveFileWriter(outputFilePath, waveIn.WaveFormat);

                waveIn.StartRecording();

                label1.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while starting the recording: " + ex.Message);
            }
        }

        private void stopRecording()
        {
            try
            {
                if (waveIn != null)
                {
                    waveIn.StopRecording();
                    label1.Visible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while stopping the recording: " + ex.Message);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {

            string bookChar = "ˑ";
            int textPlace = textBox1.SelectionStart;
            textBox1.Text = textBox1.Text.Substring(0, textPlace) + bookChar + textBox1.Text.Substring(textPlace);
            textBox1.SelectionStart = textPlace;
            textBox1.Focus();

            MessageBox.Show("push a key");

            this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
            this.textBox1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyUp);
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode >= Keys.A && e.KeyCode <= Keys.Z)
            {

                HandleAlphabetKeyDown(e.KeyCode);

            }
            else
            {
                MessageBox.Show("Wrong key, try again!");
            }

        }

        private bool isRecording = false;

        private void HandleAlphabetKeyDown(Keys key)
        {
            if (!isRecording)
            {
                isRecording = true;
                textBox1.ReadOnly = true;



                Recording();

                int textPlace = textBox1.SelectionStart;
                textBox1.Text = textBox1.Text.Substring(0, textPlace) + key + textBox1.Text.Substring(textPlace);
                textBox1.SelectionStart = textPlace + 1;
                textBox1.Focus();
            }
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (isRecording)
            {
                stopRecording();
                isRecording = false;
                textBox1.ReadOnly = false;
            }
        }


        private void exitBtn_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }





        private void Form1_keyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar))
            {
                string pressedCharOpen = e.KeyChar.ToString().ToUpper();
                bool keyFound = false;

                for (int i = 0; i < alphaList.Count; i++) // Loop through all items in alphaList
                {
                    if (pressedCharOpen == alphaList[i].ToUpper())
                    {
                        try
                        {
                            keyFound = true; // Mark that the key was found
                            string addressToOpen = allDocxFiles[i];
                            PlayMp3(voiceTagList[i]);
                            textBox1.Text = openWordDocument(addressToOpen);
                            textBox2.Text = nameList[i].Replace(".docx", "");
                            keyPressTcs.TrySetResult(true);

                            break; // Exit the loop since we found the key
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"an error occured{ex.Message}");
                        }

                    }
                }

                if (!keyFound)
                {
                    MessageBox.Show("Wrong key, try again!");
                }

                Console.WriteLine("keydown" + pressedCharOpen);
            }

            this.KeyPress -= new KeyPressEventHandler(Form1_keyPress);

        }



        // Dummy implementations of the methods to make the code compile



        private void OnDataAvailable(object sender, WaveInEventArgs e)
        {
            if (writer != null)
            {
                writer.Write(e.Buffer, 0, e.BytesRecorded);
                writer.Flush();
            }
        }

        private void onRecordingStopped(object sender, StoppedEventArgs e)
        {
            if (writer != null)
            {
                writer.Dispose();
                writer = null;
            }
            if (waveIn != null)
            {
                waveIn.Dispose();
                waveIn = null;
            }
            if (e.Exception != null)
            {
                MessageBox.Show("An error occurred during recording: " + e.Exception.Message);
            }
            else
            {
                MessageBox.Show("Recording saved in " + outputFilePath);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            if (isAlarmActive)
            {
                // Stop the alarm
                beepTimer.Stop();
                isAlarmActive = false;
            }
            else
            {
                // Start the alarm
                beepTimer.Start();
                isAlarmActive = true;
            }


            this.gotoBtn.Visible = !this.gotoBtn.Visible;
            this.bookMarkBtn.Visible = !this.bookMarkBtn.Visible;
        }

        private void gotoBtn_Click(object sender, EventArgs e)
        {
            this.startparBtn.Visible = !this.startparBtn.Visible;
            this.fontGroupBtn.Visible = !this.fontGroupBtn.Visible;
            endsentBtn.Visible = !endsentBtn.Visible;
            this.endparBtn.Visible = !this.endparBtn.Visible;
            startsentBtn.Visible = !startsentBtn.Visible;
            this.gotoBmark.Visible = !this.gotoBmark.Visible;
            this.searchBtn.Visible = !this.searchBtn.Visible;
            this.insertBtn.Visible = !this.insertBtn.Visible;
        }

        private void gotoBmark_Click(object sender, EventArgs e)
        {
            string Bname;
            textBoxValue = textBox1.Text;
            for (int i = 0; i < textBoxValue.Length; i++)
            {
                if (i.ToString() == "ˑ")
                {
                    Bname = (i++).ToString();
                    listBox1.Items.Add(Bname);

                }
            }


        }


        private void CreateWordDocument(string filePath)
        {

            Word.Application wordApp = null;
            Word.Document wordDoc = null;
            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Add();

                wordDoc.Content.Text = textBox1.Text;
                wordDoc.SaveAs2(filePath);
                MessageBox.Show($"Word document created: {filePath}");
            }
            catch (Exception er)
            {
                MessageBox.Show($"An error occurred: {er.Message}");
            }
            finally
            {
                if (wordDoc != null)
                {
                    wordDoc.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
            }



        }


        private void Form1_keyDown(object sender, KeyEventArgs e)
        {
            Console.WriteLine("flag key down");

            if (e.KeyCode >= Keys.A && e.KeyCode <= Keys.Z)
            {
                pressedChar = e.KeyCode.ToString();
                Console.WriteLine("keydown" + pressedChar);
                handelRecording(e.KeyCode);

            }
            else
            {
                MessageBox.Show("Wrong key, try again!");
            }


        }

        private void handelRecording(Keys key)
        {
            if (!isRecording)
            {
                isRecording = true;


                Recording();

            }

        }

        private void Form1_keyUp(object sender, KeyEventArgs e)
        {

            if (isRecording)
            {
                stopRecording();
                isRecording = false;

            }
            keyPressTcs.TrySetResult(true);
        }

        private TaskCompletionSource<bool> keyPressTcs;


        private async System.Threading.Tasks.Task WaitForKeyPressAsync()
        {
            keyPressTcs = new TaskCompletionSource<bool>();
            await keyPressTcs.Task;
        }


        // Ensure pressedChar is initialized
        private void saveBtn_Click(object sender, EventArgs e)
        {


            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                DialogResult result = MessageBox.Show("The file is new. Do you want to save it?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    // Handle the "Yes" button click (e.g., save the file)
                    // Attach the event handler for the saveAsBtn.Click event here
                    this.saveAsBtn.Click += saveAsBtn_Click;
                }
                else
                {
                    MessageBox.Show("You didn't specify a file name.");
                }
            }

            string wordFileName = $"{textBox2.Text}.docx";
            string wordFilePath = Path.Combine(DocumentsPath, "wordFiles", wordFileName);
            
            string voiceFilePath = Path.Combine(DocumentsPath, "subDatas", "voices", wordFileName);

            bool wordFileExists = File.Exists(wordFilePath);
            bool voiceFileExists = File.Exists(voiceFilePath);

            if (wordFileExists && voiceFileExists)
            {
                SaveTextToWordFile(wordFilePath, textBox1.Text);
                MessageBox.Show("Data saved successfully.");
            }
            else
            {
                DialogResult result = MessageBox.Show("The Word document or voice file doesn't exist. Do you want to create a new file?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    // Assuming saveAsBtn_Click is intended to handle creating new files
                    saveAsBtn_Click(sender, e);
                }
            }
        }

        private void SaveTextToWordFile(string filePath, string text)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = null;

            try
            {
                wordDoc = wordApp.Documents.Add();
                wordDoc.Content.Text = text;
                wordDoc.SaveAs2(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving document: {ex.Message}");
            }
            finally
            {
                wordDoc?.Close();
                wordApp.Quit();
            }
        }

        private void PopulateListBoxOpen() //to be change haie
        {
            listBox1.Items.Clear();

            for (int i = 0; i < allDocxFiles.Count; i++)
            {

                listBox1.Items.Add(Path.GetFileName(allDocxFiles[i]));
            }
            this.listBox1.KeyDown -= new System.Windows.Forms.KeyEventHandler(this.listBox1_keyDown);
            this.listBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.listBox1_keyDown);
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Get the selected index
            int index = listBox1.SelectedIndex;

            // Check if an item is selected
            if (index != -1)
            {
                string selectedFileName = listBox1.SelectedItem.ToString();

                Console.WriteLine("Selected index: " + index);
                Console.WriteLine("Selected item: " + selectedFileName);

                // Check if the index is within the range of voiceTagList
                if (index >= 0 && index < voiceTagList.Count)
                {
                    // Replace ".docx" with an empty string in the file name
                    string fileNameWithoutExtension = selectedFileName.Replace(".docx", "");

                    string voiceTag = Path.Combine(DocumentsPath, "subDatas", "voices", $"{fileNameWithoutExtension}.wav");

                    Console.WriteLine("Playing voice tag: " + voiceTag);
                    try
                    {
                        // Check if the voice tag file exists
                        if (File.Exists(voiceTag))
                        { 
                                PlayMp3(voiceTag);
                        }
                        else
                        {
                            MessageBox.Show("Voice file doesn't exist", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("voice doesnt exist");
                    }

                }
                else
                {
                    Console.WriteLine("Index out of range for voiceTagList");
                }
            }
            else
            {
                Console.WriteLine("No item is selected");
            }
        }



        private void listBox1_keyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Delete)
            {
                int index = listBox1.SelectedIndex;

                if (index != -1) // Check if an item is selected
                {
                    string wordName = listBox1.SelectedItem.ToString();

                    // Ensure index is within valid range for allDocxFiles
                    if (index >= 0 && index < allDocxFiles.Count)
                    {
                        string fileNameWithoutExtension = wordName.Replace(".docx", "");

                        string voiceTag = Path.Combine(DocumentsPath, "subDatas", "voices", $"{fileNameWithoutExtension}.wav");
                        string wordFilePath = Path.Combine(wordPath, wordName);

                        // Try to delete the voice file
                        if (File.Exists(voiceTag))
                        {
                            try
                            {
                                
                                

                                // Delete the voice file
                                File.Delete(voiceTag);
                                MessageBox.Show($"Voice file deleted: {voiceTag}");
                            }
                            catch (IOException ex)
                            {
                                MessageBox.Show($"Error: The file is in use and cannot be deleted. Please close any program that might be using it.\nDetails: {ex.Message}");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error deleting voice file: {ex.Message}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Voice file does not exist.");
                        }

                        // Try to delete the Word file
                        if (File.Exists(wordFilePath))
                        {
                            try
                            {
                                // Attempt to close the file if it is open elsewhere
                                

                                // Delete the Word file
                                File.Delete(wordFilePath);
                                MessageBox.Show($"Word file deleted: {wordFilePath}");
                            }
                            catch (IOException ex)
                            {
                                MessageBox.Show($"Error: The file is in use and cannot be deleted. Please close any program that might be using it.\nDetails: {ex.Message}");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error deleting Word file: {ex.Message}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Word file does not exist.");
                        }

                        // Update the dataList by removing the item at the current index
                        try
                        {
                            dataList.RemoveAt(index);

                            // Save updated dataList to file
                            using (StreamWriter writer = new StreamWriter(filePathData, false)) // false to overwrite the file
                            {
                                foreach (var item in dataList)
                                {
                                    writer.WriteLine(item); // Write each item on a new line
                                }
                            }
                            Console.WriteLine("Data list updated and saved.");

                            // Remove the item from the ListBox
                            listBox1.Items.RemoveAt(index);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error updating data list: {ex.Message}");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Index out of range. Please select a valid item.");
                    }
                }
                else
                {
                    MessageBox.Show("No item selected. Please select an item to delete.");
                }
            }


            if (e.KeyCode == Keys.Enter) { //open the file 
                int index = listBox1.SelectedIndex;

                if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    DialogResult result = MessageBox.Show("Are you sure you want to discard the changes?", "Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        // Execute the logic for discarding changes
                        // (e.g., clear the TextBox, reset the form, etc.)
                    }
                    else
                    {
                        // Only add the event handler once, or ensure it's added only if it's not already attached
                        this.saveBtn.Click -= saveBtn_Click; // Remove if already attached
                        this.saveBtn.Click += saveBtn_Click;
                    }
                }


                if (index != ListBox.NoMatches)// Check if an item is clicked
                {
                    Console.WriteLine(index);
                    wordName = listBox1.SelectedItem.ToString();
                    // Check if the index is within the range of voiceTagList
                    if (index >= 0 && index < allDocxFiles.Count)
                    {
                        wordName.Replace(".docx", "");
                        string addressToOpen = Path.Combine(DocumentsPath, "wordFiles", wordName);
                        textBox1.Text = openWordDocument(addressToOpen);
                        textBox2.Text = wordName;
                    }
                    else
                    {
                        Console.WriteLine("Index out of range for voiceTagList");
                    }
                }
            }

            if (e.KeyCode == Keys.Space)
            {
                bool removableDriveFound = false;

                DriveInfo[] drives = DriveInfo.GetDrives();
                foreach (DriveInfo drive in drives)
                {
                    if (drive.DriveType == DriveType.Removable && drive.IsReady)
                    {
                        removableDriveFound = true;
                        int index = listBox1.SelectedIndex;

                        if (index >= 0 && index < nameList.Count)
                        {
                            Console.WriteLine($"Found removable drive: {drive.Name}");

                            string sourceFilePath = Path.Combine(wordPath, nameList[index]);
                            string destinationFilePath = Path.Combine(drive.Name, nameList[index]);

                            try
                            {
                                File.Copy(sourceFilePath, destinationFilePath, overwrite: true);
                                MessageBox.Show($"File copied to {destinationFilePath}");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error copying file: {ex.Message}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("No file selected or invalid selection.");
                        }
                    }
                }

                if (!removableDriveFound)
                {
                    MessageBox.Show("Cannot find any external storages");
                }
            }
            // Detect external storage drives

            if(e.KeyCode == Keys.Back)
            {
                MessageBox.Show("enter the new name of the file");
                textBox2.Enabled = true;
                textBox2.Focus();
                this.textBox2.KeyDown -= new System.Windows.Forms.KeyEventHandler(this.textBox_keyDown_Rename);
                this.textBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_keyDown_Rename);
                
            }

        }


        private void textBox_keyDown_Rename(object sender, KeyEventArgs e)
        {
            if (string.IsNullOrEmpty(textBox2.Text))
            {
                return;
            }
            if(e.KeyCode == Keys.Enter)
            {
                string wordname = listBox1.SelectedItem.ToString();
                
                string oldWordFile = Path.Combine(wordPath, wordname);
                string newWordFile = Path.Combine(wordPath, $"{textBox2.Text}.docx");

                wordname = wordname.Replace(".docx", "");
                string oldVoiceTag = Path.Combine(mainFolderPath, "voices", $"{wordname}0.wav");
                string newVoiceTag = Path.Combine(mainFolderPath, "voices", $"{textBox2.Text}0.wav");

                try
                {
                    File.Move(oldWordFile, newWordFile);
                    File.Move(oldVoiceTag, newVoiceTag);
                    MessageBox.Show("rename was seccesful");

                    this.textBox2.KeyDown -= new System.Windows.Forms.KeyEventHandler(this.textBox_keyDown_Rename);
                    textBox2.Enabled = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"error while renameing the files: {ex.Message}");
                }
                
            }
      

        }
        private string openWordDocument(string filePath)
        {
            Word.Application wordApp = null;
            Word.Document wordDoc = null;
            string textBoxContent = string.Empty;
            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Open(filePath);
                //wordApp.Visible = true;
                textBoxContent = wordDoc.Content.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"an error occurred:{ex.Message}");
            }
            finally
            {
                wordDoc.Close();
                wordApp.Quit();

            }
            return textBoxContent;
        }

        

        private void BeepTimer_Tick(object sender, EventArgs e)
        {
            SystemSounds.Beep.Play(); // Play the beep sound
        }

        private void fileMagament_Click(object sender, EventArgs e)
        {

            if (isAlarmActive)
            {
                // Stop the alarm
                beepTimer.Stop();
                isAlarmActive = false;
                
                listBox1.Items.Clear();
            }
            else
            {
                // Start the alarm
                beepTimer.Start();
                isAlarmActive = true;
                this.listBox1.SelectedIndexChanged += new EventHandler(this.ListBox1_SelectedIndexChanged);
                listBox1.Items.Clear();
                PopulateListBoxOpen();
                listBox1.Focus();
            }
            
            DialogResult result = MessageBox.Show("do you want to use random accsess or regular access", "info");

            ToggleButtonVisibilityFileManage();
            ManageFiles();
        }


       
        private void ToggleButtonVisibilityFileManage()
        {
            this.gotoBtn.Visible = !this.gotoBtn.Visible;
            this.saveAsBtn.Visible = !this.saveAsBtn.Visible;
            this.printBtn.Visible = !this.printBtn.Visible;
            this.savePdfBtn.Visible = !this.savePdfBtn.Visible;
            this.searchBtn.Visible = !this.searchBtn.Visible;
        }

        private void ManageFiles()
        {
            string wordFolder = Path.Combine(DocumentsPath, "wordFiles");
            this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Form1_keyPress);
            this.KeyPress -= new System.Windows.Forms.KeyPressEventHandler(this.Form1_keyPress);
        }



        private void starttextBtn_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_keyPress(object sender, KeyPressEventArgs e)
        {
            if (soundMappings.ContainsKey(e.KeyChar))
            {
                PlayMp3(soundMappings[e.KeyChar]);
            }
        }

        private void printBtn_Click(object sender, EventArgs e)
        {
            Printer printer = new Printer(textBox1.Text);
            printer.Print();
        }


    private void savePdfBtn_Click(object sender, EventArgs e)
    {
        try
        {
            // Ensure the output directory exists
            string outputDir = Path.Combine(DocumentsPath, "subDatas", "pdf folder");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // File path to save the PDF
            string outputFile = Path.Combine(outputDir, $"{textBox2.Text}.pdf");

            // Create a PDF document
            Document document = new Document();
            using (FileStream fs = new FileStream(outputFile, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                document.Open();

                // Use a font that supports Arabic
                BaseFont baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 12);

                // Add content to the document using the specified font
                Paragraph paragraph = new Paragraph(textBox1.Text, font)
                {
                    Alignment = Element.ALIGN_RIGHT // Align text to the right for Arabic
                };
                document.Add(paragraph);

                document.Close();
            }

            MessageBox.Show("PDF created successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"An error occurred while creating the PDF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }


        private void saveAsBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("clicked");
            listBox1.Focus();// Focus on the list box


            textBox2.Enabled = true;
            textBox2.Focus();

            this.textBox2.KeyDown -= new System.Windows.Forms.KeyEventHandler(this.textBox2_KeyDown);
            this.textBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox2_KeyDown);




        }
        private async void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && textBox2.Text != null)
            {
                DateTime currentDateTime = DateTime.Now;
                string dateTimeString = currentDateTime.ToString("yyyy-MM-dd HH:mm:ss");

                wordName = textBox2.Text;
                wordFilePath = Path.Combine(DocumentsPath, "wordFiles", $"{wordName}.docx");

                CreateWordDocument(wordFilePath);

                MessageBox.Show("Press a key", "Press key", MessageBoxButtons.OK, MessageBoxIcon.Information);


                // Detach previous handlers to avoid multiple subscriptions


                this.Enabled = false;
                // Attach new handlers
                this.KeyDown += new KeyEventHandler(Form1_keyDown);
                textBox2.Enabled = false;

                outputFilePath = Path.Combine(mainFolderPath, "voices", $"{wordName}.wav");
                this.KeyUp += new KeyEventHandler(Form1_keyUp);

                this.Enabled = true;
                await WaitForKeyPressAsync();
                //this.Enabled = true;





                // Check if any item in alphaList starts with pressedChar followed by a possible index
                if (alphaList.Any(item => item.StartsWith(pressedChar)))
                {
                    // Find the last index of items starting with pressedChar in alphaList
                    int lastIndex = alphaList.FindLastIndex(item => item.StartsWith(pressedChar));

                    // Extract the current index from the last matching item
                    string lastItem = alphaList[lastIndex];
                    int startIndex = pressedChar.Length;
                    int currentIndex = 0;

                    // Check if there is an index to extract
                    if (lastItem.Length > startIndex && lastItem[startIndex] == '(' && lastItem.EndsWith(")"))
                    {
                        string indexString = lastItem.Substring(startIndex + 1, lastItem.Length - startIndex - 2);
                        if (int.TryParse(indexString, out int parsedIndex))
                        {
                            currentIndex = parsedIndex;
                        }
                    }

                    // Append (currentIndex + 1) to pressedChar
                    pressedChar = $"{pressedChar}({currentIndex + 1})";
                }

                // Add pressedChar to alphaList if it's not already there
                if (!alphaList.Contains(pressedChar))
                {
                    alphaList.Add(pressedChar);
                }
                Console.WriteLine("Before saving char: " + pressedChar);



                // Write data to file with appended pressedChar
                try
                {
                    using (StreamWriter writer = new StreamWriter(filePathData, true))
                    {
                        writer.WriteLine($"{pressedChar}|{wordName}.docx|{outputFilePath}|{dateTimeString}");
                    }

                    // Re-enable the form


                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving data: {ex.Message}");
                    this.Enabled = true;
                }
                finally
                {
                    this.KeyDown -= new KeyEventHandler(Form1_keyDown); ;
                    this.KeyUp -= new KeyEventHandler(Form1_keyUp);
                    pressedChar = string.Empty;


                    textBox2.Text = wordName;
                    this.textBox2.KeyDown -= new System.Windows.Forms.KeyEventHandler(this.textBox2_KeyDown);

                }

            }
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        
        private void fontGroupBtn_Click(object sender, EventArgs e)
        {
            if (isAlarmActive)
            {
                // Stop the alarm
                beepTimer.Stop();
                isAlarmActive = false;
                removeItemsListBox();
                listBox1.Items.Clear();
            }
            else
            {
                // Start the alarm
                beepTimer.Start();
                isAlarmActive = true;
                // Check if the selected item has changed


                removeItemsListBox();
                populateFontListBox();
                listBox1.Focus();
               

                // Update the last selected item
               
            }



            insertBtn.Visible = !insertBtn.Visible;
            gotoBtn.Visible = !gotoBtn.Visible;
            searchBtn.Visible = !searchBtn.Visible;
            fileManagment.Visible = !fileManagment.Visible;
            fontComboBox.Visible = !fontComboBox.Visible;
            fontSizeComboBox.Visible = !fontSizeComboBox.Visible;
            changeColor.Visible = !changeColor.Visible;
            BIUbtn.Visible = !BIUbtn.Visible;
            PBObtn.Visible = !PBObtn.Visible;
            statusBtn.Visible = !statusBtn.Visible;
            
        }
        
        private void removeItemsListBox()
        {
            
            for(int i = 0; i <= listBox1.Items.Count - 1; i++)
            {
                listBox1.Items.Remove(i);
            }
            
        }

        private void populateFontListBox()
        {
            listBox1.Items.Clear();
            fontList.Clear();
            using (StreamReader reader = new StreamReader(fontFilePath, true))
            {
                while (!reader.EndOfStream)
                {
                    string item = reader.ReadLine();
                    fontList.Add(item);
                }
            }



            foreach (string font in fontList)
            {
                string[] fonts = font.Split('|');
                for (int i = 0; i < fonts.Length; i++)
                {
                    if (i % 3 == 1) // Check if the index is odd
                    {
                        listBox1.Items.Add(fonts[i]);
                        // Add the odd-numbered element to your ListBox (listBox1.Items.Add(...))
                    }
                }

            }

            this.listBox1.KeyDown += new KeyEventHandler(this.listBox_KeyDown_Font);

        }

        private void listBox_KeyDown_Font(object sender, KeyEventArgs e)
        {
            int index = listBox1.SelectedIndex;

            if (index != -1) // Check if an item is selected
            {
                if (e.KeyCode == Keys.Delete)
                
                    // Ensure index is within valid range for allDocxFiles
                    if (index >= 0 && index < allDocxFiles.Count)
                    {
                        string voiceName = listBox1.SelectedIndex.ToString();


                        string voiceTag = Path.Combine(DocumentsPath, "subDatas", "voices", $"{voiceName}.wav");
                        

                        // Try to delete the voice file
                        if (File.Exists(voiceTag))
                        {
                            try
                            {
                                // Attempt to close the file if it is open elsewhere


                                // Delete the voice file
                                File.Delete(voiceTag);
                                MessageBox.Show($"Voice file deleted: {voiceTag}");
                            }
                            catch (IOException ex)
                            {
                                MessageBox.Show($"Error: The file is in use and cannot be deleted. Please close any program that might be using it.\nDetails: {ex.Message}");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error deleting voice file: {ex.Message}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Voice file does not exist.");
                        }


                        try
                        {
                            fontList.RemoveAt(index);

                            // Save updated dataList to file
                            using (StreamWriter writer = new StreamWriter(fontFilePath, false)) // false to overwrite the file
                            {
                                foreach (var item in fontList)
                                {
                                    writer.WriteLine(item); // Write each item on a new line
                                }
                            }
                            Console.WriteLine("Data list updated and saved.");

                            // Remove the item from the ListBox
                            listBox1.Items.RemoveAt(index);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error updating data list: {ex.Message}");
                        }

                    }   
                    else
                    {
                        MessageBox.Show("No item selected. Please select an item to delete.");
                    }
            }
        }

        private void changeColor_Click(object sender, EventArgs e)
        {
            MessageBox.Show("enter chararcter between ORGBCMYW");
            this.KeyPress += new KeyPressEventHandler(this.Form1_keyPress_color);
        }

        private void Form1_keyPress_color(object sender , KeyPressEventArgs e)
        {
            
            string getColor = e.KeyChar.ToString().ToUpper();
            switch (getColor)
            {
                case "O":
                    textBox1.ForeColor = System.Drawing.Color.Orange;
                    MessageBox.Show("color changed");
                    break;
                case "R":

                    textBox1.ForeColor = System.Drawing.Color.Red;
                    break;
                case "G":

                    textBox1.ForeColor = System.Drawing.Color.Green;
                    break;
                case "B":

                    textBox1.ForeColor = System.Drawing.Color.Blue;
                    break;
                case "W":

                    textBox1.ForeColor = System.Drawing.Color.White;
                    break;
                case "M":

                    textBox1.ForeColor = System.Drawing.Color.Maroon;
                    break;
                case "Y":

                    textBox1.ForeColor = System.Drawing.Color.Yellow;
                    break;
                case "C":

                    textBox1.ForeColor = System.Drawing.Color.Black;
                    break;
                default:
                    MessageBox.Show("invalid color key");
                    break;
            }
            //keyPressTcs.SetResult(true);
            this.KeyPress -= new KeyPressEventHandler(this.Form1_keyPress_color);

        }

        private void BIUbtn_Click(object sender, EventArgs e)
        {
            this.KeyPress += new KeyPressEventHandler(this.Form1_keyPress_style);
        }

        private void Form1_keyPress_style(object sender , KeyPressEventArgs e)
        {
            string getStyle = e.KeyChar.ToString().ToUpper();
            switch (getStyle)
            {
                case "B":
                    ToggleFontStyle(FontStyle.Bold);
                    break;
                case "I":
                    ToggleFontStyle(FontStyle.Italic);
                    break;
                case "U":
                    ToggleFontStyle(FontStyle.Underline);
                    break;
                default:
                    MessageBox.Show("wrong style key");
                    break;
            }
            this.KeyPress -= new KeyPressEventHandler(this.Form1_keyPress_style);

        }
        private void ToggleFontStyle(FontStyle style)
        {
            System.Drawing.Font currentFont = this.textBox1.Font;
            FontStyle newFontStyle;

            // Check if the style is already applied
            if (currentFont.Style.HasFlag(style))
            {
                // Remove the style if it's already applied
                newFontStyle = currentFont.Style & ~style;
            }
            else
            {
                // Add the style if it's not applied
                newFontStyle = currentFont.Style | style;
            }

            this.textBox1.Font = new System.Drawing.Font(currentFont, newFontStyle);
        }



        private void PopulateFontComboBox()
        {
            InstalledFontCollection installedFonts = new InstalledFontCollection();
            var fontFamilies = installedFonts.Families.OrderBy(f => f.Name);

            foreach (var font in fontFamilies)
            {
                fontComboBox.Items.Add(font.Name);
            }
        }

        private void FontComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();

            if (e.Index >= 0)
            {
                string fontName = fontComboBox.Items[e.Index].ToString();
                using (System.Drawing.Font font = new System.Drawing.Font(fontName, e.Font.Size, FontStyle.Regular, GraphicsUnit.Point))
                {
                    e.Graphics.DrawString(fontName, font, Brushes.Black, e.Bounds);
                }
            }

            e.DrawFocusRectangle();
        }

        private void FontComboBox_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            if (e.Index >= 0)
            {
                string fontName = fontComboBox.Items[e.Index].ToString();
                using (System.Drawing.Font font = new System.Drawing.Font(fontName, 10, FontStyle.Regular, GraphicsUnit.Point))
                {
                    SizeF size = e.Graphics.MeasureString(fontName, font);
                    e.ItemHeight = (int)size.Height;
                    e.ItemWidth = (int)size.Width;
                }
            }
        }

        private void FontComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (fontList.Count >= 10)
            {
                MessageBox.Show("you cant add more than 10 fonts in your list");
                return;
            }

            if (fontComboBox.SelectedItem != null)
            {
                getComboItem(fontComboBox.SelectedItem.ToString());
            }
        }

        private async void getComboItem(string selectedFont)
        {
            // Add the selected font to the fontList
            if (!fontList.Contains(selectedFont))
            {
                fontList.Add(selectedFont);
            }

            // Set the output file path
            

            // Ensure KeyDown and KeyUp event handlers are added only once
            this.KeyDown -= Form1_keyDown;
            this.KeyDown += Form1_keyDown;
            outputFilePath = Path.Combine(mainFolderPath, "voices", $"{selectedFont}.wav");
            
            this.KeyUp -= Form1_keyUp;
            this.KeyUp += Form1_keyUp;


            await WaitForKeyPressAsync();
            this.KeyDown -= Form1_keyDown;
            this.KeyUp -= Form1_keyUp;

            saveFontFile(selectedFont);
           
        }
        private void saveFontFile(string selectedFont)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(fontFilePath, true))
                {
                    writer.WriteLine($"{pressedChar}|{selectedFont}|{outputFilePath}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while saving the font file: {ex.Message}");
            }
            finally
            {
                this.KeyDown -= Form1_keyDown;
                this.KeyUp -= Form1_keyUp;
            }
        }

       
        private void ComboBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Suppress keypresses for alphanumeric keys
            if (char.IsLetterOrDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

   
        private void fontSizeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateFontSize();

            if (fontSizeComboBox.SelectedItem != null)
            {
                string selectedSize = fontSizeComboBox.SelectedItem.ToString();
                MessageBox.Show($"Selected font: {selectedSize}");
            }
        }

        private void updateFontSize()
        {
            if(int.TryParse(fontSizeComboBox.Text, out int newSize))
            {
                textBox1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, newSize);
            }
            else
            {
                MessageBox.Show("enter valid number");
            }
        }

        private void fontSizeComboBox_keyPress(object sender , KeyPressEventArgs e)
        {
        
            if(!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }

            if(e.KeyChar == (char)Keys.Enter)
            {
                updateFontSize();
            }
        }

        private void PBObtn_Click(object sender, EventArgs e)
        {
            Chunk p = new Chunk("normal text");
            Chunk superscript = new Chunk("superscript");
            superscript.SetTextRise(5f);
            superscript.Font.Size = 12;
            textBox1.AppendText(String.Concat(p, superscript));
            
        }

        private void statusBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show($"{textBox1.Font}{textBox1.Font.Style}{textBox1.ForeColor}");
        }

        private void endsentBtn_Click(object sender, EventArgs e)
        {
                // Get the current cursor position
                int cursorPosition = textBox1.SelectionStart;

                // Find the position of the next period after the cursor
                int endOfSentence = textBox1.Text.IndexOf('.', cursorPosition);

                // If a period is found
                if (endOfSentence != -1)
                {
                    // Move to the character after the period (accounting for the period and space)
                    textBox1.SelectionStart = endOfSentence + 1;

                    // If there is a space after the period, move the cursor after the space
                    if (endOfSentence + 1 < textBox1.Text.Length && textBox1.Text[endOfSentence + 1] == ' ')
                    {
                        textBox1.SelectionStart++;
                    }
                }
                else
                {
                    // If no period is found, move to the end of the text
                    textBox1.SelectionStart = textBox1.Text.Length;
                }

                // Ensure the TextBox has focus so the cursor is visible
                textBox1.Focus();
            
        }

        private void settingBtn_Click(object sender, EventArgs e)
        {
            setting_form sf = new setting_form();
            sf.Show();

        }
    }
}


    

    



