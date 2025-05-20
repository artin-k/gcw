using Microsoft.Office.Interop.Word;
using System;

namespace WindowsFormsApp1
{
	partial class mainForm
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}



        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        /// 
        private void InitializeComponent()
        {
            this.saveBtn = new System.Windows.Forms.Button();
            this.searchBtn = new System.Windows.Forms.Button();
            this.startparBtn = new System.Windows.Forms.Button();
            this.parGroup = new System.Windows.Forms.Button();
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
            this.btnInsertPageBreak = new System.Windows.Forms.Button();
            this.fileManagment = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.changeColor = new System.Windows.Forms.Button();
            this.btnSetSpacing = new System.Windows.Forms.Button();
            this.aligenmentBtn = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btnToggleLanguage = new System.Windows.Forms.Button();
            this.BIUbtn = new System.Windows.Forms.Button();
            this.fontGroupBtn = new System.Windows.Forms.Button();
            this.rtlBtn = new System.Windows.Forms.Button();
            this.filesListBox = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.titleTextBox = new System.Windows.Forms.TextBox();
            this.settingBtn = new System.Windows.Forms.Button();
            this.MainrichTextBox = new System.Windows.Forms.RichTextBox();
            this.spacingValue = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.exportBtn = new System.Windows.Forms.Button();
            this.bMarkList = new System.Windows.Forms.ListBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // saveBtn
            // 
            this.saveBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.saveBtn.ImageAlign = System.Drawing.ContentAlignment.TopRight;
            this.saveBtn.Location = new System.Drawing.Point(258, 0);
            this.saveBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
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
            this.searchBtn.Location = new System.Drawing.Point(251, 68);
            this.searchBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.searchBtn.Name = "searchBtn";
            this.searchBtn.Size = new System.Drawing.Size(101, 57);
            this.searchBtn.TabIndex = 2;
            this.searchBtn.Text = "جستجو";
            this.searchBtn.UseVisualStyleBackColor = true;
            // 
            // startparBtn
            // 
            this.startparBtn.Location = new System.Drawing.Point(258, 25);
            this.startparBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.startparBtn.Name = "startparBtn";
            this.startparBtn.Size = new System.Drawing.Size(73, 89);
            this.startparBtn.TabIndex = 3;
            this.startparBtn.Text = "شروع پاراگراف";
            this.startparBtn.UseVisualStyleBackColor = true;
            this.startparBtn.Visible = false;
            this.startparBtn.Click += new System.EventHandler(this.startparBtn_Click);
            // 
            // parGroup
            // 
            this.parGroup.Location = new System.Drawing.Point(199, 25);
            this.parGroup.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.parGroup.Name = "parGroup";
            this.parGroup.Size = new System.Drawing.Size(140, 58);
            this.parGroup.TabIndex = 4;
            this.parGroup.Text = "paragraph ";
            this.parGroup.UseVisualStyleBackColor = true;
            this.parGroup.Click += new System.EventHandler(this.parGroup_Click);
            // 
            // endparBtn
            // 
            this.endparBtn.Location = new System.Drawing.Point(59, 87);
            this.endparBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.endparBtn.Name = "endparBtn";
            this.endparBtn.Size = new System.Drawing.Size(140, 85);
            this.endparBtn.TabIndex = 5;
            this.endparBtn.Text = "پایان پاراگراف";
            this.endparBtn.UseVisualStyleBackColor = true;
            this.endparBtn.Visible = false;
            this.endparBtn.Click += new System.EventHandler(this.endparBtn_Click);
            // 
            // startsentBtn
            // 
            this.startsentBtn.Location = new System.Drawing.Point(6, 114);
            this.startsentBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.startsentBtn.Name = "startsentBtn";
            this.startsentBtn.Size = new System.Drawing.Size(147, 58);
            this.startsentBtn.TabIndex = 6;
            this.startsentBtn.Text = "شروع جمله";
            this.startsentBtn.UseVisualStyleBackColor = true;
            this.startsentBtn.Visible = false;
            this.startsentBtn.Click += new System.EventHandler(this.startsentBtn_Click);
            // 
            // bookMarkBtn
            // 
            this.bookMarkBtn.Location = new System.Drawing.Point(95, 74);
            this.bookMarkBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.bookMarkBtn.Name = "bookMarkBtn";
            this.bookMarkBtn.Size = new System.Drawing.Size(151, 55);
            this.bookMarkBtn.TabIndex = 7;
            this.bookMarkBtn.Text = "بوک مارک";
            this.bookMarkBtn.UseVisualStyleBackColor = true;
            this.bookMarkBtn.Visible = false;
            this.bookMarkBtn.Click += new System.EventHandler(this.bMarkBtn_Click);
            // 
            // endsentBtn
            // 
            this.endsentBtn.Location = new System.Drawing.Point(5, 20);
            this.endsentBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.endsentBtn.Name = "endsentBtn";
            this.endsentBtn.Size = new System.Drawing.Size(131, 66);
            this.endsentBtn.TabIndex = 8;
            this.endsentBtn.Text = "پایان جمله";
            this.endsentBtn.UseVisualStyleBackColor = true;
            this.endsentBtn.Visible = false;
            this.endsentBtn.Click += new System.EventHandler(this.endsentBtn_Click);
            // 
            // starttextBtn
            // 
            this.starttextBtn.Location = new System.Drawing.Point(0, 103);
            this.starttextBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.starttextBtn.Name = "starttextBtn";
            this.starttextBtn.Size = new System.Drawing.Size(111, 66);
            this.starttextBtn.TabIndex = 9;
            this.starttextBtn.Text = "شروع متن";
            this.starttextBtn.UseVisualStyleBackColor = true;
            this.starttextBtn.Visible = false;
            // 
            // exitBtn
            // 
            this.exitBtn.Location = new System.Drawing.Point(0, 15);
            this.exitBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
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
            this.endtextBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.endtextBtn.Name = "endtextBtn";
            this.endtextBtn.Size = new System.Drawing.Size(99, 58);
            this.endtextBtn.TabIndex = 11;
            this.endtextBtn.Text = "پایان متن";
            this.endtextBtn.UseVisualStyleBackColor = true;
            this.endtextBtn.Visible = false;
            // 
            // gotoBtn
            // 
            this.gotoBtn.Location = new System.Drawing.Point(1401, 21);
            this.gotoBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.gotoBtn.Name = "gotoBtn";
            this.gotoBtn.Size = new System.Drawing.Size(108, 50);
            this.gotoBtn.TabIndex = 12;
            this.gotoBtn.Text = "برو به ";
            this.gotoBtn.UseVisualStyleBackColor = true;
            this.gotoBtn.Click += new System.EventHandler(this.gotoBtn_Click);
            // 
            // gotoBmark
            // 
            this.gotoBmark.Location = new System.Drawing.Point(95, 106);
            this.gotoBmark.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.gotoBmark.Name = "gotoBmark";
            this.gotoBmark.Size = new System.Drawing.Size(244, 57);
            this.gotoBmark.TabIndex = 13;
            this.gotoBmark.Text = "go bookmark";
            this.gotoBmark.UseVisualStyleBackColor = true;
            this.gotoBmark.Visible = false;
            this.gotoBmark.Click += new System.EventHandler(this.gotoBmark_Click);
            // 
            // insertBtn
            // 
            this.insertBtn.Location = new System.Drawing.Point(152, 20);
            this.insertBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.insertBtn.Name = "insertBtn";
            this.insertBtn.Size = new System.Drawing.Size(95, 48);
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
            this.groupBox1.Controls.Add(this.gotoBmark);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(1567, 9);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox1.Size = new System.Drawing.Size(358, 281);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            // 
            // fontSizeComboBox
            // 
            this.fontSizeComboBox.FormattingEnabled = true;
            this.fontSizeComboBox.Location = new System.Drawing.Point(243, 197);
            this.fontSizeComboBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.fontSizeComboBox.Name = "fontSizeComboBox";
            this.fontSizeComboBox.Size = new System.Drawing.Size(107, 37);
            this.fontSizeComboBox.TabIndex = 26;
            this.fontSizeComboBox.Visible = false;
            this.fontSizeComboBox.SelectedIndexChanged += new System.EventHandler(this.fontSizeComboBox_SelectedIndexChanged);
            // 
            // fontComboBox
            // 
            this.fontComboBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.fontComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.fontComboBox.Location = new System.Drawing.Point(1, 111);
            this.fontComboBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.fontComboBox.Name = "fontComboBox";
            this.fontComboBox.Size = new System.Drawing.Size(112, 35);
            this.fontComboBox.TabIndex = 0;
            this.fontComboBox.Visible = false;
            this.fontComboBox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.FontComboBox_DrawItem);
            this.fontComboBox.SelectedIndexChanged += new System.EventHandler(this.FontComboBox_SelectedIndexChanged);
            this.fontComboBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ComboBox_KeyPress);
            // 
            // savePdfBtn
            // 
            this.savePdfBtn.Location = new System.Drawing.Point(120, 133);
            this.savePdfBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.savePdfBtn.Name = "savePdfBtn";
            this.savePdfBtn.Size = new System.Drawing.Size(116, 48);
            this.savePdfBtn.TabIndex = 7;
            this.savePdfBtn.Text = "save pdf";
            this.savePdfBtn.UseVisualStyleBackColor = true;
            this.savePdfBtn.Visible = false;
            this.savePdfBtn.Click += new System.EventHandler(this.savePdfBtn_Click);
            // 
            // printBtn
            // 
            this.printBtn.Location = new System.Drawing.Point(19, 47);
            this.printBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.printBtn.Name = "printBtn";
            this.printBtn.Size = new System.Drawing.Size(111, 46);
            this.printBtn.TabIndex = 6;
            this.printBtn.Text = "print";
            this.printBtn.UseVisualStyleBackColor = true;
            this.printBtn.Visible = false;
            this.printBtn.Click += new System.EventHandler(this.printBtn_Click);
            // 
            // saveAsBtn
            // 
            this.saveAsBtn.Location = new System.Drawing.Point(116, 24);
            this.saveAsBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.saveAsBtn.Name = "saveAsBtn";
            this.saveAsBtn.Size = new System.Drawing.Size(109, 46);
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
            this.groupBox2.Controls.Add(this.btnInsertPageBreak);
            this.groupBox2.Controls.Add(this.saveAsBtn);
            this.groupBox2.Controls.Add(this.fileManagment);
            this.groupBox2.Controls.Add(this.exitBtn);
            this.groupBox2.Controls.Add(this.starttextBtn);
            this.groupBox2.Controls.Add(this.endtextBtn);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(12, 12);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox2.Size = new System.Drawing.Size(269, 283);
            this.groupBox2.TabIndex = 19;
            this.groupBox2.TabStop = false;
            // 
            // statusBtn
            // 
            this.statusBtn.Location = new System.Drawing.Point(5, 136);
            this.statusBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.statusBtn.Name = "statusBtn";
            this.statusBtn.Size = new System.Drawing.Size(85, 55);
            this.statusBtn.TabIndex = 27;
            this.statusBtn.Text = "status";
            this.statusBtn.UseVisualStyleBackColor = true;
            this.statusBtn.Visible = false;
            this.statusBtn.Click += new System.EventHandler(this.statusBtn_Click);
            // 
            // btnInsertPageBreak
            // 
            this.btnInsertPageBreak.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInsertPageBreak.Location = new System.Drawing.Point(76, 54);
            this.btnInsertPageBreak.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnInsertPageBreak.Name = "btnInsertPageBreak";
            this.btnInsertPageBreak.Size = new System.Drawing.Size(171, 68);
            this.btnInsertPageBreak.TabIndex = 35;
            this.btnInsertPageBreak.Text = "break page";
            this.btnInsertPageBreak.UseVisualStyleBackColor = true;
            this.btnInsertPageBreak.Visible = false;
            this.btnInsertPageBreak.Click += new System.EventHandler(this.btnInsertPageBreak_Click);
            // 
            // fileManagment
            // 
            this.fileManagment.Location = new System.Drawing.Point(76, 205);
            this.fileManagment.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.fileManagment.Name = "fileManagment";
            this.fileManagment.Size = new System.Drawing.Size(187, 47);
            this.fileManagment.TabIndex = 7;
            this.fileManagment.Text = "file magment";
            this.fileManagment.UseVisualStyleBackColor = true;
            this.fileManagment.Click += new System.EventHandler(this.fileMagament_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox3.AutoSize = true;
            this.groupBox3.Controls.Add(this.changeColor);
            this.groupBox3.Controls.Add(this.endsentBtn);
            this.groupBox3.Controls.Add(this.startsentBtn);
            this.groupBox3.Controls.Add(this.insertBtn);
            this.groupBox3.Controls.Add(this.btnSetSpacing);
            this.groupBox3.Controls.Add(this.aligenmentBtn);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(12, 527);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox3.Size = new System.Drawing.Size(253, 204);
            this.groupBox3.TabIndex = 20;
            this.groupBox3.TabStop = false;
            // 
            // changeColor
            // 
            this.changeColor.Location = new System.Drawing.Point(168, 43);
            this.changeColor.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.changeColor.Name = "changeColor";
            this.changeColor.Size = new System.Drawing.Size(79, 55);
            this.changeColor.TabIndex = 19;
            this.changeColor.Text = "color";
            this.changeColor.UseVisualStyleBackColor = true;
            this.changeColor.Visible = false;
            this.changeColor.Click += new System.EventHandler(this.changeColor_Click);
            // 
            // btnSetSpacing
            // 
            this.btnSetSpacing.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSetSpacing.Location = new System.Drawing.Point(91, 118);
            this.btnSetSpacing.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnSetSpacing.Name = "btnSetSpacing";
            this.btnSetSpacing.Size = new System.Drawing.Size(156, 55);
            this.btnSetSpacing.TabIndex = 32;
            this.btnSetSpacing.Text = "Set Spacing";
            this.btnSetSpacing.UseVisualStyleBackColor = true;
            this.btnSetSpacing.Visible = false;
            this.btnSetSpacing.Click += new System.EventHandler(this.btnSetSpacing_Click);
            // 
            // aligenmentBtn
            // 
            this.aligenmentBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.aligenmentBtn.Location = new System.Drawing.Point(0, 43);
            this.aligenmentBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.aligenmentBtn.Name = "aligenmentBtn";
            this.aligenmentBtn.Size = new System.Drawing.Size(146, 55);
            this.aligenmentBtn.TabIndex = 30;
            this.aligenmentBtn.Text = "aligenment";
            this.aligenmentBtn.UseVisualStyleBackColor = true;
            this.aligenmentBtn.Visible = false;
            this.aligenmentBtn.Click += new System.EventHandler(this.aligenmentBtn_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox4.AutoSize = true;
            this.groupBox4.Controls.Add(this.btnToggleLanguage);
            this.groupBox4.Controls.Add(this.BIUbtn);
            this.groupBox4.Controls.Add(this.fontGroupBtn);
            this.groupBox4.Controls.Add(this.startparBtn);
            this.groupBox4.Controls.Add(this.endparBtn);
            this.groupBox4.Controls.Add(this.parGroup);
            this.groupBox4.Controls.Add(this.rtlBtn);
            this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.Location = new System.Drawing.Point(1586, 398);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox4.Size = new System.Drawing.Size(351, 341);
            this.groupBox4.TabIndex = 21;
            this.groupBox4.TabStop = false;
            // 
            // btnToggleLanguage
            // 
            this.btnToggleLanguage.Location = new System.Drawing.Point(211, 251);
            this.btnToggleLanguage.Name = "btnToggleLanguage";
            this.btnToggleLanguage.Size = new System.Drawing.Size(115, 58);
            this.btnToggleLanguage.TabIndex = 30;
            this.btnToggleLanguage.Text = "eng/per";
            this.btnToggleLanguage.UseVisualStyleBackColor = true;
            this.btnToggleLanguage.Click += new System.EventHandler(this.btnToggleLanguage_Click);
            // 
            // BIUbtn
            // 
            this.BIUbtn.Location = new System.Drawing.Point(67, 79);
            this.BIUbtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
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
            this.fontGroupBtn.Location = new System.Drawing.Point(19, 102);
            this.fontGroupBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.fontGroupBtn.Name = "fontGroupBtn";
            this.fontGroupBtn.Size = new System.Drawing.Size(93, 55);
            this.fontGroupBtn.TabIndex = 6;
            this.fontGroupBtn.Text = "font";
            this.fontGroupBtn.UseVisualStyleBackColor = true;
            this.fontGroupBtn.Click += new System.EventHandler(this.fontGroupBtn_Click);
            // 
            // rtlBtn
            // 
            this.rtlBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rtlBtn.Location = new System.Drawing.Point(6, 143);
            this.rtlBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.rtlBtn.Name = "rtlBtn";
            this.rtlBtn.Size = new System.Drawing.Size(128, 55);
            this.rtlBtn.TabIndex = 29;
            this.rtlBtn.Text = "RTL/LTR";
            this.rtlBtn.UseVisualStyleBackColor = true;
            this.rtlBtn.Visible = false;
            this.rtlBtn.Click += new System.EventHandler(this.rtlBtn_Click);
            // 
            // filesListBox
            // 
            this.filesListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.filesListBox.FormattingEnabled = true;
            this.filesListBox.ItemHeight = 29;
            this.filesListBox.Location = new System.Drawing.Point(1140, 242);
            this.filesListBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.filesListBox.Name = "filesListBox";
            this.filesListBox.Size = new System.Drawing.Size(369, 120);
            this.filesListBox.TabIndex = 22;
            this.filesListBox.SelectedIndexChanged += new System.EventHandler(this.ListBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(773, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(151, 29);
            this.label1.TabIndex = 23;
            this.label1.Text = "record status";
            this.label1.Visible = false;
            // 
            // titleTextBox
            // 
            this.titleTextBox.Enabled = false;
            this.titleTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titleTextBox.Location = new System.Drawing.Point(245, 119);
            this.titleTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.titleTextBox.Multiline = true;
            this.titleTextBox.Name = "titleTextBox";
            this.titleTextBox.Size = new System.Drawing.Size(676, 34);
            this.titleTextBox.TabIndex = 25;
            this.titleTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox2_keyPress);
            this.titleTextBox.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.title_PreviewKeyDown);
            // 
            // settingBtn
            // 
            this.settingBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.settingBtn.Location = new System.Drawing.Point(21, 281);
            this.settingBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.settingBtn.Name = "settingBtn";
            this.settingBtn.Size = new System.Drawing.Size(127, 55);
            this.settingBtn.TabIndex = 28;
            this.settingBtn.Text = "setting";
            this.settingBtn.UseVisualStyleBackColor = true;
            this.settingBtn.Click += new System.EventHandler(this.settingBtn_Click);
            // 
            // MainrichTextBox
            // 
            this.MainrichTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MainrichTextBox.Location = new System.Drawing.Point(287, 158);
            this.MainrichTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MainrichTextBox.Name = "MainrichTextBox";
            this.MainrichTextBox.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.MainrichTextBox.Size = new System.Drawing.Size(616, 302);
            this.MainrichTextBox.TabIndex = 31;
            this.MainrichTextBox.Text = "";
            this.MainrichTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
            this.MainrichTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            this.MainrichTextBox.KeyUp += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyUp);
            this.MainrichTextBox.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.textBox1_PreviewKeyDown);
            // 
            // spacingValue
            // 
            this.spacingValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.spacingValue.Location = new System.Drawing.Point(948, 29);
            this.spacingValue.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.spacingValue.Multiline = true;
            this.spacingValue.Name = "spacingValue";
            this.spacingValue.Size = new System.Drawing.Size(153, 34);
            this.spacingValue.TabIndex = 33;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(948, 281);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(103, 55);
            this.button1.TabIndex = 34;
            this.button1.Text = "باز کردن فرمت";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.UseWaitCursor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // exportBtn
            // 
            this.exportBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportBtn.Location = new System.Drawing.Point(948, 206);
            this.exportBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.exportBtn.Name = "exportBtn";
            this.exportBtn.Size = new System.Drawing.Size(103, 58);
            this.exportBtn.TabIndex = 8;
            this.exportBtn.Text = "export";
            this.exportBtn.UseVisualStyleBackColor = true;
            this.exportBtn.Visible = false;
            this.exportBtn.Click += new System.EventHandler(this.exportBtn_Click);
            // 
            // bMarkList
            // 
            this.bMarkList.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bMarkList.FormattingEnabled = true;
            this.bMarkList.ItemHeight = 29;
            this.bMarkList.Location = new System.Drawing.Point(1140, 406);
            this.bMarkList.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.bMarkList.Name = "bMarkList";
            this.bMarkList.Size = new System.Drawing.Size(369, 120);
            this.bMarkList.TabIndex = 36;
            this.bMarkList.SelectedIndexChanged += new System.EventHandler(this.bMarkList_SelectedIndexChanged);
            this.bMarkList.KeyDown += new System.Windows.Forms.KeyEventHandler(this.bMarkList_KeyDown);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1924, 740);
            this.Controls.Add(this.bMarkList);
            this.Controls.Add(this.exportBtn);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.spacingValue);
            this.Controls.Add(this.MainrichTextBox);
            this.Controls.Add(this.settingBtn);
            this.Controls.Add(this.titleTextBox);
            this.Controls.Add(this.gotoBtn);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.filesListBox);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
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


        #endregion

        private System.Windows.Forms.Button searchBtn;
        private System.Windows.Forms.Button startparBtn;
        private System.Windows.Forms.Button parGroup;
        private System.Windows.Forms.Button endparBtn;
        private System.Windows.Forms.Button startsentBtn;
        private System.Windows.Forms.Button bookMarkBtn;
        private System.Windows.Forms.Button endsentBtn;
        private System.Windows.Forms.Button starttextBtn;
        private System.Windows.Forms.Button exitBtn;
        private System.Windows.Forms.Button endtextBtn;
        private System.Windows.Forms.Button gotoBtn;
        private System.Windows.Forms.Button gotoBmark;
        private System.Windows.Forms.Button insertBtn;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button fileManagment;
        private System.Windows.Forms.Button printBtn;
        private System.Windows.Forms.ListBox filesListBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox titleTextBox;
        private System.Windows.Forms.Button saveBtn;

        private System.Windows.Forms.Button saveAsBtn;
        private System.Windows.Forms.Button savePdfBtn;
        private System.Windows.Forms.Button statusBtn;
        private System.Windows.Forms.Button changeColor;
        private System.Windows.Forms.Button BIUbtn;
        private System.Windows.Forms.Button fontGroupBtn;
        private System.Windows.Forms.ComboBox fontComboBox;
        private System.Windows.Forms.ComboBox fontSizeComboBox;
        private System.Windows.Forms.Button settingBtn;
        private System.Windows.Forms.Button rtlBtn;
        private System.Windows.Forms.Button aligenmentBtn;
        private System.Windows.Forms.RichTextBox MainrichTextBox;
        private System.Windows.Forms.Button btnSetSpacing;
        private System.Windows.Forms.TextBox spacingValue;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button exportBtn;
        private System.Windows.Forms.Button btnInsertPageBreak;
        private System.Windows.Forms.ListBox bMarkList;
        private System.Windows.Forms.Button btnToggleLanguage;
    }
}

