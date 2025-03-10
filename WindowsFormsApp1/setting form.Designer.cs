using System.Windows.Forms;

namespace WindowsFormsApp1
{
    partial class setting_form
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
        private void InitializeComponent()
        {
            this.listBoxAllFonts = new System.Windows.Forms.ListBox();
            this.listBoxSelectedFonts = new System.Windows.Forms.ListBox();
            this.buttonAdd = new System.Windows.Forms.Button();
            this.buttonRmove = new System.Windows.Forms.Button();
            this.buttonSave = new System.Windows.Forms.Button();
            this.listBoxSize = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // listBoxAllFonts
            // 
            this.listBoxAllFonts.FormattingEnabled = true;
            this.listBoxAllFonts.ItemHeight = 16;
            this.listBoxAllFonts.Location = new System.Drawing.Point(359, 113);
            this.listBoxAllFonts.Name = "listBoxAllFonts";
            this.listBoxAllFonts.Size = new System.Drawing.Size(209, 260);
            this.listBoxAllFonts.TabIndex = 0;
            // 
            // listBoxSelectedFonts
            // 
            this.listBoxSelectedFonts.FormattingEnabled = true;
            this.listBoxSelectedFonts.ItemHeight = 16;
            this.listBoxSelectedFonts.Location = new System.Drawing.Point(114, 113);
            this.listBoxSelectedFonts.Name = "listBoxSelectedFonts";
            this.listBoxSelectedFonts.Size = new System.Drawing.Size(197, 260);
            this.listBoxSelectedFonts.TabIndex = 1;
            // 
            // buttonAdd
            // 
            this.buttonAdd.Location = new System.Drawing.Point(232, 459);
            this.buttonAdd.Name = "buttonAdd";
            this.buttonAdd.Size = new System.Drawing.Size(75, 23);
            this.buttonAdd.TabIndex = 2;
            this.buttonAdd.Text = "add";
            this.buttonAdd.UseVisualStyleBackColor = true;
            this.buttonAdd.Click += new System.EventHandler(this.ButtonAdd_Click);
            // 
            // buttonRmove
            // 
            this.buttonRmove.Location = new System.Drawing.Point(389, 459);
            this.buttonRmove.Name = "buttonRmove";
            this.buttonRmove.Size = new System.Drawing.Size(75, 23);
            this.buttonRmove.TabIndex = 3;
            this.buttonRmove.Text = "remove";
            this.buttonRmove.UseVisualStyleBackColor = true;
            this.buttonRmove.Click += new System.EventHandler(this.ButtonRemove_Click);
            // 
            // buttonSave
            // 
            this.buttonSave.Location = new System.Drawing.Point(557, 459);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(75, 23);
            this.buttonSave.TabIndex = 4;
            this.buttonSave.Text = "save";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.ButtonSave_Click);
            // 
            // listBoxSize
            // 
            this.listBoxSize.FormattingEnabled = true;
            this.listBoxSize.ItemHeight = 16;
            this.listBoxSize.Items.AddRange(new object[] {
            12,
            14,
            16,
            18,
            20,
            22,
            24,
            26,
            28,
            30,
            36,
            42,
            54,
            68,
            72});
            this.listBoxSize.Location = new System.Drawing.Point(633, 113);
            this.listBoxSize.Name = "listBoxSize";
            this.listBoxSize.Size = new System.Drawing.Size(209, 260);
            this.listBoxSize.TabIndex = 5;
            // 
            // setting_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(933, 549);
            this.Controls.Add(this.listBoxSize);
            this.Controls.Add(this.buttonSave);
            this.Controls.Add(this.buttonRmove);
            this.Controls.Add(this.buttonAdd);
            this.Controls.Add(this.listBoxSelectedFonts);
            this.Controls.Add(this.listBoxAllFonts);
            this.Name = "setting_form";
            this.Text = "setting_form";
            this.Load += new System.EventHandler(this.setting_form_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox listBoxAllFonts;
        private System.Windows.Forms.ListBox listBoxSelectedFonts;
        private System.Windows.Forms.Button buttonAdd;
        private System.Windows.Forms.Button buttonRmove;
        private System.Windows.Forms.Button buttonSave;
        private System.Windows.Forms.ListBox listBoxSize;
    }
}