namespace WinFormsProsecApp
{
    partial class Form2
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
            SubmitAttendancedetails = new Button();
            CompanyNameList = new ComboBox();
            label1 = new Label();
            label2 = new Label();
            ClientList = new ComboBox();
            Month = new Label();
            MonthList = new ComboBox();
            Yearlist = new Label();
            YearList1 = new ComboBox();
            SuspendLayout();
            // 
            // SubmitAttendancedetails
            // 
            SubmitAttendancedetails.Location = new Point(438, 151);
            SubmitAttendancedetails.Name = "SubmitAttendancedetails";
            SubmitAttendancedetails.Size = new Size(75, 23);
            SubmitAttendancedetails.TabIndex = 0;
            SubmitAttendancedetails.Text = "Submit";
            SubmitAttendancedetails.UseVisualStyleBackColor = true;
            // 
            // CompanyNameList
            // 
            CompanyNameList.AllowDrop = true;
            CompanyNameList.FormattingEnabled = true;
            CompanyNameList.Location = new Point(133, 41);
            CompanyNameList.Name = "CompanyNameList";
            CompanyNameList.Size = new Size(121, 23);
            CompanyNameList.TabIndex = 1;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(28, 47);
            label1.Name = "label1";
            label1.Size = new Size(94, 15);
            label1.TabIndex = 2;
            label1.Text = "Company Name";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(285, 47);
            label2.Name = "label2";
            label2.Size = new Size(73, 15);
            label2.TabIndex = 4;
            label2.Text = "Client Name";
            // 
            // ClientList
            // 
            ClientList.FormattingEnabled = true;
            ClientList.Location = new Point(390, 41);
            ClientList.Name = "ClientList";
            ClientList.Size = new Size(121, 23);
            ClientList.TabIndex = 3;
            ClientList.SelectedIndexChanged += ClientList_SelectedIndexChanged;
            // 
            // Month
            // 
            Month.AutoSize = true;
            Month.Location = new Point(74, 108);
            Month.Name = "Month";
            Month.Size = new Size(43, 15);
            Month.TabIndex = 6;
            Month.Text = "Month";
            // 
            // MonthList
            // 
            MonthList.FormattingEnabled = true;
            MonthList.Location = new Point(132, 102);
            MonthList.Name = "MonthList";
            MonthList.Size = new Size(121, 23);
            MonthList.TabIndex = 5;
            MonthList.SelectedIndexChanged += comboBox3_SelectedIndexChanged;
            // 
            // Yearlist
            // 
            Yearlist.AutoSize = true;
            Yearlist.Location = new Point(287, 99);
            Yearlist.Name = "Yearlist";
            Yearlist.Size = new Size(29, 15);
            Yearlist.TabIndex = 8;
            Yearlist.Text = "Year";
            // 
            // YearList1
            // 
            YearList1.FormattingEnabled = true;
            YearList1.Location = new Point(392, 93);
            YearList1.Name = "YearList1";
            YearList1.Size = new Size(121, 23);
            YearList1.TabIndex = 7;
            // 
            // Form2
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(552, 204);
            Controls.Add(Yearlist);
            Controls.Add(YearList1);
            Controls.Add(Month);
            Controls.Add(MonthList);
            Controls.Add(label2);
            Controls.Add(ClientList);
            Controls.Add(label1);
            Controls.Add(CompanyNameList);
            Controls.Add(SubmitAttendancedetails);
            Name = "Form2";
            Text = "Form2";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button SubmitAttendancedetails;
        private ComboBox CompanyNameList;
        private Label label1;
        private Label label2;
        private ComboBox ClientList;
        private Label Month;
        private ComboBox MonthList;
        private Label Yearlist;
        private ComboBox YearList1;
    }
}