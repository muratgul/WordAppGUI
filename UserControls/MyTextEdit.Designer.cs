namespace WordAppGUI.UserControls
{
    partial class MyTextEdit
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtKey = new System.Windows.Forms.TextBox();
            this.txtValue = new System.Windows.Forms.TextBox();
            this.cmbTip = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // txtKey
            // 
            this.txtKey.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtKey.Location = new System.Drawing.Point(1, 0);
            this.txtKey.Name = "txtKey";
            this.txtKey.ReadOnly = true;
            this.txtKey.Size = new System.Drawing.Size(127, 20);
            this.txtKey.TabIndex = 0;
            // 
            // txtValue
            // 
            this.txtValue.Location = new System.Drawing.Point(133, 1);
            this.txtValue.Name = "txtValue";
            this.txtValue.Size = new System.Drawing.Size(127, 20);
            this.txtValue.TabIndex = 0;
            // 
            // cmbTip
            // 
            this.cmbTip.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTip.FormattingEnabled = true;
            this.cmbTip.Items.AddRange(new object[] {
            "Metin",
            "Tarih",
            "Sayı",
            "Ondalık"});
            this.cmbTip.Location = new System.Drawing.Point(269, 2);
            this.cmbTip.Name = "cmbTip";
            this.cmbTip.Size = new System.Drawing.Size(81, 21);
            this.cmbTip.TabIndex = 1;
            // 
            // MyTextEdit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cmbTip);
            this.Controls.Add(this.txtValue);
            this.Controls.Add(this.txtKey);
            this.Name = "MyTextEdit";
            this.Size = new System.Drawing.Size(266, 20);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.TextBox txtKey;
        public System.Windows.Forms.TextBox txtValue;
        public System.Windows.Forms.ComboBox cmbTip;
    }
}
