namespace Selection.Windows
{
    partial class CurvesStyle
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
            this.lb_CurvesStyle = new System.Windows.Forms.ListBox();
            this.b_OK = new System.Windows.Forms.Button();
            this.b_Cancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lb_CurvesStyle
            // 
            this.lb_CurvesStyle.FormattingEnabled = true;
            this.lb_CurvesStyle.Location = new System.Drawing.Point(120, 51);
            this.lb_CurvesStyle.Name = "lb_CurvesStyle";
            this.lb_CurvesStyle.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lb_CurvesStyle.Size = new System.Drawing.Size(247, 108);
            this.lb_CurvesStyle.TabIndex = 0;
            // 
            // b_OK
            // 
            this.b_OK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.b_OK.Location = new System.Drawing.Point(155, 190);
            this.b_OK.Name = "b_OK";
            this.b_OK.Size = new System.Drawing.Size(75, 23);
            this.b_OK.TabIndex = 1;
            this.b_OK.Text = "OK";
            this.b_OK.UseVisualStyleBackColor = true;
            // 
            // b_Cancel
            // 
            this.b_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.b_Cancel.Location = new System.Drawing.Point(255, 190);
            this.b_Cancel.Name = "b_Cancel";
            this.b_Cancel.Size = new System.Drawing.Size(75, 23);
            this.b_Cancel.TabIndex = 2;
            this.b_Cancel.Text = "Отмена";
            this.b_Cancel.UseVisualStyleBackColor = true;
            // 
            // CurvesStyle
            // 
            this.AcceptButton = this.b_OK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.b_Cancel;
            this.ClientSize = new System.Drawing.Size(477, 259);
            this.Controls.Add(this.b_Cancel);
            this.Controls.Add(this.b_OK);
            this.Controls.Add(this.lb_CurvesStyle);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CurvesStyle";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Тип линии";
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.ListBox lb_CurvesStyle;
        internal System.Windows.Forms.Button b_OK;
        internal System.Windows.Forms.Button b_Cancel;
    }
}