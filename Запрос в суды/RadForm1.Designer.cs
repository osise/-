
namespace Запрос_в_суды
{
    partial class RadForm1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RadForm1));
            this.fluentTheme1 = new Telerik.WinControls.Themes.FluentTheme();
            this.radButton1 = new Telerik.WinControls.UI.RadButton();
            this.radRichTextEditor1 = new Telerik.WinControls.UI.RadTextBoxControl();
            this.radButton6 = new Telerik.WinControls.UI.RadButton();
            this.radButton5 = new Telerik.WinControls.UI.RadButton();
            this.progressBar1 = new Telerik.WinControls.UI.RadProgressBar();
            this.radButton4 = new Telerik.WinControls.UI.RadButton();
            this.radButton2 = new Telerik.WinControls.UI.RadButton();
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radRichTextEditor1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.progressBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // radButton1
            // 
            this.radButton1.Font = new System.Drawing.Font("Maiandra GD", 12F);
            this.radButton1.Location = new System.Drawing.Point(11, 309);
            this.radButton1.Name = "radButton1";
            this.radButton1.Size = new System.Drawing.Size(180, 41);
            this.radButton1.TabIndex = 28;
            this.radButton1.Text = "Открыть обработанные файлы";
            this.radButton1.ThemeName = "Fluent";
            this.radButton1.Click += new System.EventHandler(this.radButton1_Click);
            // 
            // radRichTextEditor1
            // 
            this.radRichTextEditor1.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.radRichTextEditor1.Location = new System.Drawing.Point(210, 12);
            this.radRichTextEditor1.Multiline = true;
            this.radRichTextEditor1.Name = "radRichTextEditor1";
            this.radRichTextEditor1.Size = new System.Drawing.Size(455, 338);
            this.radRichTextEditor1.TabIndex = 27;
            this.radRichTextEditor1.ThemeName = "Fluent";
            // 
            // radButton6
            // 
            this.radButton6.Font = new System.Drawing.Font("Maiandra GD", 12F);
            this.radButton6.Location = new System.Drawing.Point(12, 190);
            this.radButton6.Name = "radButton6";
            this.radButton6.Size = new System.Drawing.Size(180, 41);
            this.radButton6.TabIndex = 26;
            this.radButton6.Text = "Остановить";
            this.radButton6.ThemeName = "Fluent";
            this.radButton6.Click += new System.EventHandler(this.radButton6_Click);
            // 
            // radButton5
            // 
            this.radButton5.Font = new System.Drawing.Font("Maiandra GD", 12F);
            this.radButton5.Location = new System.Drawing.Point(12, 133);
            this.radButton5.Name = "radButton5";
            this.radButton5.Size = new System.Drawing.Size(180, 41);
            this.radButton5.TabIndex = 25;
            this.radButton5.Text = "Начать";
            this.radButton5.ThemeName = "Fluent";
            this.radButton5.Click += new System.EventHandler(this.radButton5_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(210, 367);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(455, 41);
            this.progressBar1.TabIndex = 24;
            this.progressBar1.ThemeName = "Fluent";
            // 
            // radButton4
            // 
            this.radButton4.Font = new System.Drawing.Font("Maiandra GD", 12F);
            this.radButton4.Location = new System.Drawing.Point(11, 367);
            this.radButton4.Name = "radButton4";
            this.radButton4.Size = new System.Drawing.Size(180, 41);
            this.radButton4.TabIndex = 23;
            this.radButton4.Text = "Выход";
            this.radButton4.ThemeName = "Fluent";
            this.radButton4.Click += new System.EventHandler(this.radButton4_Click);
            // 
            // radButton2
            // 
            this.radButton2.Font = new System.Drawing.Font("Maiandra GD", 12F);
            this.radButton2.Location = new System.Drawing.Point(12, 12);
            this.radButton2.Name = "radButton2";
            this.radButton2.Size = new System.Drawing.Size(180, 41);
            this.radButton2.TabIndex = 22;
            this.radButton2.Text = "Открыть папку с файлами";
            this.radButton2.ThemeName = "Fluent";
            this.radButton2.Click += new System.EventHandler(this.radButton2_Click);
            // 
            // RadForm1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(679, 423);
            this.Controls.Add(this.radButton1);
            this.Controls.Add(this.radRichTextEditor1);
            this.Controls.Add(this.radButton6);
            this.Controls.Add(this.radButton5);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.radButton4);
            this.Controls.Add(this.radButton2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "RadForm1";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.Text = "   Создание запросов для судов v1.1";
            this.ThemeName = "Fluent";
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radRichTextEditor1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.progressBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Telerik.WinControls.Themes.FluentTheme fluentTheme1;
        private Telerik.WinControls.UI.RadButton radButton1;
        private Telerik.WinControls.UI.RadTextBoxControl radRichTextEditor1;
        private Telerik.WinControls.UI.RadButton radButton6;
        private Telerik.WinControls.UI.RadButton radButton5;
        private Telerik.WinControls.UI.RadProgressBar progressBar1;
        private Telerik.WinControls.UI.RadButton radButton4;
        private Telerik.WinControls.UI.RadButton radButton2;
    }
}