
namespace Test.UI
{
    partial class TestForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            CreateDocumentBtn = new Button();
            SuspendLayout();
            // 
            // CreateDocumentBtn
            // 
            CreateDocumentBtn.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point,  162);
            CreateDocumentBtn.Location = new Point(318, 217);
            CreateDocumentBtn.Name = "CreateDocumentBtn";
            CreateDocumentBtn.Size = new Size(175, 44);
            CreateDocumentBtn.TabIndex = 0;
            CreateDocumentBtn.Text = "Doküman Oluştur";
            CreateDocumentBtn.UseVisualStyleBackColor = true;
            CreateDocumentBtn.Click += this.CreateDocumentBtn_Click;
            // 
            // TestForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(CreateDocumentBtn);
            Name = "TestForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Form1";
            ResumeLayout(false);
        }



        #endregion

        private Button CreateDocumentBtn;
    }
}
