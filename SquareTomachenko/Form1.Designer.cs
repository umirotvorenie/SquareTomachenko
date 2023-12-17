namespace SquareTomachenko
{
    partial class Form1
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
            comboBox1 = new ComboBox();
            buttonResult = new Button();
            buttonClear = new Button();
            buttonExit = new Button();
            buttonShowToWord = new Button();
            buttonShowToExcel = new Button();
            buttonShowToPDF = new Button();
            textBoxFigure = new TextBox();
            labelFigure = new Label();
            labelResult = new Label();
            textBoxFigure1 = new TextBox();
            textBoxFigure2 = new TextBox();
            labelFigure1 = new Label();
            labelFigure2 = new Label();
            SuspendLayout();
            // 
            // comboBox1
            // 
            comboBox1.FormattingEnabled = true;
            comboBox1.Items.AddRange(new object[] { "Куб", "Прямоугольный параллелепипед", "Конус", "Цилиндр" });
            comboBox1.Location = new Point(478, 42);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(121, 23);
            comboBox1.TabIndex = 0;
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            // 
            // buttonResult
            // 
            buttonResult.Location = new Point(644, 42);
            buttonResult.Name = "buttonResult";
            buttonResult.Size = new Size(131, 54);
            buttonResult.TabIndex = 1;
            buttonResult.Text = "Рассчитать";
            buttonResult.UseVisualStyleBackColor = true;
            buttonResult.Click += buttonResult_Click;
            // 
            // buttonClear
            // 
            buttonClear.Location = new Point(644, 102);
            buttonClear.Name = "buttonClear";
            buttonClear.Size = new Size(131, 54);
            buttonClear.TabIndex = 2;
            buttonClear.Text = "Очистить";
            buttonClear.UseVisualStyleBackColor = true;
            buttonClear.Click += buttonClear_Click;
            // 
            // buttonExit
            // 
            buttonExit.Location = new Point(644, 342);
            buttonExit.Name = "buttonExit";
            buttonExit.Size = new Size(131, 54);
            buttonExit.TabIndex = 3;
            buttonExit.Text = "Выйти";
            buttonExit.UseVisualStyleBackColor = true;
            buttonExit.Click += buttonExit_Click;
            // 
            // buttonShowToWord
            // 
            buttonShowToWord.Location = new Point(644, 162);
            buttonShowToWord.Name = "buttonShowToWord";
            buttonShowToWord.Size = new Size(131, 54);
            buttonShowToWord.TabIndex = 4;
            buttonShowToWord.Text = "Показать в Word";
            buttonShowToWord.UseVisualStyleBackColor = true;
            buttonShowToWord.Click += buttonShowToWord_Click;
            // 
            // buttonShowToExcel
            // 
            buttonShowToExcel.Location = new Point(644, 222);
            buttonShowToExcel.Name = "buttonShowToExcel";
            buttonShowToExcel.Size = new Size(131, 54);
            buttonShowToExcel.TabIndex = 5;
            buttonShowToExcel.Text = "Показаать в Excel";
            buttonShowToExcel.UseVisualStyleBackColor = true;
            buttonShowToExcel.Click += buttonShowToExcel_Click;
            // 
            // buttonShowToPDF
            // 
            buttonShowToPDF.Location = new Point(644, 282);
            buttonShowToPDF.Name = "buttonShowToPDF";
            buttonShowToPDF.Size = new Size(131, 54);
            buttonShowToPDF.TabIndex = 6;
            buttonShowToPDF.Text = "Покаазать в PDF";
            buttonShowToPDF.UseVisualStyleBackColor = true;
            buttonShowToPDF.Click += buttonShowToPDF_Click;
            // 
            // textBoxFigure
            // 
            textBoxFigure.Location = new Point(27, 41);
            textBoxFigure.Name = "textBoxFigure";
            textBoxFigure.Size = new Size(167, 23);
            textBoxFigure.TabIndex = 7;
            // 
            // labelFigure
            // 
            labelFigure.AutoSize = true;
            labelFigure.Location = new Point(67, 23);
            labelFigure.Name = "labelFigure";
            labelFigure.Size = new Size(0, 15);
            labelFigure.TabIndex = 8;
            // 
            // labelResult
            // 
            labelResult.AutoSize = true;
            labelResult.Location = new Point(81, 173);
            labelResult.Name = "labelResult";
            labelResult.Size = new Size(0, 15);
            labelResult.TabIndex = 9;
            // 
            // textBoxFigure1
            // 
            textBoxFigure1.Location = new Point(27, 87);
            textBoxFigure1.Name = "textBoxFigure1";
            textBoxFigure1.Size = new Size(167, 23);
            textBoxFigure1.TabIndex = 10;
            // 
            // textBoxFigure2
            // 
            textBoxFigure2.Location = new Point(27, 133);
            textBoxFigure2.Name = "textBoxFigure2";
            textBoxFigure2.Size = new Size(167, 23);
            textBoxFigure2.TabIndex = 11;
            // 
            // labelFigure1
            // 
            labelFigure1.AutoSize = true;
            labelFigure1.Location = new Point(67, 69);
            labelFigure1.Name = "labelFigure1";
            labelFigure1.Size = new Size(0, 15);
            labelFigure1.TabIndex = 12;
            // 
            // labelFigure2
            // 
            labelFigure2.AutoSize = true;
            labelFigure2.Location = new Point(67, 116);
            labelFigure2.Name = "labelFigure2";
            labelFigure2.Size = new Size(0, 15);
            labelFigure2.TabIndex = 13;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(labelFigure2);
            Controls.Add(labelFigure1);
            Controls.Add(textBoxFigure2);
            Controls.Add(textBoxFigure1);
            Controls.Add(labelResult);
            Controls.Add(labelFigure);
            Controls.Add(textBoxFigure);
            Controls.Add(buttonShowToPDF);
            Controls.Add(buttonShowToExcel);
            Controls.Add(buttonShowToWord);
            Controls.Add(buttonExit);
            Controls.Add(buttonClear);
            Controls.Add(buttonResult);
            Controls.Add(comboBox1);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ComboBox comboBox1;
        private Button buttonResult;
        private Button buttonClear;
        private Button buttonExit;
        private Button buttonShowToWord;
        private Button buttonShowToExcel;
        private Button buttonShowToPDF;
        private TextBox textBoxFigure;
        private Label labelFigure;
        private Label labelResult;
        private TextBox textBoxFigure1;
        private TextBox textBoxFigure2;
        private Label labelFigure1;
        private Label labelFigure2;
    }
}