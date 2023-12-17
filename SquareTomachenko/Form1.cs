using Microsoft.Office.Interop.Word;
using System.Data;
using System.Diagnostics;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using Application = Microsoft.Office.Interop.Word.Application;
using Excel = Microsoft.Office.Interop.Excel;
using iText.Kernel.Pdf;
using iText.Kernel.Font;
using iText.Layout;
using iText.Layout.Element;

using Word = Microsoft.Office.Interop.Word;

using iText.IO.Font;

namespace SquareTomachenko
{
    public partial class Form1 : Form
    {
        private double sideLength = 0.0;
        private double length = 0.0;
        private double width = 0.0;
        private double height = 0.0;
        private double radius = 0.0;
        private double coneHeight = 0.0;
        private double cylinderHeight = 0.0;
        private double result = 0.0;
        private string resultText = "";

        public Form1()
        {
            InitializeComponent();
            textBoxFigure.Visible = false;
            textBoxFigure1.Visible = false;
            textBoxFigure2.Visible = false;
            comboBox1.SelectedIndex = 0;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem.ToString() == "Куб")
            {
                textBoxFigure.Text = textBoxFigure1.Text = textBoxFigure2.Text = labelResult.Text = "";
                labelFigure1.Text = "";
                labelFigure2.Text = "";
                textBoxFigure1.Visible = false;
                textBoxFigure2.Visible = false;
                labelFigure.Text = "Длина ребра";
                textBoxFigure.Visible = true;
            }

            if (comboBox1.SelectedItem.ToString() == "Прямоугольный параллелепипед")
            {
                textBoxFigure.Text = textBoxFigure1.Text = textBoxFigure2.Text = labelResult.Text = "";
                labelFigure.Text = "Длина";
                labelFigure1.Text = "Ширина";
                labelFigure2.Text = "Высота";
                textBoxFigure.Visible = true;
                textBoxFigure1.Visible = true;
                textBoxFigure2.Visible = true;
            }
            if (comboBox1.SelectedItem.ToString() == "Конус")
            {
                textBoxFigure.Text = textBoxFigure1.Text = textBoxFigure2.Text = labelResult.Text = "";
                labelFigure2.Text = "";
                textBoxFigure2.Visible = false;
                labelFigure.Text = "Радиус основания";
                labelFigure1.Text = "Высота";
                textBoxFigure.Visible = true;
                textBoxFigure1.Visible = true;
            }
            if (comboBox1.SelectedItem.ToString() == "Цилиндр")
            {
                textBoxFigure.Text = textBoxFigure1.Text = textBoxFigure2.Text = labelResult.Text = "";
                labelFigure2.Text = "";
                textBoxFigure2.Visible = false;
                labelFigure.Text = "Радиус основания";
                labelFigure1.Text = "Высота";
                textBoxFigure.Visible = true;
                textBoxFigure1.Visible = true;
            }
        }

        private bool IsValidNumber(string input)
        {
            return double.TryParse(input, out _);
        }

        private void buttonResult_Click(object sender, EventArgs e)
        {
            result = 0.0;

            if (comboBox1.SelectedItem == null || string.IsNullOrEmpty(textBoxFigure.Text) || !double.TryParse(textBoxFigure.Text, out sideLength))
            {
                MessageBox.Show("Введите корректные значения!");
                textBoxFigure.Text = "";
                textBoxFigure1.Text = "";
                textBoxFigure2.Text = "";
                return;
            }

            if (comboBox1.SelectedItem.ToString() == "Куб")
            {
                result = 6 * sideLength * sideLength;
                labelResult.Text = ($"Площадь куба равна: {result:F2}");
            }
            else if (comboBox1.SelectedItem.ToString() == "Прямоугольный параллелепипед")
            {
                if (string.IsNullOrEmpty(textBoxFigure1.Text) || string.IsNullOrEmpty(textBoxFigure2.Text) ||
                    !double.TryParse(textBoxFigure1.Text, out width) || !double.TryParse(textBoxFigure2.Text, out height))
                {
                    MessageBox.Show("Введите корректные значения!");
                    textBoxFigure.Text = "";
                    textBoxFigure1.Text = "";
                    textBoxFigure2.Text = "";
                    return;
                }

                result = 2 * (length * width + length * height + width * height);
                labelResult.Text = ($"Площадь прямоугольного параллелепипеда равна: {result:F2}");
            }
            else if (comboBox1.SelectedItem.ToString() == "Конус")
            {
                if (string.IsNullOrEmpty(textBoxFigure1.Text) || !double.TryParse(textBoxFigure1.Text, out coneHeight) ||
                    !double.TryParse(textBoxFigure.Text, out radius))
                {
                    MessageBox.Show("Введите корректные значения!");
                    textBoxFigure.Text = "";
                    textBoxFigure1.Text = "";
                    textBoxFigure2.Text = "";
                    return;
                }

                result = Math.PI * radius * (radius + Math.Sqrt(radius * radius + coneHeight * coneHeight));
                labelResult.Text = ($"Площадь конуса равна: {result:F2}");
            }
            else if (comboBox1.SelectedItem.ToString() == "Цилиндр")
            {
                if (!double.TryParse(textBoxFigure1.Text, out cylinderHeight) || !double.TryParse(textBoxFigure.Text, out radius))
                {
                    MessageBox.Show("Введите корректные значения!");
                    textBoxFigure.Text = "";
                    textBoxFigure1.Text = "";
                    textBoxFigure2.Text = "";
                    return;
                }

                result = 2 * Math.PI * radius * (radius + cylinderHeight);
                labelResult.Text = ($"Площадь цилиндра равна: {result:F2}");
            }
        }

        private double EvaluateFormula(string formula)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            return Convert.ToDouble(table.Compute(formula, ""));
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            textBoxFigure.Text = "";
            textBoxFigure1.Text = "";
            textBoxFigure2.Text = "";
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private string GetResultsText()
        {
            // Собираем значения и результаты в текстовую строку
            string resultText = "";

            if (comboBox1.SelectedItem.ToString() == "Куб")
            {
                string formula = $"6 * {sideLength:F2} * {sideLength:F2}";
                resultText += $"Фигура: Куб\nДлина ребра: {sideLength:F2}\nПлощадь: {result:F2}\nРешение: {formula:F2}";
            }
            else if (comboBox1.SelectedItem.ToString() == "Прямоугольный параллелепипед")
            {
                string formula = $"Формула: 2 * ({length:F2} * {width:F2} + {length:F2} * {height:F2} + {width:F2} * {height:F2})";
                resultText += $"Фигура: Прямоугольный параллелепипед\nДлина: {length:F2}\nШирина: {width:F2}\nВысота: {height:F2}\nПлощадь: {result:F2}\nРешение: {formula:F2}";
            }
            else if (comboBox1.SelectedItem.ToString() == "Конус")
            {
                string formula = $"Формула: PI * {radius:F2} * ({radius:F2} + ({radius:F2}^2 + {coneHeight:F2}^2)^2)";
                resultText += $"Фигура: Конус\nРадиус основания: {radius:F2}\nВысота: {coneHeight:F2}\nПлощадь: {result:F2}\nРешение: {formula:F2}";
            }
            else if (comboBox1.SelectedItem.ToString() == "Цилиндр")
            {
                string formula = $"Формула: 2 * PI * {radius:F2} * ({radius:F2} + {cylinderHeight:F2})";
                resultText += $"Фигура: Цилиндр\nРадиус основания: {radius:F2}\nВысота: {cylinderHeight:F2}\nПлощадь: {result:F2}\nРешение: {formula:F2}";
            }

            return resultText;
        }

        private void buttonShowToWord_Click(object sender, EventArgs e)
        {
            var wordApp = new Application();
            wordApp.Visible = false;

            // Создаем новый документ
            var doc = wordApp.Documents.Add();

            // Добавляем заголовок
            var paraTitle = doc.Paragraphs.Add();
            paraTitle.Range.Text = "Результаты вычислений";
            paraTitle.Range.Font.Size = 14;
            paraTitle.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; // Изменено на выравнивание по левому краю
            paraTitle.Range.InsertParagraphAfter();

            // Добавляем значения и результаты в документ
            var paraContent = doc.Paragraphs.Add();
            paraContent.Range.Text = GetResultsText();
            paraContent.Range.Font.Size = 14;
            paraContent.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; // Изменено на выравнивание по левому краю
            paraContent.Range.InsertParagraphAfter();

            // Указываем путь и имя файла в папке с проектом
            string projectFolderPath = AppDomain.CurrentDomain.BaseDirectory;
            string projectFilePath = Path.Combine(projectFolderPath, "Результаты.docx");

            // Сохраняем документ
            doc.SaveAs2(projectFilePath);

            // Закрываем Word и документ
            doc.Close();
            Marshal.ReleaseComObject(doc);
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            MessageBox.Show("Файл успешно сохранен");
        }

        private void buttonShowToExcel_Click(object sender, EventArgs e)
        {
            string projectFolderPath = AppDomain.CurrentDomain.BaseDirectory;

            // Формируем полный путь к файлу Excel в папке проекта
            string filePath = Path.Combine(projectFolderPath, "Результаты.xlsx");

            // Создаем новый объект Excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            // Создаем новую книгу
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            // Добавляем заголовок
            worksheet.Cells[1, 1] = "Результаты вычислений";

            // Добавляем данные
            string resultText = GetResultsText(); // Метод, который возвращает текст для всех данных
            worksheet.Cells[2, 1] = resultText;
            Excel.Range columnA = worksheet.Columns["A"];
            columnA.AutoFit();
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            // Сохраняем файл
            workbook.SaveAs(filePath);

            // Закрываем Excel
            workbook.Close();
            Marshal.ReleaseComObject(workbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            MessageBox.Show("Файл успешно сохранен");
        }

        private void SaveToPdf(string text)
        {
            // Указываем путь и имя файла в папке с проектом
            string projectFolderPath = AppDomain.CurrentDomain.BaseDirectory;
            string projectFilePath = Path.Combine(projectFolderPath, "Результаты.pdf");

            // Создаем новый документ PDF
            using (PdfWriter writer = new PdfWriter(projectFilePath))
            {
                using (PdfDocument pdf = new PdfDocument(writer))
                {
                    iText.Layout.Document document = new iText.Layout.Document(pdf);

                    PdfFont timesFont = PdfFontFactory.CreateFont("c:/windows/fonts/times.ttf", PdfEncodings.IDENTITY_H, true);
                    iText.Layout.Element.Paragraph titleParagraph = new iText.Layout.Element.Paragraph("Результаты вычислений");
                    document.Add(titleParagraph.SetFont(timesFont));
                    // Добавляем текст в документ
                    iText.Layout.Element.Paragraph paragraph = new iText.Layout.Element.Paragraph(text);
                    document.Add(paragraph.SetFont(timesFont));
                }
            }

            MessageBox.Show("Данные успешно сохранены в PDF файл!");
        }

        private void buttonShowToPDF_Click(object sender, EventArgs e)
        {
            string resultText = GetResultsText();

            // Сохраняем результаты в PDF
            SaveToPdf(resultText);
        }
    }
}