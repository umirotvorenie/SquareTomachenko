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
            if (comboBox1.SelectedItem.ToString() == "���")
            {
                textBoxFigure.Text = textBoxFigure1.Text = textBoxFigure2.Text = labelResult.Text = "";
                labelFigure1.Text = "";
                labelFigure2.Text = "";
                textBoxFigure1.Visible = false;
                textBoxFigure2.Visible = false;
                labelFigure.Text = "����� �����";
                textBoxFigure.Visible = true;
            }

            if (comboBox1.SelectedItem.ToString() == "������������� ��������������")
            {
                textBoxFigure.Text = textBoxFigure1.Text = textBoxFigure2.Text = labelResult.Text = "";
                labelFigure.Text = "�����";
                labelFigure1.Text = "������";
                labelFigure2.Text = "������";
                textBoxFigure.Visible = true;
                textBoxFigure1.Visible = true;
                textBoxFigure2.Visible = true;
            }
            if (comboBox1.SelectedItem.ToString() == "�����")
            {
                textBoxFigure.Text = textBoxFigure1.Text = textBoxFigure2.Text = labelResult.Text = "";
                labelFigure2.Text = "";
                textBoxFigure2.Visible = false;
                labelFigure.Text = "������ ���������";
                labelFigure1.Text = "������";
                textBoxFigure.Visible = true;
                textBoxFigure1.Visible = true;
            }
            if (comboBox1.SelectedItem.ToString() == "�������")
            {
                textBoxFigure.Text = textBoxFigure1.Text = textBoxFigure2.Text = labelResult.Text = "";
                labelFigure2.Text = "";
                textBoxFigure2.Visible = false;
                labelFigure.Text = "������ ���������";
                labelFigure1.Text = "������";
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
                MessageBox.Show("������� ���������� ��������!");
                textBoxFigure.Text = "";
                textBoxFigure1.Text = "";
                textBoxFigure2.Text = "";
                return;
            }

            if (comboBox1.SelectedItem.ToString() == "���")
            {
                result = 6 * sideLength * sideLength;
                labelResult.Text = ($"������� ���� �����: {result:F2}");
            }
            else if (comboBox1.SelectedItem.ToString() == "������������� ��������������")
            {
                if (string.IsNullOrEmpty(textBoxFigure1.Text) || string.IsNullOrEmpty(textBoxFigure2.Text) ||
                    !double.TryParse(textBoxFigure1.Text, out width) || !double.TryParse(textBoxFigure2.Text, out height))
                {
                    MessageBox.Show("������� ���������� ��������!");
                    textBoxFigure.Text = "";
                    textBoxFigure1.Text = "";
                    textBoxFigure2.Text = "";
                    return;
                }

                result = 2 * (length * width + length * height + width * height);
                labelResult.Text = ($"������� �������������� ��������������� �����: {result:F2}");
            }
            else if (comboBox1.SelectedItem.ToString() == "�����")
            {
                if (string.IsNullOrEmpty(textBoxFigure1.Text) || !double.TryParse(textBoxFigure1.Text, out coneHeight) ||
                    !double.TryParse(textBoxFigure.Text, out radius))
                {
                    MessageBox.Show("������� ���������� ��������!");
                    textBoxFigure.Text = "";
                    textBoxFigure1.Text = "";
                    textBoxFigure2.Text = "";
                    return;
                }

                result = Math.PI * radius * (radius + Math.Sqrt(radius * radius + coneHeight * coneHeight));
                labelResult.Text = ($"������� ������ �����: {result:F2}");
            }
            else if (comboBox1.SelectedItem.ToString() == "�������")
            {
                if (!double.TryParse(textBoxFigure1.Text, out cylinderHeight) || !double.TryParse(textBoxFigure.Text, out radius))
                {
                    MessageBox.Show("������� ���������� ��������!");
                    textBoxFigure.Text = "";
                    textBoxFigure1.Text = "";
                    textBoxFigure2.Text = "";
                    return;
                }

                result = 2 * Math.PI * radius * (radius + cylinderHeight);
                labelResult.Text = ($"������� �������� �����: {result:F2}");
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
            // �������� �������� � ���������� � ��������� ������
            string resultText = "";

            if (comboBox1.SelectedItem.ToString() == "���")
            {
                string formula = $"6 * {sideLength:F2} * {sideLength:F2}";
                resultText += $"������: ���\n����� �����: {sideLength:F2}\n�������: {result:F2}\n�������: {formula:F2}";
            }
            else if (comboBox1.SelectedItem.ToString() == "������������� ��������������")
            {
                string formula = $"�������: 2 * ({length:F2} * {width:F2} + {length:F2} * {height:F2} + {width:F2} * {height:F2})";
                resultText += $"������: ������������� ��������������\n�����: {length:F2}\n������: {width:F2}\n������: {height:F2}\n�������: {result:F2}\n�������: {formula:F2}";
            }
            else if (comboBox1.SelectedItem.ToString() == "�����")
            {
                string formula = $"�������: PI * {radius:F2} * ({radius:F2} + ({radius:F2}^2 + {coneHeight:F2}^2)^2)";
                resultText += $"������: �����\n������ ���������: {radius:F2}\n������: {coneHeight:F2}\n�������: {result:F2}\n�������: {formula:F2}";
            }
            else if (comboBox1.SelectedItem.ToString() == "�������")
            {
                string formula = $"�������: 2 * PI * {radius:F2} * ({radius:F2} + {cylinderHeight:F2})";
                resultText += $"������: �������\n������ ���������: {radius:F2}\n������: {cylinderHeight:F2}\n�������: {result:F2}\n�������: {formula:F2}";
            }

            return resultText;
        }

        private void buttonShowToWord_Click(object sender, EventArgs e)
        {
            var wordApp = new Application();
            wordApp.Visible = false;

            // ������� ����� ��������
            var doc = wordApp.Documents.Add();

            // ��������� ���������
            var paraTitle = doc.Paragraphs.Add();
            paraTitle.Range.Text = "���������� ����������";
            paraTitle.Range.Font.Size = 14;
            paraTitle.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; // �������� �� ������������ �� ������ ����
            paraTitle.Range.InsertParagraphAfter();

            // ��������� �������� � ���������� � ��������
            var paraContent = doc.Paragraphs.Add();
            paraContent.Range.Text = GetResultsText();
            paraContent.Range.Font.Size = 14;
            paraContent.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; // �������� �� ������������ �� ������ ����
            paraContent.Range.InsertParagraphAfter();

            // ��������� ���� � ��� ����� � ����� � ��������
            string projectFolderPath = AppDomain.CurrentDomain.BaseDirectory;
            string projectFilePath = Path.Combine(projectFolderPath, "����������.docx");

            // ��������� ��������
            doc.SaveAs2(projectFilePath);

            // ��������� Word � ��������
            doc.Close();
            Marshal.ReleaseComObject(doc);
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            MessageBox.Show("���� ������� ��������");
        }

        private void buttonShowToExcel_Click(object sender, EventArgs e)
        {
            string projectFolderPath = AppDomain.CurrentDomain.BaseDirectory;

            // ��������� ������ ���� � ����� Excel � ����� �������
            string filePath = Path.Combine(projectFolderPath, "����������.xlsx");

            // ������� ����� ������ Excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            // ������� ����� �����
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            // ��������� ���������
            worksheet.Cells[1, 1] = "���������� ����������";

            // ��������� ������
            string resultText = GetResultsText(); // �����, ������� ���������� ����� ��� ���� ������
            worksheet.Cells[2, 1] = resultText;
            Excel.Range columnA = worksheet.Columns["A"];
            columnA.AutoFit();
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            // ��������� ����
            workbook.SaveAs(filePath);

            // ��������� Excel
            workbook.Close();
            Marshal.ReleaseComObject(workbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            MessageBox.Show("���� ������� ��������");
        }

        private void SaveToPdf(string text)
        {
            // ��������� ���� � ��� ����� � ����� � ��������
            string projectFolderPath = AppDomain.CurrentDomain.BaseDirectory;
            string projectFilePath = Path.Combine(projectFolderPath, "����������.pdf");

            // ������� ����� �������� PDF
            using (PdfWriter writer = new PdfWriter(projectFilePath))
            {
                using (PdfDocument pdf = new PdfDocument(writer))
                {
                    iText.Layout.Document document = new iText.Layout.Document(pdf);

                    PdfFont timesFont = PdfFontFactory.CreateFont("c:/windows/fonts/times.ttf", PdfEncodings.IDENTITY_H, true);
                    iText.Layout.Element.Paragraph titleParagraph = new iText.Layout.Element.Paragraph("���������� ����������");
                    document.Add(titleParagraph.SetFont(timesFont));
                    // ��������� ����� � ��������
                    iText.Layout.Element.Paragraph paragraph = new iText.Layout.Element.Paragraph(text);
                    document.Add(paragraph.SetFont(timesFont));
                }
            }

            MessageBox.Show("������ ������� ��������� � PDF ����!");
        }

        private void buttonShowToPDF_Click(object sender, EventArgs e)
        {
            string resultText = GetResultsText();

            // ��������� ���������� � PDF
            SaveToPdf(resultText);
        }
    }
}