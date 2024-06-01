using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace Lab_26_Danylko
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Word.Application _wordApp;
        private Document _wordDoc;
        private string _templatePath;

        public Form1()
        {
            InitializeComponent();
            // Додаємо обробники подій для прапорців
            checkBoxEducation1.CheckedChanged += checkBoxEducation1_CheckedChanged;
            checkBoxEducation2.CheckedChanged += checkBoxEducation2_CheckedChanged;
            checkBoxWork1.CheckedChanged += checkBoxWork1_CheckedChanged;
            checkBoxWork2.CheckedChanged += checkBoxWork2_CheckedChanged;
        }

        // Обробник кнопки для генерації
        private void buttonGenerate_Click(object sender, EventArgs e)
        {
            if (!ValidateInput()) return;

            if (!string.IsNullOrEmpty(textBoxTemplatePath.Text))
            {
                _templatePath = textBoxTemplatePath.Text;
            }
            else
            {
                MessageBox.Show("Виберіть шаблон для генерації.");
                return;
            }

            try
            {
                _wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
                _wordDoc = _wordApp.Documents.Add(_templatePath);

                // Заповнення даних у шаблоні
                FillVisitingCardsData();

                _wordApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Помилка: {ex.Message}", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Звільнення об'єктів
                ReleaseWordObject(_wordDoc);
                ReleaseWordObject(_wordApp);
            }
        }

        // Метод для заповнення даних у шаблоні
        private void FillVisitingCardsData()
        {
                ReplaceContentControlText(_wordDoc, $"Name", textBoxName.Text);
                ReplaceContentControlText(_wordDoc, $"Address", textBoxAddress.Text);
                ReplaceContentControlText(_wordDoc, $"Region-index", textBoxRegionIndex.Text);
                ReplaceContentControlText(_wordDoc, $"Phone", textBoxPhone.Text);
                ReplaceContentControlText(_wordDoc, $"Email", textBoxEmail.Text);
                ReplaceContentControlText(_wordDoc, $"Aim", textBoxAim.Text);

                if (checkBoxEducation1.Checked)
                {
                    ReplaceContentControlText(_wordDoc, $"Education", textBoxEducation.Text);
                    ReplaceContentControlText(_wordDoc, $"EduYears", textBoxEduYears.Text);
                }
                if (checkBoxEducation2.Checked)
                {
                    ReplaceContentControlText(_wordDoc, $"Education2", textBoxEducation2.Text);
                    ReplaceContentControlText(_wordDoc, $"EduYears2", textBoxEduYears2.Text);
                }
                if (checkBoxWork1.Checked)
                {
                    ReplaceContentControlText(_wordDoc, $"Work1", textBoxWork1.Text);
                    ReplaceContentControlText(_wordDoc, $"Post1", textBoxPost1.Text);
                    ReplaceContentControlText(_wordDoc, $"WorkDate1", textBoxWorkDate1.Text);
                }
                if (checkBoxWork2.Checked)
                {
                    ReplaceContentControlText(_wordDoc, $"Work2", textBoxWork2.Text);
                    ReplaceContentControlText(_wordDoc, $"Post2", textBoxPost2.Text);
                    ReplaceContentControlText(_wordDoc, $"WorkDate2", textBoxWorkDate2.Text);
                }
        }

        // Метод для заміни тексту в елементах управління вмістом за тегом
        private void ReplaceContentControlText(Document doc, string tag, string text)
        {
            foreach (ContentControl cc in doc.ContentControls)
            {
                if (cc.Tag == tag)
                {
                    // Очищаємо існуючий текст у контент-контролі
                    cc.Range.Text = string.Empty;
                    // Додаємо новий текст
                    cc.Range.InsertAfter(text);
                }
            }
        }

        // Метод для звільнення ресурсів Word
        private void ReleaseWordObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                }
                catch
                {
                    // Ігноруємо помилки при звільненні
                }
                finally
                {
                    obj = null;
                }
            }
        }

        // Обробник кнопки для вибору шаблону
        private void buttonSelectTemplate_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Word Templates (*.dotx)|*.dotx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    textBoxTemplatePath.Text = openFileDialog.FileName;
                }
            }
        }

        private string FormatDate(string date)
        {
            // Определяем регулярное выражение для диапазона дат
            string pattern = @"^(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})$";
            Match match = Regex.Match(date, pattern);

            if (match.Success)
            {
                // Парсим и форматируем каждую дату в диапазоне
                if (DateTime.TryParse(match.Groups[1].Value, out DateTime startDate) &&
                    DateTime.TryParse(match.Groups[2].Value, out DateTime endDate))
                {
                    return $"{startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}";
                }
            }
            else
            {
                // Если не диапазон, проверяем одну дату
                if (DateTime.TryParse(date, out DateTime singleDate))
                {
                    return singleDate.ToString("dd.MM.yyyy");
                }
            }

            // Если формат даты неверен
            MessageBox.Show($"Неправильний формат дати: {date}. Використовуйте формат dd.mm.yyyy або dd.mm.yyyy - dd.mm.yyyy.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return date; // повертаємо оригінальну дату, якщо вона не вдається перетворити
        }



        // Перевірка введених даних
        private bool ValidateInput()
        {
            if (string.IsNullOrEmpty(textBoxName.Text))
            {
                MessageBox.Show("Введіть ім'я.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (string.IsNullOrEmpty(textBoxAddress.Text))
            {
                MessageBox.Show("Введіть адресу.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (!Regex.IsMatch(textBoxRegionIndex.Text, @"^[^,]+,\s*\d+$"))
            {
                MessageBox.Show("Введіть регіон та індекс у правильному форматі (Регіон, індекс).", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (!Regex.IsMatch(textBoxPhone.Text, @"^[\d\+\-\(\)]+$"))
            {
                MessageBox.Show("Введіть номер телефону, використовуючи тільки цифри, +, -, (, ).", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (!Regex.IsMatch(textBoxEmail.Text, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                MessageBox.Show("Введіть правильну адресу електронної пошти.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (string.IsNullOrEmpty(textBoxAim.Text))
            {
                MessageBox.Show("Введіть мету.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (checkBoxEducation1.Checked)
            {
                if (string.IsNullOrEmpty(textBoxEducation.Text) || !Regex.IsMatch(textBoxEduYears.Text, @"^[\d\.\-\s]+$"))
                {
                    MessageBox.Show("Введіть правильні дані для першого освіти.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                textBoxEduYears.Text = FormatDate(textBoxEduYears.Text);
            }

            if (checkBoxEducation2.Checked)
            {
                if (string.IsNullOrEmpty(textBoxEducation2.Text) || !Regex.IsMatch(textBoxEduYears2.Text, @"^[\d\.\-\s]+$"))
                {
                    MessageBox.Show("Введіть правильні дані для другого освіти.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                textBoxEduYears2.Text = FormatDate(textBoxEduYears2.Text);
            }

            if (checkBoxWork1.Checked)
            {
                if (string.IsNullOrEmpty(textBoxWork1.Text) || string.IsNullOrEmpty(textBoxPost1.Text) || !Regex.IsMatch(textBoxWorkDate1.Text, @"^[\d\.\-\s]+$"))
                {
                    MessageBox.Show("Введіть правильні дані для першого місця роботи.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                textBoxWorkDate1.Text = FormatDate(textBoxWorkDate1.Text);
            }

            if (checkBoxWork2.Checked)
            {
                if (string.IsNullOrEmpty(textBoxWork2.Text) || string.IsNullOrEmpty(textBoxPost2.Text) || !Regex.IsMatch(textBoxWorkDate2.Text, @"^[\d\.\-\s]+$"))
                {
                    MessageBox.Show("Введіть правильні дані для другого місця роботи.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                textBoxWorkDate2.Text = FormatDate(textBoxWorkDate2.Text);
            }

            return true;
        }


        // Обробники подій для прапорців
        private void checkBoxEducation1_CheckedChanged(object sender, EventArgs e)
        {
            textBoxEducation.Enabled = checkBoxEducation1.Checked;
            textBoxEduYears.Enabled = checkBoxEducation1.Checked;
        }

        private void checkBoxEducation2_CheckedChanged(object sender, EventArgs e)
        {
            textBoxEducation2.Enabled = checkBoxEducation2.Checked;
            textBoxEduYears2.Enabled = checkBoxEducation2.Checked;
        }

        private void checkBoxWork1_CheckedChanged(object sender, EventArgs e)
        {
            textBoxWork1.Enabled = checkBoxWork1.Checked;
            textBoxPost1.Enabled = checkBoxWork1.Checked;
            textBoxWorkDate1.Enabled = checkBoxWork1.Checked;
        }

        private void checkBoxWork2_CheckedChanged(object sender, EventArgs e)
        {
            textBoxWork2.Enabled = checkBoxWork2.Checked;
            textBoxPost2.Enabled = checkBoxWork2.Checked;
            textBoxWorkDate2.Enabled = checkBoxWork2.Checked;
        }
    }
}
