using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp26
{
    public partial class Form1 : Form
    {
        private DataTable scheduleDataTable;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void LoadTaskScheduleFromFile(string filePath)
        {

            scheduleDataTable = new DataTable();
            scheduleDataTable.Columns.Add("Завдання", typeof(string));
            scheduleDataTable.Columns.Add("Пріорітет", typeof(string));
            scheduleDataTable.Columns.Add("Термін виконання", typeof(string));
            scheduleDataTable.Columns.Add("Стан завдання", typeof(string));

            try
            {
                //Зчитуємо з файлу 
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        string[] scheduleEntry = line.Split(',');
                        if (scheduleEntry.Length == 4)
                        {
                            scheduleDataTable.Rows.Add(scheduleEntry);
                        }
                    }
                }

                dataGridView1.DataSource = scheduleDataTable;
            }
            catch (IOException ex)
            {
                MessageBox.Show("Помилка при читанні файлу: " + ex.Message);
            }
        }

        private void завантажитиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text Files (*.txt)|*.txt";
            openFileDialog.Title = "Виберіть файл ";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedFilePath = openFileDialog.FileName;
                LoadTaskScheduleFromFile(selectedFilePath);
            }
        }

        private void зберегтиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text Files (*.txt)|*.txt";
            saveFileDialog.Title = "Зберегти завдання";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedFilePath = saveFileDialog.FileName;

                try
                {
                    using (StreamWriter writer = new StreamWriter(selectedFilePath))
                    {
                        foreach (DataRow row in scheduleDataTable.Rows)
                        {
                            string scheduleEntry = string.Join(",", row.ItemArray);
                            writer.WriteLine(scheduleEntry);
                        }
                    }

                    MessageBox.Show("Завдання були успішно збережені у файл.", "Інформація",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Помилка при збереженні файлу: " + ex.Message, "Помилка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void збергToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Зберегти у файл Excel";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(saveFileDialog.FileName, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    SheetData sheetData = new SheetData();
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(sheetData);
                    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                    Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                    sheets.Append(sheet);

                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        Cell cell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(dataGridView1.Columns[i].HeaderText) };
                        sheetData.AppendChild(new Row(cell));
                    }

                    foreach (DataGridViewRow dataGridViewRow in dataGridView1.Rows)
                    {
                        Row row = new Row();

                        foreach (DataGridViewCell dataGridViewCell in dataGridViewRow.Cells)
                        {
                            Cell cell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(dataGridViewCell.Value.ToString()) };
                            row.AppendChild(cell);
                        }

                        sheetData.AppendChild(row);
                    }
                }
                MessageBox.Show("Дані успішно збережено у файл Excel.", "Інформація", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void SearchBystate(string searchValue)
        {
            //Функція для пошуку предметів в датагріді за індексом 2 (3 стовпець)
            int columnIndex = 3;


            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[columnIndex].Value != null &&
                    row.Cells[columnIndex].Value.ToString().Equals(searchValue, StringComparison.OrdinalIgnoreCase))
                {
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.LightSeaGreen;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = dataGridView1.DefaultCellStyle.BackColor;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string searchValue = textBox1.Text.Trim();
            SearchBystate(searchValue);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 addSubjectForm = new Form2();
            if (addSubjectForm.ShowDialog() == DialogResult.OK)
            {
                DataRow newRow = scheduleDataTable.NewRow();
                newRow["Завдання"] = addSubjectForm.Task;
                newRow["Пріорітет"] = addSubjectForm.Priority;
                newRow["Термін виконання"] = addSubjectForm.Term;
                newRow["Стан завдання"] = addSubjectForm.State;
                scheduleDataTable.Rows.Add(newRow);
            }
        }
    }
}
