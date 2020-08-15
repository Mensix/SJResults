using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace SJResults
{
    public partial class Form1 : Form
    {
        private static readonly List<Competitor> competitors = JsonConvert.DeserializeObject<List<Competitor>>(File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory + @"\database.json")));
        private static List<Result> results = new List<Result>();
        private static int place = 1;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.DataSource = competitors.Select(x => x.Name).ToList();
            textBox11.Text = place.ToString();
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < competitors.Count; i++)
            {
                if (competitors[i].Name == (string)comboBox1.SelectedItem)
                {
                    textBox12.Focus();
                    textBox6.Text = competitors[i].Code;
                    textBox8.Text = competitors[i].Birth;
                    textBox9.Text = competitors[i].FirstCountry;
                    textBox10.Text = competitors[i].SecondCountry;
                    break;
                }
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            results.Add(new Result
            {
                Code = textBox6.Text,
                Name = (string)comboBox1.Text,
                Birth = textBox8.Text,
                FirstCountry = textBox9.Text,
                SecondCountry = textBox10.Text,
                Place = textBox11.Text,
                FirstLength = textBox12.Text,
                SecondLength = textBox13.Text,
                ThirdLength = textBox14.Text,
                Point = textBox15.Text,
                FirstNote = textBox16.Text,
                SecondNote = textBox17.Text
            });

            textBox6.Text = String.Empty;
            comboBox1.Text = null;
            for (int i = 8; i < 16; i++)
            {
                if (i == 11) continue;
                this.Controls.OfType<TextBox>().Where(x => x.Name == $"textBox{i}").ToList().ForEach(x => x.Text = String.Empty);
            }
            if (textBox11.Text == null || textBox11.Text == "") place = 1;
            else place++;
            textBox11.Text = place.ToString();
            comboBox1.Focus();
        }

        private void Button2_Click(object sender, EventArgs e)
        {

            string[,] excelInit = {
                {
                "Sezon",
                "Center",
                textBox1.Text
                },
                {
                "Kolej",
                "Center",
                null
                },
                {
                "Lp",
                "Center",
                null
                },
                {
                "Data",
                "Center",
                textBox2.Text
                },
                {
                "miejsce",
                "Left",
                textBox3.Text
                },
                {
                "Kraj",
                "Center",
                textBox4.Text
                },
                {
                "Typ",
                "Center",
                textBox5.Text
                },
                {
                "Mirz",
                "Center",
                null
                },
                {
                "mie",
                "Center",
                null
                },
                {
                "kod",
                "Center",
                null
                },
                {
                "nazwisko",
                "Left",
                null
                },
                {
                "Rok ur",
                "Center",
                null
                },
                {
                "kraj",
                "Center",
                null
                },
                {
                "Kraj rz",
                "Center",
                null
                },
                {
                "1s",
                "Center",
                null
                },
                {
                "2s",
                "Center",
                null
                },
                {
                "3s",
                "Center",
                null
                },
                {
                "punwko",
                "Center",
                null
                },
                {
                "Uwagi",
                "Center",
                null
                },
                {
                "Uwagi 2",
                "Center",
                null
                }
            };
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory + @"\Arkusz.xlsx"))))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add("Wyniki");
                excelWorksheet.Cells.Style.Font.Name = "Arial";
                excelWorksheet.Cells.Style.Font.Size = 10;

                for (int i = 0; i < 20; i++)
                {
                    excelWorksheet.Cells[1, i + 1].Value = excelInit[i, 0];
                    excelWorksheet.Cells[2, i + 1].Value = excelInit[i, 2];
                    excelWorksheet.Cells[1, i + 1, 100, i + 1].Style.HorizontalAlignment = excelInit[i, 1] == "Center" ? OfficeOpenXml.Style.ExcelHorizontalAlignment.Center : OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                }

                int startingIndex = 8;
                for (int i = 0; i < results.Count; i++)
                {
                    excelWorksheet.Cells[i + 2, startingIndex].Value = results[i].Place;
                    excelWorksheet.Cells[i + 2, startingIndex + 2].Value = results[i].Code;
                    excelWorksheet.Cells[i + 2, startingIndex + 3].Value = results[i].Name;
                    excelWorksheet.Cells[i + 2, startingIndex + 4].Value = results[i].Birth;
                    excelWorksheet.Cells[i + 2, startingIndex + 5].Value = results[i].FirstCountry;
                    excelWorksheet.Cells[i + 2, startingIndex + 6].Value = results[i].SecondCountry;
                    excelWorksheet.Cells[i + 2, startingIndex + 7].Value = results[i].FirstLength;
                    excelWorksheet.Cells[i + 2, startingIndex + 8].Value = results[i].SecondLength;
                    excelWorksheet.Cells[i + 2, startingIndex + 9].Value = results[i].ThirdLength;
                    excelWorksheet.Cells[i + 2, startingIndex + 10].Value = results[i].Point;
                    excelWorksheet.Cells[i + 2, startingIndex + 11].Value = results[i].FirstNote;
                    excelWorksheet.Cells[i + 2, startingIndex + 12].Value = results[i].FirstNote;
                }

                excelWorksheet.Cells[1, 1, 100, 100].Style.Numberformat.Format = "@";
                excelPackage.Save();
            }

            Application.Exit();
        }
    }
}