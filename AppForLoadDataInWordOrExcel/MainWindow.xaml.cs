using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.Office.Interop.Word;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
namespace AppForLoadDataInWordOrExcel
{
  
    public partial class MainWindow : System.Windows.Window
    {
        List<Info> infoList = new List<Info>();
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            LoadDataGrid();
        }
        public void LoadDataGrid()
        {
            dataGrid.ItemsSource = null;  // Clear the existing ItemsSource
            dataGrid.ItemsSource = infoList;  // Set the new ItemsSource
        }
        private void ExportToWord_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
                winword.Visible = true;
                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add();

                // Add header
                Microsoft.Office.Interop.Word.Paragraph headerPar = document.Content.Paragraphs.Add();
                headerPar.Range.Text = "Накладная";
                headerPar.Range.Font.Name = "Times new roman";
                headerPar.Range.Font.Size = 16;
                headerPar.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                headerPar.Range.InsertParagraphAfter();

                // Add invoice information
                Microsoft.Office.Interop.Word.Paragraph invoicePar = document.Content.Paragraphs.Add();
                invoicePar.Range.Text = "Строк: " + (infoList.Count + 1);
                invoicePar.Range.Font.Name = "Times new roman";
                invoicePar.Range.Font.Size = 14;
                invoicePar.Range.InsertParagraphAfter();
                Microsoft.Office.Interop.Word.Paragraph supplier = document.Content.Paragraphs.Add();
                supplier.Range.Text = "Поставщик: " + (supplierTextBox.Text );
                supplier.Range.Font.Name = "Times new roman";
                supplier.Range.Font.Size = 14;
                supplier.Range.InsertParagraphAfter();
                Microsoft.Office.Interop.Word.Paragraph client = document.Content.Paragraphs.Add();
                client.Range.Text = "Клиент: " + (clientTextBox.Text);
                client.Range.Font.Name = "Times new roman";
                client.Range.Font.Size = 14;
                client.Range.InsertParagraphAfter();
                Microsoft.Office.Interop.Word.Paragraph number = document.Content.Paragraphs.Add();
                number.Range.Text = "Номер заказа: " + (numberTextBox.Text);
                number.Range.Font.Name = "Times new roman";
                number.Range.Font.Size = 14;
                number.Range.InsertParagraphAfter();
                Microsoft.Office.Interop.Word.Paragraph paragraph = document.Content.Paragraphs.Add();
                paragraph.Range.Text = "Итого: " + (Convert.ToString(AmountSum(infoList)));
                paragraph.Range.Font.Name = "Times new roman";
                paragraph.Range.Font.Size = 14;
                paragraph.Range.InsertParagraphAfter();
                // Add table
                Microsoft.Office.Interop.Word.Table myTable = document.Tables.Add(invoicePar.Range, infoList.Count + 1, 5);
                myTable.Borders.Enable = 1;

                // Add table headers
                string[] headers = { "Номер", "Продукт", "Количество", "Цена", "Кол*Цена" };
                for (int i = 1; i <= headers.Length; i++)
                {
                    myTable.Cell(1, i).Range.Text = headers[i - 1];
                }

                // Add data to the table
                for (int i = 0; i < infoList.Count; i++)
                {
                    var info = infoList[i];
                    myTable.Cell(i + 2, 1).Range.Text = info.Number.ToString();
                    myTable.Cell(i + 2, 2).Range.Text = info.Product;
                    myTable.Cell(i + 2, 3).Range.Text = info.Quantity.ToString();
                    myTable.Cell(i + 2, 4).Range.Text = info.Price.ToString();
                    myTable.Cell(i + 2, 5).Range.Text = info.Amount.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error exporting to Word: " + ex.Message);
            }

        }
        public decimal AmountSum(List<Info> list)
        {
            decimal result = 0;
            foreach(Info info in list) {
                result += info.Amount;
            }
            return result;
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = true;
                Workbook workbook = excelApp.Workbooks.Add();
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                // Add headers
                string[] headers = { "Number", "Product", "Quantity", "Price", "Amount" };
                for (int i = 1; i <= headers.Length; i++)
                {
                    worksheet.Cells[1, i] = headers[i - 1];
                }

                // Add data to the worksheet
                for (int i = 0; i < infoList.Count; i++)
                {
                    var info = infoList[i];
                    worksheet.Cells[i + 2, 1] = info.Number;
                    worksheet.Cells[i + 2, 2] = info.Product;
                    worksheet.Cells[i + 2, 3] = info.Quantity;
                    worksheet.Cells[i + 2, 4] = info.Price;
                    worksheet.Cells[i + 2, 5] = info.Amount;
                }

                // Add additional information
                worksheet.Cells[infoList.Count + 4, 1] = "Поставщик:";
                worksheet.Cells[infoList.Count + 4, 2] = supplierTextBox.Text;
                worksheet.Cells[infoList.Count + 5, 1] = "Клиент:";
                worksheet.Cells[infoList.Count + 5, 2] = clientTextBox.Text;
                worksheet.Cells[infoList.Count + 6, 1] = "Номер заказа:";
                worksheet.Cells[infoList.Count + 6, 2] = numberTextBox.Text;
                worksheet.Cells[infoList.Count + 7, 1] = "Итого";
                worksheet.Cells[infoList.Count + 7, 2] = Convert.ToString(AmountSum(infoList));


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error exporting to Excel: " + ex.Message);
            }
        }

        private void AddNewColumn_Click(object sender, RoutedEventArgs e)
        {
            Info info = new Info();
            info.Number = infoList.Count+1;
            info.Product = productTextBox.Text;
            info.Quantity = Convert.ToInt32(lotTextBox.Text);
            info.Price = Convert.ToDecimal(priceTextBox.Text);
            info.Amount = info.Price * info.Quantity;
            infoList.Add(info);
            LoadDataGrid();

        }

        private void DeleteNewColumn_Click(object sender, RoutedEventArgs e)
        {
            infoList.RemoveAt(infoList.Count-1);
        }
    }
    public class Info
    { 
        public int Number { get; set; }
        public string Product { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public decimal Amount { get; set; }
        


    }

}
