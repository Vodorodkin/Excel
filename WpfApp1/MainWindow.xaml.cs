using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        PaymentsEntities db = new PaymentsEntities();
        
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
            ex.Visible = true;
            ex.SheetsInNewWorkbook = 1;
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
            ex.DisplayAlerts = false;
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            sheet.Name = "Платежи";
            List<Category> categories = db.Categories.ToList();
            int i = 1;
            int j = 1;
            
            Excel.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 4]];
            range.Merge();
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.Cells.Font.Size = 15;
            range.Cells.Font.Bold = 2;

            sheet.Cells[i, j] = string.Format($"Список платежей");
            i+=2;
            decimal allSum = 0;
            foreach (Category iCat in categories)
            {
                decimal sum = 0;
                sheet.Cells[i, j] = string.Format($"{iCat.Name}");
                 range = sheet.Range[sheet.Cells[i, j], sheet.Cells[i, j + 1]];
                range.Merge();
                range.Cells.Font.Bold = 2;
                i++;
                List<Paying> payings = db.Payings.Where(c => c.Category.ID == iCat.ID).ToList();
                foreach (Paying iPay in payings)
                {
                    sheet.Cells[i, j] = string.Format($"{iPay.Name}");
                    
                    sheet.Cells[i, j+3] = iPay.Sum;
                    
                    sum += iPay.Sum;
                    i++;
                }
                i++;
                sheet.Cells[i, j] = string.Format($"Сумма:");
                
                sheet.Cells[i, j+3] = sum;
                i += 2;
                allSum += sum;
                sum = 0;
            }
            sheet.Cells[i, j] = string.Format("Итого:");
            range = sheet.Range[sheet.Cells[i, j], sheet.Cells[i, j + 1]];
            sheet.Cells[i, j+3] = allSum;
            range.Cells.Font.Bold = 2;




        }
    }
}
