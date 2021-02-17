using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    class VMmain : OnPropertyChangedClass
    {
      public PaymentsEntities db  = new PaymentsEntities();
        private ObservableCollection<Category> _categories;
        public ObservableCollection<Category> categories { get => new ObservableCollection<Category>(db.Categories); set => SetProperty(ref _categories, value); }
        private Category _currentCategory;
        public Category currentCategory { get => _currentCategory; set => SetProperty(ref _currentCategory, value); }
        private ObservableCollection<Paying> _payings;
        public ObservableCollection<Paying> payings { get =>_payings= new ObservableCollection<Paying>(db.Payings); set => SetProperty(ref _payings,value); }
        private RelayCommand _com1;
        public RelayCommand com1 => _com1 ?? (_com1 = new RelayCommand(
            p =>
            {
            string a = "";
            Excel.Application ex = new Excel.Application();
            ex.Visible = true;
            ex.SheetsInNewWorkbook = 2;
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
            ex.DisplayAlerts = false;
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            Excel.Worksheet sheetForChart = (Excel.Worksheet)ex.Worksheets.get_Item(2);
            sheetForChart.Name = "Лист для графика";
            sheetForChart.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            sheet.Name = "Платежи";
            List<Category> categories = new List<Category>();
            if (_currentCategory == null)
            {
                categories = db.Categories.ToList();
            }
            else
            {
                categories.Clear();
                categories.Add(_currentCategory);
            }
            int i = 1;
            int j = 1;

            Excel.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 4]];
            Excel.Range rangeForChart = null;
            range.Merge();
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.Cells.Font.Size = 15;
            range.Cells.Font.Bold = 2;

            sheet.Cells[i, j] = string.Format($"Список платежей");
            i += 1;
            range = sheet.Range[sheet.Cells[i, j], sheet.Cells[i, j + 2]];
            range.Merge();
            sheet.Cells[i, j] = string.Format($"Наименование");

            sheet.Cells[i, j + 3] = string.Format($"Стоимость");
            int x = 1;

            //sheetForChart.Cells[x, 5] = string.Format($"Наименование");
            //sheetForChart.Cells[x, 6] = string.Format($"Стоимость");
            //x++;

            i++;
            decimal allSum = 0;
            foreach (Category iCat in categories)
            {
                decimal sum = 0;
                sheet.Cells[i, j] = string.Format($"{iCat.Name}");
                sheetForChart.Cells[x, 1] = string.Format($"{iCat.Name}");
                range = sheet.Range[sheet.Cells[i, j], sheet.Cells[i, j + 3]];
                range.Merge();
                range.Cells.Font.Bold = 2;
                i++;
                List<Paying> payings = db.Payings.Where(c => c.Category.ID == iCat.ID).ToList();
                foreach (Paying iPay in payings)
                {
                    sheet.Cells[i, j] = string.Format($"{iPay.Name}");
                    range = sheet.Range[sheet.Cells[i, j], sheet.Cells[i, j + 2]];
                    range.Merge();
                    sheet.Cells[i, j + 3] = iPay.Sum;
                    sum += iPay.Sum;
                    i++;
                }
                range = sheet.Range[sheet.Cells[i, j], sheet.Cells[i, j + 3]];
                range.Merge();
                i++;
                range = sheet.Range[sheet.Cells[i, j], sheet.Cells[i, j + 2]];
                range.Merge();
                sheet.Cells[i, j] = string.Format($"Сумма:");
                sheet.Cells[i, j + 3] = sum;
                sheetForChart.Cells[x, 2] = sum;
                x++;
                a += sheet.Cells[i, j + 3].Address + "";
                rangeForChart = sheet.Cells[i, j + 3];
                range = sheet.Range[sheet.Cells[i + 1, j], sheet.Cells[i + 1, j + 3]];
                range.Merge();
                i += 2;
                allSum += sum;
                sum = 0;
            }
            sheet.Cells[i, j] = string.Format("Итого:");
            range = sheet.Range[sheet.Cells[i, j], sheet.Cells[i, j + 2]];
            sheet.Cells[i, j + 3] = allSum;
            range.Cells.Font.Bold = 2;
            range.Merge();
            range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[i, 4]];
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.EntireColumn.AutoFit();
            Excel.SeriesCollection seriesCollection;
            Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(250, 10, 500, 350);
            Excel.Chart chart = myChart.Chart;           
            chart.ChartType = Excel.XlChartType.xlPie;
            chart.HasTitle = true;
            chart.ChartTitle.Text = "Платежи";
            chart.SetSourceData(sheetForChart.Range[$"$A$1:$B${x-1}"]);
            

            }
            ));
        private RelayCommand _com2;
        public RelayCommand com2 => _com2 ?? (_com2 = new RelayCommand(
            p =>
            {
                Excel.Application ex = new Excel.Application();
                ex.Visible = true;
                ex.SheetsInNewWorkbook = 1;
                Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
                ex.DisplayAlerts = false;
                Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
                sheet.Name = "Платежи";
                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 500, 350);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = sheet.get_Range("A3", "B7");
                chartPage.ChartStyle = 1;
                chartPage.HasTitle = true;
                chartPage.ChartTitle.Text = "HeaderText Title";
                chartPage.SetSourceData(chartRange);
                chartPage.ChartType = Excel.XlChartType.xl3DPieExploded;
                chartPage.Elevation = 35;
                //chartPage.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, false, true, true, false, true, false, true, true, Separator: System.Environment.NewLine);
            }
            ));


    }
}
