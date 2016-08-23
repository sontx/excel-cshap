using ClosedXML.Excel;
using System.Windows.Forms;

namespace ExcelCSharp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            // create new excel workbook
            var workbook = new XLWorkbook();
            // set author for this workbook
            workbook.Author = "www.sontx.in";
            // create a sheet
            var worksheet = workbook.Worksheets.Add("new worksheet");
            // set value for A1 cell in this worksheet
            worksheet.Cell("A1").Value = "This is value in A1 cell";
            // export excel file to disk
            workbook.SaveAs("my_report.xlsx");
            // clean up
            workbook.Dispose();
        }
    }
}
