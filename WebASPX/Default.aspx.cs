using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WebASPX.Model;
using System.IO;
using OfficeOpenXml;

namespace WebASPX
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void OnClickClickMeButton(object sender, EventArgs e)
        {
            MyExcelModel model = new MyExcelModel();
            model.name = "TestName";
            model.surName = "SurName";
            model.rowId = 1;

            string uniqeFileName = Guid.NewGuid().ToString() + ".xlsx";
            string filePath = "D:\\TempFile\\" + uniqeFileName;
            FileUpload1.SaveAs(filePath);

            readExcel(filePath);
        }

        private void readExcel(string filePath)
        {
            //create a list to hold all the values
            List<string> excelData = new List<string>();

            //read the Excel file as byte array
            byte[] bin = File.ReadAllBytes(filePath);

            //or if you use asp.net, get the relative path
            //byte[] bin = File.ReadAllBytes(Server.MapPath("ExcelDemo.xlsx"));

            //create a new Excel package in a memorystream
            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                //loop all worksheets
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    //loop all rows
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        //loop all columns in a row
                        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                        {
                            //add the cell data to the List
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                excelData.Add(worksheet.Cells[i, j].Value.ToString());
                            }
                        }
                    }
                }
            }

            var xx = 55;
        }
    }
}
