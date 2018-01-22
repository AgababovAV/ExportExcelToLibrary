using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

//Импорт из списка в библиотеку напрямую в виде excel

namespace exportExcel
{
    class Program
    {

        public static void Main(string[] args)
        {
            using (SPSite siteCollection = new SPSite("portal"))
            {

                SPWebCollection sites = siteCollection.AllWebs;
                SPWeb site = siteCollection.OpenWeb();
                SPList list = site.GetList(site.Url + "/Lists/AnswerForStore/AllItems.aspx");
                SPView wpView = list.DefaultView;
                wpView.RowLimit = 2000000;
                SPQuery query = new SPQuery(wpView);
               
                SPListItemCollection exportListItems = list.GetItems(query);
                string documentLibraryName = site.Url + "/Documents/RetailReport";
                //SPListItemCollection exportListItems = site.GetList(siteCollection.Url + "/Lists/AnswerForStore/AllItems.aspx").Items;
               
                
                DataTable dt = new DataTable();
                if (exportListItems.Count > 0)
                {
                    dt = exportListItems.GetDataTable();
                    
                }
                
                ExportDataTableToExcel(dt, "Отзывы о магазине " + DateTime.Now.ToString("MM.yyyy") + ".xls", site, documentLibraryName);
               
            }


        }

        public static void ExportDataTableToExcel(DataTable sourceTable, string fileName, SPWeb site,string documentLibraryName)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            MemoryStream memoryStream = new MemoryStream();
            var sheet = workbook.CreateSheet("Отчет");
            var headerRow = sheet.CreateRow(0);
            DataView dvListViewData = sourceTable.DefaultView;

           
            foreach (DataColumn column in sourceTable.Columns)
            {
                
                if (column.Ordinal < sourceTable.Columns.Count - 3) //убираем 3 последних стобца id,create,modified
                {
                    headerRow.CreateCell(column.Ordinal).SetCellValue(column.Caption);
                    
                }
            }
                

            // handling value.
            int rowIndex = 1;

            foreach (DataRow row in sourceTable.Rows)
            {
                var dataRow = sheet.CreateRow(rowIndex);
                
                foreach (DataColumn column in sourceTable.Columns)
                {

                    if (column.Ordinal < sourceTable.Columns.Count - 3)
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue(row[column.Ordinal].ToString());
                    }
                   
                }

                rowIndex++;
            }

            workbook.Write(memoryStream);
            memoryStream.Flush();

            try
            {
                SPFolder myLibrary = site.GetFolder(documentLibraryName);
                site.AllowUnsafeUpdates = true;
                SPFile spfile = myLibrary.Files.Add(fileName, memoryStream, true);
                // Commit 
                myLibrary.Update();
            }
            catch (Exception ex)
            {
                string urlFolder = site.Url + "/Documents/Forms/AllItems.aspx";
                SPFolder myLibrary = site.GetList(urlFolder).RootFolder;
                site.AllowUnsafeUpdates = true;
                SPFile spfile = myLibrary.Files.Add(fileName, memoryStream, true);
                // Commit 
                myLibrary.Update();
            }

        }
    }
}

