using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Web.Mvc;
using Aspose.Words;
using ILHG_TEST.Models;
using Aspose.Words.Tables;
using System.Drawing;
using Aspose.Words.Lists;

namespace ILHG_TEST.Controllers
{
    public class SampleController : _Controller
    {
        // GET: Sample
        public ActionResult Index()
        {

            return View();
        }
        [HttpPost]
        public void Print()
        {
            //string tempPath = Server.MapPath("../FileTemplete/ProjectMain.docx");
            
            string tempPath = Path.Combine(Server.MapPath("../FileTemplete/"), "ProjectMain.docx");

            DocConvertController docconverter = new DocConvertController();
            
            // 產一份新的 doc
            Document doc = docconverter.GenerateDocument(tempPath);
            
            DocConvertConfigModel.WordToPdf model = new DocConvertConfigModel.WordToPdf()
            {
                TMPL_PATH = tempPath,
                TARGET_PATH = "D:/testWaterMarkFile.doc",
                TITLE = "test",
                IS_DISPLAY_WATERMARK = true,
                WATERMARK_TEXT="GSSSSSSS",
                // 要放入範本檔中的資料
                DATA = new ProjectMain()
                {
                   Project_Id= 1,
                   Project_Name = "測試測試測試",
                   Project_Date = Convert.ToDateTime("2018-03-07"),
                   Promoter_Name = "TEST_promoter",
                   Plan_Apply = "APPLY_TEST"
                }

            };
            //using (MemoryStream stream = new MemoryStream())
            //{
            //    // 存取檔案的路徑
            //    //doc.Save("D:/testFile1.odt");

            //    doc.Save(stream, SaveFormat.Odt);
            //    OpenOdt(stream.ToArray(), "test.odt");
            //   using(DocConvertController docc = new DocConvertController())
            //    {
            //        docc.WordToOdt(model);
            //        //docc.WordToOdtBuffer(model);
            //    }
            //}
            //Document doc = new Document(Server.MapPath("../FileTemplete/123.docx"));
            //doc.Save("C:/Users/USER/Downloads/test.pdf", SaveFormat.Pdf);
           
        }

        /// <summary>
        /// 產生內含 TABLE 的 DOC 檔 ( TABLE 沒有內框線)
        /// </summary>
        [HttpPost]
        public void PrintDocWithTable()
        {
            Document doc = new Document();
            // 新增一個 4*4 且每一個儲存格內容皆為 123 的 TABLE 
            Table table = CreateTable(doc, 4, 4, "123");
            
            // Align the table to the center of the page.
            table.Alignment = TableAlignment.Center;

            // Clear any existing borders from the table.
            table.ClearBorders();

            // Set a black border around the table but not inside.
            // TABLE 樣式設定為有黑色外框線但無內框線 
            table.SetBorder(BorderType.Left, LineStyle.Single, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Right, LineStyle.Single, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Black, true);

            // 網底為白色
            // Fill the cells with a white solid color.
            table.SetShading(TextureIndex.TextureSolid, Color.White, Color.Empty);

            // 將處理好的 TABLE 放入 DOC 中
            doc.FirstSection.Body.AppendChild(table);
            HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            doc.FirstSection.HeadersFooters.Add(footer);
            footer.AppendParagraph("i am footer");

            // 產出DOC
            doc.Save("D:/Table.SetOutlineBorders Out.doc");
        }

        /// <summary>
        /// Creates a new table in the document with the given dimensions and text in each cell.
        /// </summary>
        private Table CreateTable(Document doc, int rowCount, int cellCount, string cellText)
        {
            Table table = new Table(doc);

            // Create the specified number of rows.
            for (int rowId = 1; rowId <= rowCount; rowId++)
            {
                Row row = new Row(doc);
                table.AppendChild(row);

                // Create the specified number of cells for each row.
                for (int cellId = 1; cellId <= cellCount; cellId++)
                {
                    Cell cell = new Cell(doc);
                    row.AppendChild(cell);
                    // Add a blank paragraph to the cell.
                    cell.AppendChild(new Paragraph(doc));

                    // Add the text.
                    cell.FirstParagraph.AppendChild(new Run(doc, cellText));
                }
            }

            return table;
        }
        /// <summary>
        /// 產出表格(合併儲存格)
        /// </summary>
        [HttpPost]
        public void PrintMergeColumnTable()
        {
            Document doc = new Document();
            Table table = new Table(doc);
            int number = 1;
            
            for (int i = 0; i < 3; i++) // 資料比數 (foreach
            {
                Row row = new Row(doc);
                List list = doc.Lists.Add(ListTemplate.NumberDefault);
                // 標號都一樣 壹 QQ
                list.ListLevels[0].Font.Color = Color.Red;
                list.ListLevels[0].Font.Size = 24;
                list.ListLevels[0].NumberStyle = NumberStyle.SimpChinNum2;
                list.ListLevels[0].StartAt = 1;
                //list.ListLevels[0].NumberFormat = "\x0000";
                //level1.NumberFormat = "\x0000";
                table.AppendChild(row);
                for(int j = 0; j < 4; j++) // 欄位數量 (foreach
                {
                    // 新增儲存格
                    Cell cell = new Cell(doc);
                    row.AppendChild(cell);
                    Paragraph p = new Paragraph(doc);
                    cell.AppendChild(p);
                    // 列合併儲存格(垂直合併)
                    //cell.CellFormat.VerticalMerge = CellMerge.First;
                    //cell.CellFormat.VerticalMerge = CellMerge.Previous;
                    // 寫入儲存格內容
                    if (j == 0)
                    {
                        //p.ListFormat.ListIndent();
                        p.ListFormat.List = list;
                        
                    }
                    string text = "123";
                    
                   
                    cell.FirstParagraph.AppendChild(new Run(doc, text));

                }
                //EndRow
                number++;
            }
            // EndTable
            SetTableBorder(table);
            doc.FirstSection.Body.AppendChild(table);

            doc.Save("D:/mergeColumnTable.doc");
        }
        /// <summary>
        /// 設定表格邊線(只有外框線)
        /// </summary>
        /// <param name="table"></param>
        private void SetTableBorder(Table table)
        {
            table.ClearBorders();
            //table.AutoFit(AutoFitBehavior.AutoFitToContents);
            table.SetBorder(BorderType.Left, LineStyle.Single, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Right, LineStyle.Single, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Black, true);
        }
    }
}