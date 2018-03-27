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
        #region first test doc
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
                WATERMARK_TEXT = "GSSSSSSS",
                // 要放入範本檔中的資料
                DATA = new ProjectMain()
                {
                    Project_Id = 1,
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
        #endregion

        #region table doc
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

        #endregion

        #region 配合格式設定
        /// <summary>
        /// 產出表格(合併儲存格)
        /// </summary>
        [HttpPost]
        public void PrintMergeColumnTable()
        {
            #region 初始設定
            License license = new License();
            license.SetLicense("Aspose.Total.655.lic");
            Document doc = new Document();
            doc.RemoveAllChildren();
            #endregion
            #region 報表第一頁設定
            Section section = new Section(doc);
            doc.AppendChild(section);
            // 設定邊界
            section.PageSetup.TopMargin = 42.5;
            section.PageSetup.BottomMargin = 42.5;
            section.PageSetup.LeftMargin = 42.5;
            section.PageSetup.RightMargin = 42.5;
            section.PageSetup.SectionStart = SectionStart.NewPage;
            // 頁面尺寸
            section.PageSetup.PaperSize = PaperSize.A4;
            // 頁碼起始設定
            section.PageSetup.PageStartingNumber = 1;
            // 頁碼格式
            section.PageSetup.PageNumberStyle = NumberStyle.Arabic;
            Body body = new Body(doc);
            section.AppendChild(body);
            #region 頁碼
            // 頁尾
            HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            doc.FirstSection.HeadersFooters.Add(footer);
            Paragraph para = new Paragraph(doc);
            // 放入頁碼
            para.InsertField(Aspose.Words.Fields.FieldType.FieldPage, false, null, true);
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            Style footerStyle = doc.Styles.Add(StyleType.Paragraph, "footerStyle");
            // 設定頁碼大小字型
            footerStyle.Font.Size = 10;
            footerStyle.Font.Name = "Times New Roman";
            para.ParagraphFormat.Style = footerStyle;
            footer.AppendChild(para);
            #endregion
            #endregion
            #region 文件最上方大標題
            Paragraph firstPageTitle = new Paragraph(doc);
            Style firstPageTitleStyle = doc.Styles.Add(StyleType.Paragraph, "FirstPageTitleStyle");
            firstPageTitleStyle.Font.Size = 24;
            firstPageTitleStyle.Font.Name = "Times New Roman";
            firstPageTitleStyle.Font.NameFarEast = "標楷體";
            firstPageTitle.ParagraphFormat.SpaceAfter = 12;
            firstPageTitle.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            firstPageTitle.ParagraphFormat.Style = firstPageTitleStyle;
            string firstPageTitleContent = "宜蘭縣政府105年度施政目標與重點";
            body.AppendChild(firstPageTitle);
            firstPageTitle.AppendChild(new Run(doc, firstPageTitleContent));

            #region
            Paragraph firstPageConten = new Paragraph(doc);
            Style firstPageContentStyle = doc.Styles.Add(StyleType.Paragraph, "FirstPageContentStyle");
            firstPageContentStyle.Font.Size = 14;
            firstPageContentStyle.Font.Name = "Times New Roman";
            firstPageContentStyle.Font.NameFarEast = "標楷體";
            firstPageConten.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
            firstPageConten.ParagraphFormat.Style = firstPageContentStyle;
            string firstPageContentStr = "參、工商旅遊處\r\n一、項目\r\n（一）項目\r\n１、項目";
            body.AppendChild(firstPageConten);
            firstPageConten.AppendChild(new Run(doc, firstPageContentStr));
            #endregion

            #endregion
            #region 報表第二頁後設定
            Section secondSection = new Section(doc);
            doc.AppendChild(secondSection);
            secondSection.PageSetup.SectionStart = SectionStart.NewPage;
            secondSection.PageSetup.TopMargin = 42.5;
            secondSection.PageSetup.BottomMargin = 42.5;
            secondSection.PageSetup.LeftMargin = 42.5;
            secondSection.PageSetup.RightMargin = 42.5;
            secondSection.PageSetup.PaperSize = PaperSize.A4;
            secondSection.PageSetup.RestartPageNumbering = false;
            Body secondBody = new Body(doc);
            secondSection.AppendChild(secondBody);
            #endregion

            #region 表格標題設定
            Paragraph docTitle = new Paragraph(doc);
            Style secondPageTitleStyle = doc.Styles.Add(StyleType.Paragraph, "SecondPageTitleStyle");
            secondPageTitleStyle.ParagraphFormat.SpaceAfter = 12;
            secondPageTitleStyle.Font.Size = 16;
            secondPageTitleStyle.Font.Name = "Times New Roman";
            secondPageTitleStyle.Font.NameFarEast = "標楷體";
            secondPageTitleStyle.Font.Bold = true;
            docTitle.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            docTitle.ParagraphFormat.Style = secondPageTitleStyle;
            string firsttitle = "宜蘭縣政府105年度重要施政計畫";
            secondBody.AppendChild(docTitle);
            docTitle.AppendChild(new Run(doc, firsttitle));
            #endregion
            #region 生產表格
            Table table = new Table(doc);
            doc.LastSection.Body.AppendChild(table);
            int number = 1;
            string[] tabletitle = { "業務別", "重要施政計畫項目", "施政內容", "預算金額\r\n(單位：千元)", "備註" };
            Row titleRow = new Row(doc);
            table.AppendChild(titleRow);
            foreach (var title in tabletitle)
            {
                Cell cell = new Cell(doc);
                cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                titleRow.AppendChild(cell);
                Paragraph p = new Paragraph(doc);
                p.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                p.ParagraphFormat.Style.Font.Name = "Times New Roman";
                p.ParagraphFormat.Style.Font.NameFarEast = "標楷體";
                p.ParagraphFormat.Style.Font.Bold = false;
                p.ParagraphFormat.Style.Font.Size = 12;
                p.ParagraphFormat.KeepWithNext = false;
                cell.AppendChild(p);
                cell.FirstParagraph.AppendChild(new Run(doc, title));
            }
            #region 表格換頁設計
            titleRow.RowFormat.HeadingFormat = true;
            titleRow.RowFormat.AllowBreakAcrossPages = false;
            setRowWidth(titleRow);
            #endregion



            for (int i = 1; i < 70; i++) // 資料比數 (foreach
            {
                Row row = new Row(doc);
                #region firstList
                List firstList = doc.Lists.Add(ListTemplate.NumberDefault);
                // 
                firstList.ListLevels[0].Font.Color = Color.Black;
                firstList.ListLevels[0].Alignment = ListLevelAlignment.Center;
                firstList.ListLevels[0].Font.Size = 12;
                firstList.ListLevels[0].Font.Name = "Times New Roman";
                firstList.ListLevels[0].Font.NameFarEast = "標楷體";
                firstList.ListLevels[0].NumberStyle = NumberStyle.TradChinNum2;
                firstList.ListLevels[0].TrailingCharacter = ListTrailingCharacter.Nothing;
                firstList.ListLevels[0].TextPosition = 0;
                #endregion

                #region secondList
                List secondList = doc.Lists.Add(ListTemplate.NumberDefault);
                // 
                secondList.ListLevels[0].Font.Color = Color.Black;
                secondList.ListLevels[0].Alignment = ListLevelAlignment.Center;
                secondList.ListLevels[0].Font.Size = 12;
                secondList.ListLevels[0].Font.Name = "Times New Roman";
                secondList.ListLevels[0].Font.NameFarEast = "標楷體";
                secondList.ListLevels[0].NumberStyle = NumberStyle.TradChinNum1;
                secondList.ListLevels[0].TrailingCharacter = ListTrailingCharacter.Nothing;
                secondList.ListLevels[0].TextPosition = 0;
                #endregion
                
                table.AppendChild(row);
                #region cell
                for (int j = 0; j < 5; j++) // 欄位數量 (foreach
                {
                    // 新增儲存格
                    Cell cell = new Cell(doc);
                    row.AppendChild(cell);
                    Paragraph cellParagraph = new Paragraph(doc);
                    cellParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                    cellParagraph.ParagraphFormat.Style.Font.Name = "Times New Roman";
                    cellParagraph.ParagraphFormat.Style.Font.NameFarEast = "標楷體";
                    cellParagraph.ParagraphFormat.Style.Font.Bold = false;
                    cellParagraph.ParagraphFormat.Style.Font.Size = 12;
                    cell.AppendChild(cellParagraph);
                    string text;


                    if (j == 0)
                    {
                        // 業務別
                        firstList.ListLevels[0].StartAt = number;
                        cellParagraph.ListFormat.List = firstList;
                        number++;
                    }
                    else if (j == 1)
                    {
                        // 重要施政計畫項目
                        #region rowspan
                        // 列合併儲存格(垂直合併)
                        //cell.CellFormat.VerticalMerge = CellMerge.First;
                        //cell.CellFormat.VerticalMerge = CellMerge.Previous;
                        #endregion
                        secondList.ListLevels[0].StartAt = i;
                        cellParagraph.ListFormat.List = secondList;
                    }
                    text = "123";

                    cell.FirstParagraph.AppendChild(new Run(doc, text));
                    
                }
                row.RowFormat.AllowBreakAcrossPages = true;
                row.RowFormat.HeightRule = HeightRule.AtLeast;
                setRowWidth(row);

                //EndRow
                #endregion
                

            }
            // EndTable

            SetTableStyle(table);
            #endregion
            doc.Save("D:/mergeColumnTable.odt",SaveFormat.Odt);
        }
        /// <summary>
        /// 設定表格邊線(只有外框線)
        /// </summary>
        /// <param name="table"></param>
        private void SetTableStyle(Table table)
        {
            table.ClearBorders();
            table.Alignment = TableAlignment.Center;
            table.SetBorder(BorderType.Left, LineStyle.Single, 1, Color.Black, true);
            table.SetBorder(BorderType.Right, LineStyle.Single, 1, Color.Black, true);
            table.SetBorder(BorderType.Top, LineStyle.Single, 1, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 1, Color.Black, true);
        }
        /// <summary>
        /// 設定欄位寬度
        /// </summary>
        /// <param name="row"></param>
        private void setRowWidth(Row row)
        {
            // For Odt
            row.Cells[0].CellFormat.Width = 90;
            row.Cells[1].CellFormat.Width = 110;
            row.Cells[2].CellFormat.Width = 180;
            row.Cells[3].CellFormat.Width = 90;
            row.Cells[4].CellFormat.Width = 50;
            // For Word
            //row.Cells[0].CellFormat.Width = 150;
            //row.Cells[1].CellFormat.Width = 220;
            //row.Cells[2].CellFormat.Width = 380;
            //row.Cells[3].CellFormat.Width = 150;
            //row.Cells[4].CellFormat.Width = 100;
        }
        
        #endregion
    }
}