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
            
            #endregion
            #region 文件最上方大標題
            Paragraph firstPageTitle = new Paragraph(doc);
            Style firstPageTitleStyle = doc.Styles.Add(StyleType.Paragraph, "FirstPageTitleStyle");
            firstPageTitleStyle.Font.Size = 24;
            //firstPageTitleStyle.Font.Name = "Times New Roman";
            firstPageTitleStyle.Font.NameAscii = "Times New Roman";
            firstPageTitleStyle.Font.NameFarEast = "標楷體";
            firstPageTitle.ParagraphFormat.SpaceAfter = 12;
            firstPageTitle.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            firstPageTitle.ParagraphFormat.Style = firstPageTitleStyle;
            string firstPageTitleContent = "宜蘭縣政府105年度施政目標與重點";
            body.AppendChild(firstPageTitle);
            firstPageTitle.AppendChild(new Run(doc, firstPageTitleContent));

            #region 第一頁
            Style firstPageContentStyle = doc.Styles.Add(StyleType.Paragraph, "FirstPageContentStyle");
            firstPageContentStyle.Font.Size = 14;
            firstPageContentStyle.Font.NameAscii = "Times New Roman";
            firstPageContentStyle.Font.NameFarEast = "標楷體";
            #region 第一頁內容設定
            List firstPageContentList = doc.Lists.Add(ListTemplate.NumberDefault);
            firstPageContentList.ListLevels[0].NumberStyle = NumberStyle.None;
            firstPageContentList.ListLevels[0].TrailingCharacter = ListTrailingCharacter.Nothing;
            firstPageContentList.ListLevels[0].NumberFormat = "\u0000";
            firstPageContentList.ListLevels[0].NumberPosition = 10;
            firstPageContentList.ListLevels[1].NumberStyle = NumberStyle.None;
            firstPageContentList.ListLevels[1].TrailingCharacter = ListTrailingCharacter.Nothing;
            firstPageContentList.ListLevels[1].NumberFormat = "\u0000";
            firstPageContentList.ListLevels[1].NumberPosition = 15;
            firstPageContentList.ListLevels[2].NumberStyle = NumberStyle.None; //(一)
            firstPageContentList.ListLevels[2].TrailingCharacter = ListTrailingCharacter.Nothing;
            firstPageContentList.ListLevels[2].NumberFormat = "\u0000";
            firstPageContentList.ListLevels[2].NumberPosition = 25;
            firstPageContentList.ListLevels[3].NumberStyle = NumberStyle.None;
            firstPageContentList.ListLevels[3].TrailingCharacter = ListTrailingCharacter.Nothing;
            firstPageContentList.ListLevels[3].NumberFormat = "\u0000";
            firstPageContentList.ListLevels[3].NumberPosition = 40;
            #endregion
            string firstPageContentStr = @" 專責宜蘭縣工商與觀光旅遊發展推動，並貫徹公平交易法之執行工作；配合中央各項經濟政策與措施，積極營造良好的投資環境，訂頒獎勵辦法鼓勵設廠投資；另一方面，本府善用好山好水的觀光資源，創造了宜蘭觀光顯著的能見度，引領台灣觀光發展的新方向，並持續宜蘭縣全方位的觀光建設推動工作；然而隨著社經環境的快速變遷，消費型態日趨多元化與複雜，在推動工商與觀光旅遊的同時，消費者的權益亦趨重視與加強。
配合縣政施政重點藍圖、中長程施政計畫及核定預算額度，針對當前社經狀況及未來發展需要，編定101年度施政計畫，其目標與重點如次：
一、工商類：
    (一)持續加強簡政便民措施及更新資訊系統，以強化商業行政業務，改善及強化工商登記業務功能。
    (二)活絡商業發展，推動各鄉鎮成立特色商圈工作，協助各鄉鎮市辦理「地方小鎮振興藍圖規劃計畫」、「小鎮商街再造計畫」輔導管理工作。
    (三)健全產業發展，推動檢討閒置工業區及不合時宜之區位變更調整為工商綜合區之設置，以改善投資環境，促進地方產業升級與發展。
    (四)協助科學園區設置之各項業務推動及招商工作。
    (五)貫徹公平交易法之執行。
    (六)加強視聽歌唱等8種行業、電子遊戲場業、資訊休閒等特種行業之輔導管理，以維社區安寧及公共安全。
    (七)輔導傳統產業走向文化化、觀光化，並與觀光旅遊業結合，使工廠具體呈顯地方特色及人文采風，轉型為生產、銷售與觀光體驗一貫化，以加值傳統產業。
    (八)推動「綠能及新興科技（低污染）產業計畫」，並對於太陽能、風力發電、生質能源研發產業，引進潔淨、低污染生產技術，以及高產值的產業，例如通訊、軟體設計產業予以獎勵措施及行政協助，以鼓勵其於本縣進駐。
    (九)推動「產業微型園區發展計畫」，於本縣設置產業園區，俾供本縣具有特色之中小企業，例如低污染且具有特有的歷史、文化、創意，特性，並運用本地素材、自然資源、傳統技藝、勞動力等從事生產或提供服務的產業，予引導形成產業聚落。
    (十)推動「本府幸福創業貸款計畫」，協助縣民創業或擴大經營績效予以貸款，並予以創業計畫輔導等協助。
    (十一)推動「青年創意產業」，輔導青年返鄉從事有機、精緻、創意農業，協助發展「品牌農業」，並鼓勵在地青年發展新興的文化創意產業，例如具文化創意特色的商品或手工業產品。
    (十二)積極辦理招商業務，並針對本府未來想發展之重點產業（如綠能、影視、新興科技及低污染）主動拜訪，並爭取其於本縣進駐，並透過相關行政作為，加速縮短其辦理證照時間。
二、觀光行銷類：
    (一)持續辦理觀光旅遊服務網路，提供旅客旅遊資訊查詢，並配合大型活動設置旅遊服務單一窗口。
    (二)辦理觀光宣傳、推廣及參加國內外旅展及邀請香港、日本、大陸等國家至本縣踩線（熟悉之旅），加強與觀光業界之聯繫與合作，俾強化觀光政策之執行。
    (三)針對自由行旅客提供完整觀光護照，以鼓勵國外遊客至本縣旅遊。
    (四)鼓勵業者開發新的套裝旅遊行程，推動觀光活動，持續加強觀光遊樂設施督導考核及辦理民宿旅館評鑑工作。
    (五)強化本縣旅賓館業及民宿經營管理及輔導，提升住宿服務品質，並辦理觀光產業從業人員之培育、訓練等事項。
    (六)辦理自行車甲租乙還、自行車挑戰賽及逍遙遊活動套裝旅遊。
    (七)輔導縣內飯店業者辦理暨爭取國際會議、企業獎勵旅遊、會議旅遊等。
三、遊憩規劃類：
    (一)創造良好觀光投資經營環境，以吸引民間投資。
    (二)整體規劃全縣觀光遊憩地區發展計畫，逐步建設各風景據點，以提供優質旅遊環境。
    (三)促進民間參與公共建設案業務之後續履約管理，以提升本縣觀光服務品質，加速社會經濟發展。
四、遊憩管理類：
    (一)安全、整潔、便利，為執行轄屬風景遊憩區旅遊環境首要。
    (二)辦理轄屬各風景遊憩區公共及遊憩設施維護。
    (三)辦理轄屬各風景遊憩區安全管理。
    (四)辦理旅遊資訊化服務。
    (五)辦理遊客服務解說等經營管理業務。
    (六)自行車步道環境清潔管理維護。
    (七)策略性帶動及推廣周邊觀光旅遊產業。
    (八)辦理轄屬遊憩景點賣店委外業務。";
            string[] Content = firstPageContentStr.Replace('\r',' ').Split('\n');
          
            int levelnumber = 0;
            foreach(string c in Content)
            {
                // 找c的前面有幾個空白
                int count = c.TakeWhile(Char.IsWhiteSpace).Count();
                if (count == 0)
                {
                    levelnumber = 0;
                }else if (count == 2)
                {
                    levelnumber = 1;
                }
                Paragraph paragraph = new Paragraph(doc);
                paragraph.ListFormat.List = firstPageContentList;
                //paragraph.ListFormat.ListLevelNumber = levelnumber;
                paragraph.ParagraphFormat.Style = firstPageContentStyle;
                paragraph.ParagraphFormat.LeftIndent = 4;
                paragraph.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                paragraph.AppendChild(new Run(doc, c));
                doc.FirstSection.Body.AppendChild(paragraph);
                //paragraph.ListFormat.RemoveNumbers();
                levelnumber++;
            }
            
            //firstPageConten.AppendChild(new Run(doc, firstPageContentStr));
            #endregion
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
            secondPageTitleStyle.Font.NameAscii = "Times New Roman";
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
                p.ParagraphFormat.Style.Font.NameAscii = "Times New Roman";
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
                firstList.ListLevels[0].Font.NameAscii = "Times New Roman";
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
                secondList.ListLevels[0].Font.NameAscii = "Times New Roman";
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
                    // cellParagraphStyle START
                    Style cellParagraphStyle = doc.Styles.Add(StyleType.Paragraph, "CellParagraphStyle");
                    cellParagraphStyle.Font.NameAscii = "Times New Roman";
                    cellParagraphStyle.Font.NameFarEast = "標楷體";
                    cellParagraphStyle.Font.Bold = false;
                    cellParagraphStyle.Font.Size = 12;
                    //cellParagraphStyle END
                    cellParagraph.ParagraphFormat.Style = cellParagraphStyle;
                    cell.AppendChild(cellParagraph);
                    string text;


                    if (j == 0) // 第一欄
                    {
                        // 業務別
                        firstList.ListLevels[0].StartAt = number;
                        cellParagraph.ListFormat.List = firstList;
                        number++;
                    }
                    else if (j == 1) // 第二欄
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
            row.Cells[0].CellFormat.Width = 80;
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