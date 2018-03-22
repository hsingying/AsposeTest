using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ILHG_TEST.Models;
using Aspose.Words;
using System.IO;
using Aspose.Words.Drawing;
using System.Reflection;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

namespace ILHG_TEST.Controllers
{
    /// <summary>
    /// 文件轉換器
    /// </summary>
    /// <remarks>
    ///     Created by Steven Tsai at 2017/08/07
    /// </remarks>
    [Authorize]
    [SessionState(System.Web.SessionState.SessionStateBehavior.ReadOnly)]
    public class DocConvertController : Controller
    {
        #region 相關 Word 應用（一般固定格式產生用）

        /// <summary>
        /// 將 Word 轉成 Odt 檔案輸出
        /// </summary>
        /// <param name="cfgModel">將 Word 轉成 Odt 檔案輸出的設定</param>
        public void WordToOdt(DocConvertConfigModel.WordToPdf cfgModel)
        {
            #region 防呆

            // 需指定產出 Odt 的路徑（樣版）
            if (string.IsNullOrEmpty(cfgModel.TARGET_PATH))
            {
                throw new ArgumentNullException("cfgModel.TARGET_PATH");
            }

            #endregion

            // 初始化 Word 檔案
            Document doc = __InitialWord(cfgModel);

            // 產出 Odt 檔案
            doc.Save(cfgModel.TARGET_PATH);
        }

        /// <summary>
        /// 將 Word 轉成 Odt 檔案串流輸出
        /// </summary>
        /// <param name="cfgModel">將 Word 轉成 Odt 檔案輸出的設定</param>
        /// <returns>轉換完成的 Odt 檔案串流輸出</returns>
        public byte[] WordToOdtBuffer(DocConvertConfigModel.WordToPdf cfgModel)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                // 初始化 Word 檔案
                Document doc = __InitialWord(cfgModel);

                // 產出 Pdf 檔案串流
                doc.Save(ms, SaveFormat.Odt);

                // 轉換成 byte[] 回傳
                return ms.ToArray();
            }
        }

        /// <summary>
        /// 初始化 Word 檔案
        /// </summary>
        /// <param name="cfgModel">將 Word 轉成 Odt 檔案輸出的設定</param>
        /// <remarks>
        ///     
        ///     1. 使用 Aspose 套件讀取 Word 檔案
        ///     2. 初始動作包括：
        ///         2.1. 讀取文件（樣版）→ 加入浮水印（可選） → 取代內容（可選） → 設定標題（可選）
        /// </remarks>
        /// <returns>讀取的 Word 檔案</returns>
        private Document __InitialWord(DocConvertConfigModel.WordToPdf cfgModel)
        {
            // 讀取 Word 檔案
            Document doc = GenerateDocument(cfgModel.TMPL_PATH);

            // 加入浮水印
            if (cfgModel.IS_DISPLAY_WATERMARK)
            {
                InsertWatermarkText(doc, cfgModel.WATERMARK_TEXT);
            }

            // 逐筆取代文件中的指定元素內容
            ReplaceTag(doc, cfgModel.DATA);

            // 設定文件標題
            SetDocTitle(doc, cfgModel.TITLE);

          

            return doc;
        }

        #endregion

        #region 產製 Word 通用（進階客製也可呼叫）

        /// <summary>
        /// 文件取代標籤樣版
        /// </summary>
        private readonly string replaceTagTmpl = "<<{0}>>";

        /// <summary>
        /// 預設的浮水印文字
        /// </summary>
        private readonly string defaultWatermarkText = "新北市政府衛生局";

        /// <summary>
        /// 初始化一份新的文件
        /// </summary>
        /// <param name="tmplPath">文件樣版讀取位置</param>
        /// <returns>文件</returns>
        public Document GenerateDocument(string tmplPath)
        {
            #region 防呆

            // 需指定讀取 Word 的路徑（樣版）
            if (string.IsNullOrEmpty(tmplPath))
            {
                throw new ArgumentNullException("tmplPath");
            }

            // 判斷指定 Word 的樣版是否存在
            if (!System.IO.File.Exists(tmplPath))
            {
                throw new FileNotFoundException("指定樣版檔案 Word 不存在!", Path.GetFileName(tmplPath));
            }

            #endregion

            // 設定 Aspose 的 license
            License license = new License();
            license.SetLicense("Aspose.Total.655.lic");

            // 讀取 Word 檔案
            return new Document(tmplPath);
        }

        /// <summary>
        /// Insert watermark into Document
        /// </summary>
        /// <param name="doc">Document</param>
        /// <param name="watermarkText">Watermark Text</param>
        /// <remarks>
        ///     Created by Steven Tsai at 2017/08/07
        /// </remarks>
        public void InsertWatermarkText(Document doc, String watermarkText = null)
        {
            // Use default watermarkText when watermarkText is null.
            watermarkText = string.IsNullOrEmpty(watermarkText) ? defaultWatermarkText : watermarkText;

            // Create a watermark shape. This will be a WordArt shape.
            // You are free to try other shape types as watermarks.
            Shape watermark = new Shape(doc, ShapeType.TextPlainText);

            // Set up the text of the watermark.
            watermark.TextPath.Text = watermarkText;
            watermark.TextPath.FontFamily = "Microsoft JhengHei";
            watermark.Width = 500;
            watermark.Height = 100;

            // Text will be directed from the bottom-left to the top-right corner.
            watermark.Rotation = -40;

            // Remove the following two lines if you need a solid black text.
            watermark.Fill.Color = System.Drawing.Color.LightGray; // Try LightGray to get more Word-style watermark
            watermark.StrokeColor = System.Drawing.Color.LightGray; // Try LightGray to get more Word-style watermark

            // Place the watermark in the page center.
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.WrapType = WrapType.None;
            watermark.VerticalAlignment = VerticalAlignment.Center;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;

            // Create a new paragraph and append the watermark to this paragraph.
            Paragraph watermarkPara = new Paragraph(doc);
            watermarkPara.AppendChild(watermark);

            // Insert the watermark into all headers of each document section.
            foreach (Section sect in doc.Sections)
            {
                // There could be up to three different headers in each section, since we want
                // the watermark to appear on all pages, insert into all headers.
                __InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderPrimary);
                __InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderFirst);
                __InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderEven);
            }
        }

        /// <summary>
        /// Insert watermark into Document Header
        /// </summary>
        /// <param name="watermarkPara">Watermark Paragraph</param>
        /// <param name="sect">Document Section</param>
        /// <param name="headerType">Document HeaderType</param>
        private void __InsertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, HeaderFooterType headerType)
        {
            HeaderFooter header = sect.HeadersFooters[headerType];

            if (header == null)
            {
                // There is no header of the specified type in the current section, create it.
                header = new HeaderFooter(sect.Document, headerType);
                sect.HeadersFooters.Add(header);
            }

            // Insert a clone of the watermark into the header.
            header.AppendChild(watermarkPara.Clone(true));
        }

        /// <summary>
        /// 文件標題設定
        /// </summary>
        /// <param name="doc">編輯中的文件</param>
        /// <param name="title">文件標題</param>
        public void SetDocTitle(Document doc, string title)
        {
            if (!string.IsNullOrEmpty(title))
            {
                doc.BuiltInDocumentProperties.Title = title;
            }
        }

        /// <summary>
        /// 取代文件中的標籤文字
        /// </summary>
        /// <param name="doc">編輯中的文件</param>
        /// <param name="props">欲取代的文字內容</param>
        public void ReplaceTag(Document doc, dynamic props)
        {
            if (props != null)
            {
                foreach (PropertyInfo prop in props.GetType().GetProperties())
                {
                    object value = prop.GetValue(props, null);
                    doc.Range.Replace(string.Format(replaceTagTmpl, prop.Name), value == null ? string.Empty : value.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                }
            }
        }

        /// <summary>
        /// 取代表格中的標籤文字
        /// </summary>
        /// <param name="table">編輯中的表格</param>
        /// <param name="props">欲取代的文字內容</param>
        public void ReplaceTag(Table table, dynamic props)
        {
            if (props != null)
            {
                foreach (PropertyInfo prop in props.GetType().GetProperties())
                {
                    object value = prop.GetValue(props, null);
                    table.Range.Replace(string.Format(replaceTagTmpl, prop.Name), value == null ? string.Empty : value.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                }
            }
        }

        #endregion
    }
    }