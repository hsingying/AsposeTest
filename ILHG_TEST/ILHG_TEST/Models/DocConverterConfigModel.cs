using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ILHG_TEST.Models
{
    public class DocConvertConfigModel
    {
        /// <summary>
        /// 將 Word 轉成 PDF 檔案輸出的設定
        /// </summary>
        /// <remarks>
        ///     如欲增加其他設定, 可自行擴充
        /// </remarks>
        public class WordToPdf
        {
            /// <summary>
            /// 讀取 Word 的路徑（樣版）
            /// </summary>
            public string TMPL_PATH { get; set; }

            /// <summary>
            /// 產出 Pdf 的路徑（含檔案名稱）
            /// </summary>
            public string TARGET_PATH { get; set; }

            /// <summary>
            /// 取代樣版內容
            /// </summary>
            /// <remarks>
            ///     使用 Reflection 物件, 逐個取代樣版內指定 Property 的值
            /// </remarks>
            public dynamic DATA { get; set; }

            /// <summary>
            /// 是否顯示浮水印文字
            /// </summary>
            public bool IS_DISPLAY_WATERMARK { get; set; }

            /// <summary>
            /// 浮水印文字
            /// </summary>
            public string WATERMARK_TEXT { get; set; }

            /// <summary>
            /// 顯示標題
            /// </summary>
            /// <remarks>
            ///     此設定會影響瀏覽器上 PDF 檢視器中所顯示的標題
            /// </remarks>
            public string TITLE { get; set; }
            
        }

        /// <summary>
        /// 圖片資訊
        /// </summary>
        //public class ImgInfo
        //{
        //    /// <summary>
        //    /// 圖片串流
        //    /// </summary>
        //    public MemoryStream STREAM { get; set; }

        //    /// <summary>
        //    /// 圖片 X 位置
        //    /// </summary>
        //    public double POSITION_X { get; set; }

        //    /// <summary>
        //    /// 圖片 Y 位置
        //    /// </summary>
        //    public double POSITION_Y { get; set; }

        //    /// <summary>
        //    /// 圖片高度
        //    /// </summary>
        //    public double HEIGHT { get; set; }

        //    /// <summary>
        //    /// 圖片寬度
        //    /// </summary>
        //    public double WIDTH { get; set; }
        //}
    }
}