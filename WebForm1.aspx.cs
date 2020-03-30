using System.Data;
using System.Configuration;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO;
using System;
using MSWord = Microsoft.Office.Interop.Word;

namespace WebApplication2
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        /// <summary> 
        /// 将word转换为Html 
        /// </summary> 
        /// <param name="sender"></param> 
        /// <param name="e"></param> 
        protected void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {

                //上传 

                //应当先把文件上传至服务器然后再解析文件为html 
                string filePath = uploadWord(File1);
                var datas = new Dictionary<string, string>();
                datas.Add("${姓名}", "BlockChain");
                datas.Add("${合同编号}", "0000111122223333");

                if (WordReplace(filePath, datas))
                {
                    WordConvertPDF(filePath, Server.MapPath("~\\html\\pd.pdf"), MSWord.WdExportFormat.wdExportFormatPDF);
                }
                //转换 
                // wordToHtml(File1);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Response.Write("恭喜，转换成功！");
            }

        }

        //上传文件并转换为html wordToHtml(wordFilePath) 
        ///<summary> 
        ///上传文件并转存为html 
        ///</summary> 
        ///<param name="wordFilePath">word文档在客户机的位置</param> 
        ///<returns>上传的html文件的地址</returns> 
        public string wordToHtml(System.Web.UI.HtmlControls.HtmlInputFile wordFilePath)
        {
            MSWord.ApplicationClass word = new MSWord.ApplicationClass();
            Type wordType = word.GetType();
            MSWord.Documents docs = word.Documents;

            // 打开文件 
            Type docsType = docs.GetType();

            //应当先把文件上传至服务器然后再解析文件为html 
            string filePath = uploadWord(wordFilePath);

            //判断是否上传文件成功 
            if (filePath == "0")
                return "0";
            //判断是否为word文件 
            if (filePath == "1")
                return "1";

            object fileName = filePath;

            MSWord.Document doc = (MSWord.Document)docsType.InvokeMember("Open",
            System.Reflection.BindingFlags.InvokeMethod, null, docs, new Object[] { fileName, true, true });

            // 转换格式，另存为html 
            Type docType = doc.GetType();

            string filename = System.DateTime.Now.Year.ToString() + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Day.ToString() +
            System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString() + System.DateTime.Now.Second.ToString();

            // 判断指定目录下是否存在文件夹，如果不存在，则创建 
            if (!Directory.Exists(Server.MapPath("~\\html")))
            {
                // 创建up文件夹 
                Directory.CreateDirectory(Server.MapPath("~\\html"));
            }

            //被转换的html文档保存的位置 
            string ConfigPath = HttpContext.Current.Server.MapPath("html/" + filename + ".html");
            object saveFileName = ConfigPath;

            /*下面是Microsoft Word 9 Object Library的写法，如果是10，可能写成： 
            * docType.InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod, 
            * null, doc, new object[]{saveFileName, Word.WdSaveFormat.wdFormatFilteredHTML}); 
            * 其它格式： 
            * wdFormatHTML 
            * wdFormatDocument 
            * wdFormatDOSText 
            * wdFormatDOSTextLineBreaks 
            * wdFormatEncodedText 
            * wdFormatRTF 
            * wdFormatTemplate 
            * wdFormatText 
            * wdFormatTextLineBreaks 
            * wdFormatUnicodeText 
            */
            docType.InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod,
            null, doc, new object[] { saveFileName, MSWord.WdSaveFormat.wdFormatFilteredHTML });

            //关闭文档 
            docType.InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod,
            null, doc, new object[] { null, null, null });

            // 退出 Word 
            wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, word, null);
            //转到新生成的页面 
            return ("/" + filename + ".html");

        }


        public string uploadWord(System.Web.UI.HtmlControls.HtmlInputFile uploadFiles)
        {
            if (uploadFiles.PostedFile != null)
            {
                string fileName = uploadFiles.PostedFile.FileName;

                int extendNameIndex = fileName.LastIndexOf(".");
                string extendName = fileName.Substring(extendNameIndex);
                string newName = "";
                try
                {
                    //验证是否为word格式 
                    if (extendName == ".doc" || extendName == ".docx")
                    {

                        DateTime now = DateTime.Now;
                        newName = now.DayOfYear.ToString() + uploadFiles.PostedFile.ContentLength.ToString();

                        // 判断指定目录下是否存在文件夹，如果不存在，则创建 
                        if (!Directory.Exists(Server.MapPath("~\\wordTmp")))
                        {
                            // 创建up文件夹 
                            Directory.CreateDirectory(Server.MapPath("~\\wordTmp"));
                        }

                        //上传路径 指当前上传页面的同一级的目录下面的wordTmp路径 
                        uploadFiles.PostedFile.SaveAs(System.Web.HttpContext.Current.Server.MapPath("wordTmp/" + newName + extendName));
                    }
                    else
                    {
                        return "1";
                    }
                }
                catch
                {
                    return "0";
                }
                //return "http://" + HttpContext.Current.Request.Url.Host + HttpContext.Current.Request.ApplicationPath + "/wordTmp/" + newName + extendName; 
                return System.Web.HttpContext.Current.Server.MapPath("wordTmp/" + newName + extendName);
            }
            else
            {
                return "0";
            }
        }

        public static bool WordReplace(string filePath, Dictionary<string, string> datas)
        {
            bool result = false;
            Microsoft.Office.Interop.Word._Application app = null;
            Microsoft.Office.Interop.Word._Document doc = null;

            object nullobj = Type.Missing;
            object file = filePath;


            try
            {
                app = new Microsoft.Office.Interop.Word.Application();//new Microsoft.Office.Interop.Word.ApplicationClass();

                doc = app.Documents.Open(
                    ref file, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj, ref nullobj) as Microsoft.Office.Interop.Word._Document;
                object tmp = "tt1";
                object missingValue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Range startRange = app.ActiveDocument.Bookmarks.get_Item(ref tmp).Range;


                //删除指定书签位置后的第一个表格
                //Microsoft.Office.Interop.Word.Table tbl = startRange.Tables[1];
                //tbl.Delete();

                //添加表格
                for (int i = 0; i < 8; i++)
                {
                    MSWord.Table table = doc.Tables.Add(startRange, 3, 2, ref missingValue, ref missingValue);
                    table.AllowAutoFit = true;
                    table.Borders.Enable = 1;//这个值可以设置得很大，例如5、13等等
                    table.Rows.HeightRule = MSWord.WdRowHeightRule.wdRowHeightExactly;//高度规则是：行高有最低值下限？
                    table.Rows.Height = 15;// 
                    table.Columns.PreferredWidthType = MSWord.WdPreferredWidthType.wdPreferredWidthPercent;

                    table.Columns[1].PreferredWidth = 30;
                    table.Columns[2].PreferredWidth = 70;

                    table.Range.Font.Size = 10.5F;
                    table.Range.Font.Bold = 0;
                    table.Cell(1, 1).Range.Text = "文件名";
                    table.Cell(2, 1).Range.Text = "SHA256校验码";
                    table.Cell(3, 1).Range.Text = "时间戳时间";
                    table.Cell(1, 2).Range.Text = "智臻链";
                    table.Cell(2, 2).Range.Text = "rerererererererererererererererererererererere\r\nerererererererererererererererererererererererer\r\nerererererererererererererererererererererererererere";
                    table.Cell(2, 2).HeightRule = MSWord.WdRowHeightRule.wdRowHeightAuto;
                    table.Cell(3, 2).Range.Text = "15393034920943048";

                }




                object objReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;


                foreach (var item in datas)
                {
                    app.Selection.Find.ClearFormatting();
                    app.Selection.Find.Replacement.ClearFormatting();
                    string oldStr = item.Key;//需要被替换的文本
                    string newStr = item.Value; //替换文本 


                    if (newStr == null)
                    {
                        newStr = "";
                    }
                    int newStrLength = newStr.Length;
                    if (newStrLength <= 255) // 长度未超过255时，直接替换
                    {
                        app.Selection.Find.Text = oldStr;
                        app.Selection.Find.Replacement.Text = newStr;


                        app.Selection.Find.Execute(ref nullobj, ref nullobj, ref nullobj,
                                     ref nullobj, ref nullobj, ref nullobj,
                                     ref nullobj, ref nullobj, ref nullobj,
                                     ref nullobj, ref objReplace, ref nullobj,
                                     ref nullobj, ref nullobj, ref nullobj);
                        continue;
                    }


                    //word文本替换长度不能超过255
                    int hasSubLength = 0;
                    int itemLength = 255 - oldStr.Length;
                    while (true)
                    {
                        string itemStr = "";
                        int subLength = 0;
                        if (newStrLength - hasSubLength <= 255)  // 剩余的内容不超过255，直接替换
                        {
                            app.Selection.Find.ClearFormatting();
                            app.Selection.Find.Replacement.ClearFormatting();


                            app.Selection.Find.Text = oldStr;
                            app.Selection.Find.Replacement.Text = newStr.Substring(hasSubLength, newStrLength - hasSubLength);


                            app.Selection.Find.Execute(ref nullobj, ref nullobj, ref nullobj,
                                          ref nullobj, ref nullobj, ref nullobj,
                                          ref nullobj, ref nullobj, ref nullobj,
                                          ref nullobj, ref objReplace, ref nullobj,
                                          ref nullobj, ref nullobj, ref nullobj);


                            break; // 结束循环
                        }


                        // 由于Word中换行为“^p”两个字符不能分割
                        // 如果分割位置将换行符分开了，则本次少替换一个字符
                        if (newStr.Substring(hasSubLength, itemLength).EndsWith("^") &&
                            newStr.Substring(hasSubLength, itemLength + 1).EndsWith("p"))
                        {
                            subLength = itemLength - 1;
                        }
                        else
                        {
                            subLength = itemLength;
                        }


                        itemStr = newStr.Substring(hasSubLength, subLength) + oldStr;


                        app.Selection.Find.ClearFormatting();
                        app.Selection.Find.Replacement.ClearFormatting();


                        app.Selection.Find.Text = oldStr;
                        app.Selection.Find.Replacement.Text = itemStr;


                        app.Selection.Find.Execute(ref nullobj, ref nullobj, ref nullobj,
                                      ref nullobj, ref nullobj, ref nullobj,
                                      ref nullobj, ref nullobj, ref nullobj,
                                      ref nullobj, ref objReplace, ref nullobj,
                                      ref nullobj, ref nullobj, ref nullobj);


                        hasSubLength += subLength;
                    }
                }

                //    object tmp = "tttttt";
                //object missingValue = System.Reflection.Missing.Value;

                //Microsoft.Office.Interop.Word.Range startRange = app.ActiveDocument.Bookmarks.get_Item(ref tmp).Range;


                ////删除指定书签位置后的第一个表格
                //Microsoft.Office.Interop.Word.Table tbl = startRange.Tables[1];
                //    tbl.Delete();

                ////添加表格
                //doc.Tables.Add(startRange, 5, 4, ref missingValue, ref missingValue);

                //    //为表格划线
                //    startRange.Tables[1].Borders[MSWord.WdBorderType.wdBorderTop].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                //    startRange.Tables[1].Borders[MSWord.WdBorderType.wdBorderLeft].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                //    startRange.Tables[1].Borders[MSWord.WdBorderType.wdBorderRight].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                //    startRange.Tables[1].Borders[MSWord.WdBorderType.wdBorderBottom].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                //    startRange.Tables[1].Borders[MSWord.WdBorderType.wdBorderHorizontal].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                //    startRange.Tables[1].Borders[MSWord.WdBorderType.wdBorderVertical].LineStyle = MSWord.WdLineStyle.wdLineStyleSingle;


                //保存
                doc.Save();


                result = true;
            }
            catch (Exception e)
            {
                result = false;
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(ref nullobj, ref nullobj, ref nullobj);
                    doc = null;
                }
                if (app != null)
                {
                    app.Quit(ref nullobj, ref nullobj, ref nullobj);
                    app = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;

        }

        private bool WordConvertPDF(string sourcePath, string targetPath, MSWord.WdExportFormat exportFormat)
        {
            bool result;
            object paramMissing = Type.Missing;
            MSWord.ApplicationClass wordApplication = new MSWord.ApplicationClass();
            MSWord.Document wordDocument = null;
            try
            {
                object paramSourceDocPath = sourcePath;
                string paramExportFilePath = targetPath;

                MSWord.WdExportFormat paramExportFormat = exportFormat;
                bool paramOpenAfterExport = false;
                MSWord.WdExportOptimizeFor paramExportOptimizeFor =
                        MSWord.WdExportOptimizeFor.wdExportOptimizeForPrint;
                MSWord.WdExportRange paramExportRange = MSWord.WdExportRange.wdExportAllDocument;
                int paramStartPage = 0;
                int paramEndPage = 0;
                MSWord.WdExportItem paramExportItem = MSWord.WdExportItem.wdExportDocumentContent;
                bool paramIncludeDocProps = true;
                bool paramKeepIRM = true;
                MSWord.WdExportCreateBookmarks paramCreateBookmarks =
                        MSWord.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                bool paramDocStructureTags = true;
                bool paramBitmapMissingFonts = true;
                bool paramUseISO19005_1 = false;

                wordDocument = wordApplication.Documents.Open(
                        ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing);

                if (wordDocument != null)
                    wordDocument.ExportAsFixedFormat(paramExportFilePath,
                            paramExportFormat, paramOpenAfterExport,
                            paramExportOptimizeFor, paramExportRange, paramStartPage,
                            paramEndPage, paramExportItem, paramIncludeDocProps,
                            paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                            paramBitmapMissingFonts, paramUseISO19005_1,
                            ref paramMissing);
                result = true;
            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

    }
}