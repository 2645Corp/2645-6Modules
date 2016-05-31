using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Word;
namespace DocPrinter //CrossDomain.WebFinder
{
    public class WordHelper
    {
        public ApplicationClass wordApp;
        public Document wordDoc;
        object missing = System.Reflection.Missing.Value;


        public ApplicationClass WordApplication
        {
            get
            {
                return wordApp;// _wordApplication;
            }
        }
        public Document WordDocument
        {
            get
            {
                return wordDoc;
            }
        }

        public WordHelper()
        {
            wordApp = new ApplicationClass();//_wordApplication
        }
        public WordHelper(ApplicationClass wordApplication)
        {
            wordApp = wordApplication;//_wordApplication
        }
        public void CreateAndActive()
        {
            wordDoc = CreateOneDocument(missing, missing, missing, missing);
            wordDoc.Activate();
        }

        public bool OpenAndActive(string FileName, bool IsReadOnly, bool IsVisibleWin)
        {
            if (string.IsNullOrEmpty(FileName))
            {
                return false;
            }
            try
            {
                wordDoc = OpenOneDocument(FileName, missing, IsReadOnly, missing, missing, missing, missing, missing, missing, missing, missing, IsVisibleWin, missing, missing, missing, missing);
                wordDoc.Activate();
                return true;
            }
            catch
            {
                return false;
            }
        }
        public void Close()
        {
            if (wordDoc != null)
            {
                wordDoc.Close(ref missing, ref missing, ref missing);
                wordApp.Application.Quit(ref missing, ref missing, ref missing);//_wordApplication
            }
        }
        public void Save()
        {
            if (wordDoc == null)
            {
                wordDoc = wordApp.ActiveDocument;//_wordApplication
            }
            wordDoc.Save();
        }

        public void SaveAs(string FileName)
        {
            if (wordDoc == null)
            {
                wordDoc = wordApp.ActiveDocument;//_wordApplication
            }
            object objFileName = FileName;
            wordDoc.SaveAs(ref objFileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
        }
        public Document CreateOneDocument(object template, object newTemplate, object documentType, object visible)
        {
            return wordDoc = wordApp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);//_wordApplication
        }

        public Document OpenOneDocument(object FileName, object ConfirmConversions, object ReadOnly,
            object AddToRecentFiles, object PasswordDocument, object PasswordTemplate, object Revert,
            object WritePasswordDocument, object WritePasswordTemplate, object Format, object Encoding,
            object Visible, object OpenAndRepair, object DocumentDirection, object NoEncodingDialog, object XMLTransform)
        {
            try
            {//_wordApplication
                // return wordApp.Documents.Add(ref FileName, ref ConfirmConversions, ref ReadOnly, ref AddToRecentFiles,
                return wordApp.Documents.Open(ref FileName, ref ConfirmConversions, ref ReadOnly, ref AddToRecentFiles,
                   ref PasswordDocument, ref PasswordTemplate, ref Revert, ref WritePasswordDocument, ref WritePasswordTemplate,
                   ref Format, ref Encoding, ref Visible, ref OpenAndRepair, ref DocumentDirection, ref NoEncodingDialog, ref XMLTransform);
            }
            catch
            {
                return null;
            }
        }

        public bool GoToBookMark(string bookMarkName)
        {
            //是否存在书签
            if (wordDoc.Bookmarks.Exists(bookMarkName))
            {
                object what = WdGoToItem.wdGoToBookmark;
                object name = bookMarkName;
                GoTo(what, missing, missing, name);
                return true;
            }
            return false;
        }
        public void GoTo(object what, object which, object count, object name)
        {
            wordApp.Selection.GoTo(ref what, ref which, ref count, ref name);//_wordApplication
        }

        public void ReplaceBookMark(string bookMarkName, string text)
        {
            bool isExist = GoToBookMark(bookMarkName);
            if (isExist)
            {
                InsertText(text);
            }
        }
        public bool Replace(string oldText, string newText, string replaceType, bool isCaseSensitive)
        {
            if (wordDoc == null)
            {
                wordDoc = wordApp.ActiveDocument;//_wordApplication
            }
            object findText = oldText;
            object replaceWith = newText;
            object wdReplace;
            object matchCase = isCaseSensitive;
            switch (replaceType)
            {
                case "All":
                    wdReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
                    break;
                case "None":
                    wdReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
                    break;
                case "FirstOne":
                    wdReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne;
                    break;
                default:
                    wdReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne;
                    break;
            }
            wordDoc.Content.Find.ClearFormatting();//移除Find的搜索文本和段落格式设置
            return wordDoc.Content.Find.Execute(ref findText, ref matchCase, ref missing, ref missing,
                  ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                  ref wdReplace, ref missing, ref missing, ref missing, ref missing);
        }

        public void InsertText(string text)
        {
            wordApp.Selection.TypeText(text);//_wordApplication
        }
        public void InsertLineBreak()
        {
            wordApp.Selection.TypeParagraph();//_wordApplication
        }

        public void InsertPageBreak()
        {
            //插入分页   if(i!=DT.Rows.Count-1)
            {
                object mymissing = System.Reflection.Missing.Value;
                object myunit = WdUnits.wdStory;
                wordApp.Selection.EndKey(ref myunit, ref mymissing);
                object pBreak = (int)WdBreakType.wdPageBreak;
                wordApp.Selection.InsertBreak(ref pBreak);
            }
        }
        public void InsertPic(string fileName)
        {
            object range = wordApp.Selection.Range;//_wordApplication
            InsertPic(fileName, missing, missing, range);
        }

        public void InsertPic(string fileName, float width, float height)
        {
            object range = wordApp.Selection.Range;//_wordApplication
            InsertPic(fileName, missing, missing, range, width, height);
        }

        public void InsertPic(string fileName, float width, float height, string caption)
        {
            object range = wordApp.Selection.Range;//_wordApplication
            InsertPic(fileName, missing, missing, range, width, height, caption);
        }

        public void InsertPic(string FileName, object LinkToFile, object SaveWithDocument, object Range, float Width, float Height, string caption)
        {//_wordApplication
            wordApp.Selection.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Range).Select();
            if (Width > 0)
            {
                wordApp.Selection.InlineShapes[1].Width = Width;//_wordApplication
            }
            if (Height > 0)
            {
                wordApp.Selection.InlineShapes[1].Height = Height;//_wordApplication
            }
            object Label = WdCaptionLabelID.wdCaptionFigure;
            object Title = caption;
            object Position = WdCaptionPosition.wdCaptionPositionBelow;
            object ExcludeLabel = true;
            wordApp.Selection.InsertCaption(ref Label, ref Title, "文档", ref Position, ref ExcludeLabel);//_wordApplication
            MoveRight();
        }

        public void InsertPic(string FileName, object LinkToFile, object SaveWithDocument, object Range, float Width, float Height)
        {
            wordApp.Selection.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Range).Select();//_wordApplication
            wordApp.Selection.InlineShapes[1].Width = Width;//_wordApplication
            wordApp.Selection.InlineShapes[1].Height = Height;//_wordApplication
            MoveRight();
        }

        public void InsertPic(string FileName, object LinkToFile, object SaveWithDocument, object Range)
        {
            wordApp.Selection.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Range);//_wordApplication
        }

        public void InsertBookMark(string bookMarkName)
        {
            //存在则先删除
            if (wordDoc.Bookmarks.Exists(bookMarkName))
            {
                DeleteBookMark(bookMarkName);
            }
            object range = wordApp.Selection.Range;//_wordApplication
            wordDoc.Bookmarks.Add(bookMarkName, ref range);
        }

        public void DeleteBookMark(string bookMarkName)
        {
            if (wordDoc.Bookmarks.Exists(bookMarkName))
            {
                var bookMarks = wordDoc.Bookmarks;
                for (int i = 1; i <= bookMarks.Count; i++)
                {
                    object index = i;
                    var bookMark = bookMarks.get_Item(ref index);
                    if (bookMark.Name == bookMarkName)
                    {
                        bookMark.Delete();
                        break;
                    }
                }
            }
        }

        public void DeleteAllBookMark()
        {
            for (; wordDoc.Bookmarks.Count > 0; )
            {
                object index = wordDoc.Bookmarks.Count;
                var bookmark = wordDoc.Bookmarks.get_Item(ref index);
                bookmark.Delete();
            }
        }

        public Table AddTable(int NumRows, int NumColumns)
        {
            return AddTable(wordApp.Selection.Range, NumRows, NumColumns, missing, missing);//_wordApplication
        }

        public Table AddTable(int NumRows, int NumColumns, WdAutoFitBehavior AutoFitBehavior)
        {
            return AddTable(wordApp.Selection.Range, NumRows, NumColumns, missing, AutoFitBehavior);//_wordApplication
        }

        public Table AddTable(Range Range, int NumRows, int NumColumns, object DefaultTableBehavior, object AutoFitBehavior)
        {
            if (wordDoc == null)
            {
                wordDoc = wordApp.ActiveDocument;//_wordApplication
            }
            return wordDoc.Tables.Add(Range, NumRows, NumColumns, ref DefaultTableBehavior, ref AutoFitBehavior);
        }

        public Row AddRow(Table table)
        {
            return AddRow(table, missing);
        }

        public Row AddRow(Table table, object beforeRow)
        {
            return table.Rows.Add(ref beforeRow);
        }
        public void InsertRows(int numRows)
        {
            object NumRows = numRows;
            object wdCollapseStart = WdCollapseDirection.wdCollapseStart;
            wordApp.Selection.InsertRows(ref NumRows);//_wordApplication
            wordApp.Selection.Collapse(ref wdCollapseStart);//_wordApplication
        }

        public void MoveLeft(WdUnits unit = WdUnits.wdCharacter, int count = 1, int extend_flag = 0)
        {
            object extend;
            if (extend_flag == 1) extend = WdMovementType.wdExtend;
            else extend = missing;
            wordApp.Selection.MoveLeft(unit, count, extend);//_wordApplication
        }

        public void MoveUp(WdUnits unit = WdUnits.wdCharacter, int count = 1, int extend_flag = 0)
        {
            object extend;
            if (extend_flag == 1) extend = WdMovementType.wdExtend;
            else extend = missing;
            wordApp.Selection.MoveUp(unit, count, extend);//_wordApplication
        }
        public void MoveRight(WdUnits unit = WdUnits.wdCharacter, int count = 1, int extend_flag = 0)
        {
            object extend;
            if (extend_flag == 1) extend = WdMovementType.wdExtend;
            else extend = missing;
            wordApp.Selection.MoveRight(unit, count, extend);//_wordApplication
        }
        public void MoveDown(WdUnits unit = WdUnits.wdCharacter, int count = 1, int extend_flag = 0)
        {
            object extend;
            if (extend_flag == 1) extend = WdMovementType.wdExtend;
            else extend = missing;
            wordApp.Selection.MoveDown(unit, count, extend);//_wordApplication
        }
        //2 另外的地方复制过来

        public void SetLinesPage(int size = 40)
        {
            wordApp.ActiveDocument.PageSetup.LinesPage = size;
        }
        public void SetPageHeaderFooter(string context, int HeaderFooter = 0)
        {

            //            wordApp.ActiveWindow.View.Type = Word.WdViewType.wdOutlineView;
            //            wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekPrimaryHeader;
            //            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter(context);
            //            wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //跳出页眉设置   
            //            wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
            //
            ////添加页眉方法二：
            if (wordApp.ActiveWindow.ActivePane.View.Type == WdViewType.wdNormalView ||
                wordApp.ActiveWindow.ActivePane.View.Type == WdViewType.wdOutlineView)
            {
                wordApp.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
            }
            if (HeaderFooter == 0)
            {//设置页眉内容";
                wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
            }
            else
            {//设置页脚内容";
                wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
            }
            wordApp.Selection.HeaderFooter.LinkToPrevious = false;
            wordApp.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.Selection.HeaderFooter.Range.Text = context;// "页眉 页脚 内容";
            //跳出页眉页脚设置
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;
        }
        public void InsertFormatText(string context, int fontSize, WdColor fontColor, int fontBold, string familyName, WdParagraphAlignment align)
        {
            //设置字体样式以及方向   
            wordApp.Application.Selection.Font.Size = fontSize;
            wordApp.Application.Selection.Font.Bold = fontBold;
            wordApp.Application.Selection.Font.Color = fontColor;
            wordApp.Selection.Font.Name = familyName;
            wordApp.Application.Selection.ParagraphFormat.Alignment = align;
            wordApp.Application.Selection.TypeText(context);
        }

        public Table CreatTable(int rowNum, int cellNum)
        {
            return this.wordDoc.Tables.Add(wordApp.Selection.Range, rowNum, cellNum, ref missing, ref missing);
        }
        public void setPageLay(string paper = "A4", int orient = 0)
        {
            //页面设置，设置页面为纵向布局，设置纸张类型为A4纸
            if (orient == 0)
            {
                wordDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            }
            else
            {
                wordDoc.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
            }
            if (paper == "A4")
            {
                wordDoc.PageSetup.PageWidth = wordApp.CentimetersToPoints(29.7F);
                wordDoc.PageSetup.PageHeight = wordApp.CentimetersToPoints(21F);
            }
        }
        public Table getTable(int index)
        {
            //    int table_num = this.wordDoc.Tables.Count;
            //    Word.Tables tables = wordDoc.Tables;
            //    return wordDoc.Content.Tables[index];
            //    wordDoc.Tables.Item(index);

            int no = 0;
            foreach (Table table in wordDoc.Tables)//tables)
            {
                if (no == index) return table;
                no++;
            }
            return null;
        }

        public void SetColumnWidth(float[] widths, Table tb)
        {
            if (widths.Length > 0)
            {
                int len = widths.Length;
                for (int i = 0; i < len; i++)
                {
                    tb.Columns[i + 1].Width = widths[i];
                }
            }
        }

        public void MergeColumn(Table tb, Cell[] cells)
        {
            if (cells.Length > 1)
            {
                Cell c = cells[0];
                int len = cells.Length;
                for (int i = 1; i < len; i++)
                {
                    c.Merge(cells[i]);
                }
            }
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
        }
    }
}