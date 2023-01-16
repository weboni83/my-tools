using DevExpress.Pdf;
using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Presentation.ViewModels;
using System.Threading.Tasks;
using System.Threading;
using System.Drawing;
using System.Diagnostics;
using DevExpress.XtraSplashScreen;
using Core.Exceptions;

namespace Presentation.Views.Excel
{
    public partial class EBRViewer : UserControl
    {
        const string BLOB_FIELD = "file_stream";
        const string MBR_TYPE = "mbr_type";
        const string CONNECTION_STRING = "";
        const string DEFAULT_SQL = @"
SELECT b.mbr_type, a.order_process_id, a.file_stream, a.added_page_no, c.template_id, c.sp_name, a.order_detailproc_id, a.mbr_worksheet_id , a.ebr_worksheet_id, b.process_nm
,case when d.order_process_id IS NOT NULL then CONVERT(BIT, 1) else CONVERT(BIT, 0) END AS IS_SELECTED
FROM ebr_order_worksheet a 
INNER JOIN op_order_process b 
	ON	 a.order_process_id = b.order_process_id
LEFT JOIN mbr_template c 
	ON	a.template_id = c.template_id
LEFT JOIN op_ebr_request_detail d 
	ON	a.plant_cd = d.plant_cd 
	AND a.order_process_id = d.order_process_id
	AND d.ebr_request_id = 2593
WHERE a.file_type = 'PROCESS'
	AND a.order_no = 'ADPM22-008'
	AND a.item_type = 'MakingItem'
ORDER BY order_process_seq, ISNULL(added_page_no, 0)
";
        public EventHandler<FormExceptionEventArgs> ExceptionTrigger;
        public EBRViewer()
        {
            InitializeComponent();

            gridView_excelList.FocusedRowChanged += (s, e) =>
            {
                try
                {
                    var mbrType = gridView_excelList.GetFocusedRowCellValue(BLOB_FIELD);
                    var obj = gridView_excelList.GetFocusedRowCellValue(BLOB_FIELD);
                    byte[] blob = null;
                    if(obj != null)
                        blob = (obj as byte[]);

                    spreadsheetControl1.LoadDocument(blob, DocumentFormat.Xlsx);
                }
                catch (Exception ex)
                {
                    // add event handler
                    throw new Exception(ex.Message);
                }
                finally
                {

                }
            };

            this.Load += (s, e) =>
            {
                memoEdit_sql.Text = DEFAULT_SQL;
            };

            this.dockPanel_scripts.CustomButtonClick += (s, e) =>
            {
                ICommand command = null;
                if(e.Button.Properties.Caption.ToUpper().Contains("RUN"))
                    command = new SelectCommand(CONNECTION_STRING);

                if(command is null)
                    return;

                SplashScreenManager.ShowForm(this.ParentForm, typeof(SplashScreenForm), true, true, false);
                try
                {
                    SplashScreenManager.Default.SendCommand(SplashScreenForm.SplashScreenCommand.SetMessage, "조회");
                    SelectCommand(command);
                }
                finally
                {

                    SplashScreenManager.CloseForm(false);
                }
            };

            this.dockPanel_pdfViewer.CustomButtonClick += (s, e) =>
            {
                ConvertWorkSheetToPDF();
            };

            this.pdfViewer1.ScrollPositionChanged += (s, e) =>
            {
                barStaticItem_message.Caption = $"{pdfViewer1.CurrentPageNumber}/{pdfViewer1.PageCount}";
            };

        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if(keyData == Keys.F5)
            {
                ICommand command = new SelectCommand(CONNECTION_STRING);
                SelectCommand(command);
                return true;
            }

            if (keyData == Keys.F9)
            {
                ConvertWorkSheetToPDF();
                pdfViewer1.NavigationPaneWidth = 200;
                pdfViewer1.NavigationPaneVisibility = DevExpress.XtraPdfViewer.PdfNavigationPaneVisibility.Visible;
                dockPanel_pdfViewer.ShowSliding();

                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        //void RunScript(ICommand command)
        void SelectCommand(ICommand command)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            try
            {
                command.Execute(memoEdit_sql.Text);

                DataTable dt = command.GetCache();
                gridControl_excelList.DataSource = dt;
                ((DevExpress.XtraGrid.Views.Grid.GridView)gridControl_excelList.MainView).BestFitColumns();
            }
            catch(Exception ex)
            {
                ExceptionTrigger?.Invoke(this, new FormExceptionEventArgs(ex.Message));
            }
            finally
            {
                barStaticItem_performance.Caption = $"{stopwatch.Elapsed.ToString("mm\\:ss\\.ff")}";
            }
        }

        void ConvertWorkSheetToPDF()
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            try
            {
                byte[] blob = null;
                using(MemoryStream ms = new MemoryStream())
                {
                    spreadsheetControl1.SaveDocument(ms, DocumentFormat.Xlsx);
                    blob = ms.ToArray();
                }

                Workbook workbook = new Workbook();
                workbook.History.IsEnabled = false;
                workbook.BeginUpdate();

                workbook.LoadDocument(blob, DocumentFormat.Xlsx);
                // PDF 로 변환 하기 전에 엑셀 마진 조정
                fncDocumentSettingBeforeExport(MBR_TYPE, workbook);

                #region Sheet 에서 페이지 목차를 만드는 방법 ( process followed by detail process)
                Worksheet sheet = workbook.Worksheets.ActiveWorksheet;

                Range usedRange = sheet.GetUsedRange();
                const int SUBNUMBER_COLUMN_INDEX = 0;
                const int SUBMUMBER_TEXT_INDEX = 1;

                //기존 이름 관리자(DefinedName) 삭제 후 재등록. 삭제시 공정, 세부공정 이름 관리자는 삭제하지 않는다.
                TreeNode<Page> pageNode = new TreeNode<Page>(new Page() { Description = "공정" });
                DefinedName[] definedNames = sheet.DefinedNames
                    .Where(d => d.Name.StartsWith("세부공정"))
                    .OrderBy(d => d.Range.TopRowIndex).ToArray();
                foreach(DefinedName definedName in definedNames)
                {
                    var tocNo = usedRange[definedName.Range.TopRowIndex, SUBNUMBER_COLUMN_INDEX].Value.TextValue;
                    var tocText = usedRange[definedName.Range.TopRowIndex, definedName.Range.LeftColumnIndex].Value.TextValue;
                    pageNode.AddChild(new Page() { Index = definedName.Range.TopRowIndex
                        , Chapter = tocNo
                        , Description = tocText
                    });
                        
                }

                const string PAGE_PREAKE_TEXT = "N/A";
                int cursorNO = 1;
                for(int i = 0; i < usedRange.RowCount; i++)
                {
                    // NOTE: { Cell: A21 formula:"" value: "6" type: Numeric}
                    Cell subNOCell = usedRange[i, SUBNUMBER_COLUMN_INDEX];
                    // NOTE: 공정과 세부공정 사이에 붙임 문서등이 삽입된 경우, FindedNode가 Null일 수 있으므로 예외처리 한다.
                    var findedNode = pageNode.Children.Where(p => p.Value.Index < subNOCell.RowIndex).LastOrDefault();
                    if(findedNode == null)
                        continue;

                    if(subNOCell.Value.Type == CellValueType.None)
                        continue;

                    int no = 0;

                    if(subNOCell.Value.Type == CellValueType.Numeric)
                        no = ((int)subNOCell.Value.NumericValue);

                    if (subNOCell.Value.Type == CellValueType.Text)
                        int.TryParse(subNOCell.Value.TextValue, out no);

                    if(no > 0)
                    {
                        Cell subTextCell = usedRange[i, SUBMUMBER_TEXT_INDEX];
                        var subText = subTextCell.Value.TextValue;

                        findedNode.AddChild(new Page()
                        {
                            Index = subNOCell.RowIndex
                            , LogicalNumber = no
                            , PysicalNumber = cursorNO
                            , Chapter = $"{findedNode.Value.Chapter}.{no}"
                            , Description = subText
                        });

                        Console.WriteLine($"IDX:{subNOCell.RowIndex}, {findedNode.Value.Description} => LogicalNO:{findedNode.Value.Chapter}.{no} PysicalNO:{cursorNO}, MAX:{findedNode.Children.Max(m=> m.Value.PysicalNumber)}");
                    }

                    if(subNOCell.Value.TextValue == PAGE_PREAKE_TEXT)
                    {
                        cursorNO++;
                        //findedNode.Children.ToList().ForEach(p => p.Value.PysicalNumber = findedNode.Value.PysicalNumber);
                        Console.WriteLine($"IDX:{subNOCell.RowIndex}, {findedNode.Value.Description} => Fined Page Break TEXT Logical Page NO => {findedNode.Children.Count}");
                    }

                }

                #endregion

                workbook.EndUpdate();
                workbook.History.IsEnabled = true;

                // 1. Load
                PdfDocumentProcessor processPdf = new PdfDocumentProcessor();
                processPdf.CreateEmptyDocument();
                byte[] bytes = null;
                using(MemoryStream ms = new MemoryStream())
                {
                    workbook.ExportToPdf(ms);
                    processPdf.AppendDocument(ms);
                    bytes = ms.ToArray();
                }

                using(MemoryStream ms = new MemoryStream(bytes))
                {
                    // 북마크 만들어서 확인하자
                    byte[] bookmarkBlob = GetCreateBookmarks(ms, pageNode);
                    // 3. PDF Load (Viewer 에서 PDF 전체를 보려면 파일로 저장하거나 메모리스트림을 유지해야 한다. using X)
                    pdfViewer1.LoadDocument(new MemoryStream(bookmarkBlob));
                }


                
            }
            catch(Exception ex)
            {
                ExceptionTrigger?.Invoke(this, new FormExceptionEventArgs(ex.Message));
            }
            finally
            {
                barStaticItem_performance.Caption = $"{stopwatch.Elapsed.ToString("mm\\:ss\\.ff")}";
            }
        }

        byte[] GetCreateBookmarks(MemoryStream memoryStream, TreeNode<Page> pageNode)
        {
            byte[] bookmarksBlob = null;
            using(PdfDocumentProcessor documentProcessor = new PdfDocumentProcessor())
            {
                // Load a document
                documentProcessor.LoadDocument(memoryStream);

                // Specify search parameters
                PdfTextSearchParameters searchParameters = new PdfTextSearchParameters();
                searchParameters.CaseSensitive = false;
                searchParameters.WholeWords = true;

                foreach(var node in pageNode.Children)
                {
                    var page = node.Value;
                    var word = page.Description;

                    Console.WriteLine($"{page.Chapter} {word}");


                    if(word.Contains("%"))
                    {
                        int lastSpaceIndex = word.LastIndexOf('%');
                        word = word.Substring(0, lastSpaceIndex);
                    }
                    // Get search results
                    PdfTextSearchResults results = null;

                    do
                    {
                        results = documentProcessor.FindText(word, searchParameters);

                        if(results.Status == PdfTextSearchStatus.Found)
                        {
                            var lastItem = pageNode.Children.Where(c => c.Value.PysicalNumber > 0).LastOrDefault();

                            if(lastItem !=null)
                                if(results.PageNumber < lastItem.Value.PysicalNumber)
                                    continue;

                            break;
                        }

                        if(results.Status == PdfTextSearchStatus.NotFound)
                            break;

                    } while(results.Status != PdfTextSearchStatus.Finished);

                    // 세부공정을 못찾으면 다음으로
                    if(results.Status == PdfTextSearchStatus.NotFound)
                    {
                        Console.WriteLine($"#. Exception {page.Chapter} {word}=> Not Found");
                        continue;
                    }

                    // If the text is found, create a destination that positions the found text
                    // at the upper window corner
                    if(results.Status == PdfTextSearchStatus.Found)
                    {
                        // 세부공정의 페이지번호 할당
                        page.PysicalNumber = results.PageNumber;
                        Console.WriteLine($"{page.Chapter} {word} => Found");

                        PdfXYZDestination destination = new PdfXYZDestination(results.Page, 0, results.Rectangles[0].Top, null);

                        // Create a bookmark associated with the destination
                        PdfBookmark bookmark = new PdfBookmark() { Title = page.Chapter, Destination = destination };
                        // Add the bookmark to the bookmark list
                        documentProcessor.Document.Bookmarks.Add(bookmark);

                        foreach (var childNode in node.Children)
                        {
                            var childPage = childNode.Value;
                            var childWord = childPage.Description;

                            Console.WriteLine($"{childPage.Chapter} {childWord}");

                            PdfTextSearchStatus finedTextStatus = PdfTextSearchStatus.NotFound;
                            PdfXYZDestination subDestination = null;
                            PdfTextSearchResults subResults = null;

                            // 1. 문자열 검색해서 찾으면 다음 문자열이 있는지 반복 검색
                            do
                            {
                                FindTextRecursive(documentProcessor, searchParameters, childPage, childWord, out subResults);

                                // 반복해서 문자열 검색(찾으면 종료)
                                if(subResults.Status == PdfTextSearchStatus.Found)
                                {
                                    subDestination = new PdfXYZDestination(subResults.Page, 0, subResults.Rectangles[0].Top, null);

                                    // 세부공정 페이지 이전에 검색된 내용은 제외한다.
                                    if(destination.PageIndex > subDestination.PageIndex)
                                        continue;

                                    // !HasBookmarks ? bookmark.Children.Add(childBookmark);
                                    if(documentProcessor.Document.Bookmarks.Where(p => p.Children.Count(b => b.Destination == subDestination) > 0).Count() == 0)
                                    {
                                        PdfBookmark childBookmark = new PdfBookmark() { Title = childPage.Chapter, Destination = subDestination };
                                        bookmark.Children.Add(childBookmark);
                                        finedTextStatus = PdfTextSearchStatus.Found;
                                        break;
                                    }

                                    finedTextStatus = PdfTextSearchStatus.Found;
                                }

                                if(subResults.Status == PdfTextSearchStatus.NotFound)
                                    break;

                            } while(subResults.Status != PdfTextSearchStatus.Finished);

                            if(finedTextStatus == PdfTextSearchStatus.Found)
                                continue;

                            // 2. 검색에 실패할 경우, 줄바꿈 문자로 한글/영문 분리해서 검색
                            string[] splits = new string[] { };
                            if(childWord.Contains("\n"))
                            {
                                splits = childWord.Split(new string[] { Environment.NewLine, "\n" }, StringSplitOptions.None);
                            }

                            foreach(var text in splits)
                            {
                                do
                                {
                                    FindTextRecursive(documentProcessor, searchParameters, childPage, text, out subResults);

                                    // 반복해서 문자열 검색(찾으면 종료)
                                    if(subResults.Status == PdfTextSearchStatus.Found)
                                    {
                                        Console.WriteLine($"2. SplitsTextSearh {childPage.Chapter} {childWord}=> Found");

                                        subDestination = new PdfXYZDestination(subResults.Page, 0, subResults.Rectangles[0].Top, null);

                                        // 세부공정 페이지 이전에 검색된 내용은 제외한다.
                                        if(destination.PageIndex > subDestination.PageIndex)
                                            continue;

                                        // !HasBookmarks ? bookmark.Children.Add(childBookmark);
                                        if(documentProcessor.Document.Bookmarks.Where(p => p.Children.Count(b => b.Destination == subDestination) > 0).Count() == 0)
                                        {
                                            PdfBookmark childBookmark = new PdfBookmark() { Title = childPage.Chapter, Destination = subDestination };
                                            bookmark.Children.Add(childBookmark);
                                            finedTextStatus = PdfTextSearchStatus.Found;
                                            break;
                                        }

                                        finedTextStatus = PdfTextSearchStatus.Found;
                                    }

                                    if(subResults.Status == PdfTextSearchStatus.NotFound)
                                        break;

                                } while(subResults.Status != PdfTextSearchStatus.Finished);

                                // 문자열을 나눠서 검색한다. 첫번재 Text 검색에 성공하면 나뉜 Text 검색은 생략한다. 
                                if(finedTextStatus == PdfTextSearchStatus.Found)
                                    break;
                            }

                            if(finedTextStatus == PdfTextSearchStatus.Found)
                                continue;

                            // 3. 내용만으로 검색에 실패할 경우, 특수한 케이스 대한 보정
                            // #3.1 Single Quote(') 가 포함된 문자열 검색에 실패하는 오류에 대한 예외처리
                            if(childWord.Contains("\'"))
                            {
                                childWord = childWord.Split(new string[] { @"'" }, StringSplitOptions.None).FirstOrDefault();

                                if(!string.IsNullOrWhiteSpace(childWord))
                                {
                                    var lastText = childWord.Substring(childWord.Length - 1, 1);
                                    // 마지막 문자가 공백이 아니라면 문자 + ' 가 결합된 형태
                                    if(!string.IsNullOrWhiteSpace(lastText))
                                    {
                                        // 마지막 단어 잘라내기
                                        int lastSpaceIndex = childWord.LastIndexOf(' ');
                                        childWord = childWord.Substring(0, lastSpaceIndex);
                                    }
                                }
                            }
                            // #3.2. word와 특수 문자가 결합된 형태의 문자는 검색에 오류를 발생시킨다.
                            if(childWord.Last() == ']')
                            {
                                int lastSpaceIndex = childWord.LastIndexOf(' ');
                                childWord = childWord.Substring(0, lastSpaceIndex);
                            }

                            do
                            {
                                FindTextRecursive(documentProcessor, searchParameters, childPage, childWord, out subResults);

                                // 반복해서 문자열 검색(찾으면 종료)
                                if(subResults.Status == PdfTextSearchStatus.Found)
                                {
                                    Console.WriteLine($"3. AdjustTextSearh {childPage.Chapter} {childWord}=> Found");

                                    subDestination = new PdfXYZDestination(subResults.Page, 0, subResults.Rectangles[0].Top, null);

                                    // 세부공정 페이지 이전에 검색된 내용은 제외한다.
                                    if(destination.PageIndex > subDestination.PageIndex)
                                        continue;

                                    // !HasBookmarks ? bookmark.Children.Add(childBookmark);
                                    if(documentProcessor.Document.Bookmarks.Where(p => p.Children.Count(b => b.Destination == subDestination) > 0).Count() == 0)
                                    {
                                        PdfBookmark childBookmark = new PdfBookmark() { Title = childPage.Chapter, Destination = subDestination };
                                        bookmark.Children.Add(childBookmark);
                                        finedTextStatus = PdfTextSearchStatus.Found;
                                        break;
                                    }

                                    finedTextStatus = PdfTextSearchStatus.Found;
                                }

                                if(subResults.Status == PdfTextSearchStatus.NotFound)
                                    break;

                            } while(subResults.Status != PdfTextSearchStatus.Finished);

                            if(finedTextStatus == PdfTextSearchStatus.Found)
                                continue;

                            Console.WriteLine($"#. Exception {childPage.Chapter} {childWord}=> Not Found");
                        }

                    }
                }

                // 테스트용 데이터
                const int INTERLEAVE_PAGE_COUNT = 0;
                const string USER_NAME = "박서준";
                string PRINT_DATE = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                const string PRINT_PURPOSE = "DEV 출력 검증용";
                const string ORIGINAL = "사본(COPY)";
                using(Font font = new Font("굴림", 10, FontStyle.Regular))
                {
                    using(SolidBrush brush = new SolidBrush(Color.Black))
                    {
                        int totalPageNO = documentProcessor.Document.Pages.Count - INTERLEAVE_PAGE_COUNT;

                        int pageNO = 1;
                        for(int i = 0; i < documentProcessor.Document.Pages.Count; ++i)
                        {
                            string pageString = string.Format("{0}/{1}", pageNO++, totalPageNO);
                            String pageWaterMark = String.Format("발행자 : {0} 발행일시 : {1} 발행사유 : {2}", USER_NAME, PRINT_DATE, PRINT_PURPOSE);
                            String pageInfo = $"문서상태 : {ORIGINAL}";
                            // Write PageNumber ( Page/Total )
                            using(PdfGraphics graph = documentProcessor.CreateGraphics())
                            {
                                if(pageString != string.Empty)
                                    graph.DrawString(pageString, font, brush, 380, 1100);

                                graph.DrawString(pageWaterMark, font, brush, 10, 10);
                                graph.DrawString(pageInfo, font, brush, 10, 1100);

                                graph.AddToPageForeground(documentProcessor.Document.Pages[i]);
                            }
                        }
                    }
                }


                using(MemoryStream ms = new MemoryStream())
                {
                    // Save the modified document
                    documentProcessor.SaveDocument(ms);
                    bookmarksBlob = ms.ToArray();
                }
            }

            return bookmarksBlob;
        }

        private PdfTextSearchResults FindTextRecursive(PdfDocumentProcessor documentProcessor, PdfTextSearchParameters searchParameters, Page childPage, string childWord, out PdfTextSearchResults subResults)
        {
            // 1. 전체 내용 검색
            subResults = documentProcessor.FindText(childWord, searchParameters);
            //PdfXYZDestination subDestination = null;
            if(subResults.Status == PdfTextSearchStatus.Found)
            {
                Console.WriteLine($"1. FullTextSearh {childPage.Chapter} => Found");

                return subResults;
                //subDestination = new PdfXYZDestination(subResults.Page, 0, subResults.Rectangles[0].Top, null);
            }

            

            return subResults;
        }

        #region LegacyMethod

        /// <summary>
        /// PDF 출력 전 마진 조정
        /// </summary>
        /// <param name="mbr_type">PROCESS, ETC</param>
        /// <param name="pWorkbook"></param>
        void fncDocumentSettingBeforeExport(string mbr_type, IWorkbook pWorkbook)
        {
            //PDF 변환 전 cell width, page print margin 조정

            // Access page margins.
            pWorkbook.BeginUpdate();

            // Access page margins.
            Margins pageMargins = pWorkbook.Worksheets[0].ActiveView.Margins;

            if(mbr_type == "PROCESS")
            {
                pWorkbook.Unit = DevExpress.Office.DocumentUnit.Document;

                pWorkbook.Worksheets.ActiveWorksheet.Columns["D"].Width = 881.25;
                pWorkbook.Worksheets.ActiveWorksheet.Columns["D"].WidthInCharacters = 39.619998931884766F;
                pWorkbook.Worksheets.ActiveWorksheet.PrintOptions.Scale = 100;
            }
            else
            {
                pWorkbook.Worksheets.ActiveWorksheet.PrintOptions.FitToPage = true;
                pWorkbook.Worksheets.ActiveWorksheet.PrintOptions.FitToWidth = 1;
                pWorkbook.Worksheets.ActiveWorksheet.PrintOptions.FitToHeight = 0;
                pWorkbook.Worksheets.ActiveWorksheet.PrintOptions.Scale = 100;
            }

            pWorkbook.Unit = DevExpress.Office.DocumentUnit.Inch;

            // Specify page margins.
            pageMargins.Left = 0.25F;
            pageMargins.Top = 0.60F;
            pageMargins.Right = 0.25F;
            pageMargins.Bottom = 0.60F;

            pageMargins.Header = 0.3F;
            pageMargins.Footer = 0.3F;

            pWorkbook.EndUpdate();
        }

        #endregion
    }
}
