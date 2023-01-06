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

                SelectCommand(command);
            };

            this.dockPanel_pdfViewer.CustomButtonClick += (s, e) =>
            {
                ConvertWorkSheetToPDF();
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
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        //void RunScript(ICommand command)
        void SelectCommand(ICommand command)
        {
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

            }
        }

        void ConvertWorkSheetToPDF()
        {
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
                Dictionary<int, string> TOCs = new Dictionary<int, string>();

                //기존 이름 관리자(DefinedName) 삭제 후 재등록. 삭제시 공정, 세부공정 이름 관리자는 삭제하지 않는다.
                TreeNode<Page> pageNode = new TreeNode<Page>(new Page() { Description = "공정" });
                DefinedName[] definedNames = sheet.DefinedNames.ToArray();
                foreach(DefinedName definedName in definedNames)
                {
                    // 세부공정을 찾아서 첫번째 컬럼에 있는 번호가 순번이다.
                    if(definedName.Name.StartsWith("세부공정"))
                        TOCs.Add(definedName.Range.TopRowIndex
                            , usedRange[definedName.Range.TopRowIndex, SUBNUMBER_COLUMN_INDEX].Value.TextValue);
                }

                // DefinedName에 설정된 세부공정의 목록을 생성한다.
                TOCs.ToList().ForEach(t => {
                    Console.WriteLine($"{t.Key}=> {t.Value}");
                    pageNode.AddChild(new Page() { Index = t.Key, Description = t.Value});
                });


                const string PAGE_PREAKE_TEXT = "N/A";
                int cursorNO = 1;
                for(int i = 0; i < usedRange.RowCount; i++)
                {
                    // NOTE: { Cell: A21 formula:"" value: "6" type: Numeric}
                    Cell cell = usedRange[i, SUBNUMBER_COLUMN_INDEX];

                    if(cell.Value.Type == CellValueType.None)
                        continue;

                    int no = 0;

                    if(cell.Value.Type == CellValueType.Numeric)
                        no = ((int)cell.Value.NumericValue);

                    if (cell.Value.Type == CellValueType.Text)
                        int.TryParse(cell.Value.TextValue, out no);

                    var findedNode = pageNode.Children.Where(p => p.Value.Index < cell.RowIndex).LastOrDefault();
                    if(no > 0)
                    {
                        findedNode.AddChild(new Page()
                        {
                            Index = cell.RowIndex
                            , LogicalNumber = no
                            , PysicalNumber = cursorNO
                        });

                        Console.WriteLine($"IDX:{cell.RowIndex}, {findedNode.Value.Description}.{no} => Cursor:{cursorNO}, MAX:{findedNode.Children.Max(m=> m.Value.PysicalNumber)}");
                    }

                    if(cell.Value.TextValue == PAGE_PREAKE_TEXT)
                    {
                        //findedNode.Children.ToList().ForEach(p => p.Value.PysicalNumber = findedNode.Value.PysicalNumber);
                        Console.WriteLine($"IDX:{cell.RowIndex}, {findedNode.Value.Description} => Fined Page Break TEXT Logical Page NO => {findedNode.Children.Count}");
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

                // 3. PDF Load (Viewer 에서 PDF 전체를 보려면 파일로 저장하거나 메모리스트림을 유지해야 한다. using X)
                pdfViewer1.LoadDocument(new MemoryStream(bytes));
            }
            catch(Exception ex)
            {
                ExceptionTrigger?.Invoke(this, new FormExceptionEventArgs(ex.Message));
            }
            finally
            {

            }
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
