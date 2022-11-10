using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.XtraBars.Helpers;
using DevExpress.XtraBars.Ribbon;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using static DevExpress.XtraExport.Helpers.TableRowControl;

namespace ExcelToSQL
{
    public partial class FrmMain : RibbonForm
    {
        string _currentFilePath = string.Empty;
        enum BUTTON
        {
            SOME_STATUS = 31,
            SOME_INFO = 32,
            FILE_OPEN = 62,
            CREATION_SCRIPT = 63,
            FILE_RELOAD = 64,
            CONVERT_SQL = 65,
        }
        public FrmMain()
        {
            InitializeComponent();
            InitSkinGallery();
            this.ribbonControl.OptionsTouch.ShowTouchUISelectorInQAT = true;

            this.ribbonControl.ItemClick += (s, e) =>
            {
                switch(e.Item.Id)
                {
                    case (int)BUTTON.FILE_OPEN:
                        FileOpen();
                        break;
                    case (int)BUTTON.CREATION_SCRIPT:
                        richEditControlScriptText.Text = CreationDataTable();
                        if (richEditControlScriptText.Text.Length > 0)
                            dockPanel1.ShowSliding();
                        break;
                    case (int)BUTTON.CONVERT_SQL:
                        richEditControlScriptText.Text = ConvertExecutesqlToQuery();
                        if(richEditControlScriptText.Text.Length > 0)
                        {
                            Clipboard.SetText(richEditControlScriptText.Text);
                            dockPanel1.ShowSliding();
                        }
                        break;

                    case (int)BUTTON.FILE_RELOAD:
                        if(string.IsNullOrEmpty(_currentFilePath))
                            break;
                        spreadsheetControl.LoadDocument(_currentFilePath);
                        break;
                    case (int)BUTTON.SOME_STATUS:
                        break;
                    case (int)BUTTON.SOME_INFO:
                        this.Capture = true;
                        break;
                    default:
                        DevExpress.XtraBars.Docking2010.Customization.FlyoutDialog.Show(this, e.Item.Id.ToString(), MessageBoxButtons.OK);
                        break;
                }

            };

            this.Load += (s, e) =>
            {
            };

            this.Activated += (s, e) =>
            {
            };

            this.Shown += (s, e) =>
            {

            };

            this.MouseMove += (s, e) =>
            {
                Control ctl = ControlHelper.FindControlAtPoint(this, Control.MousePosition);
                if(ctl == null)
                    return;

                this.siInfo.BeginUpdate();
                this.siInfo.Caption = ctl.Name;
                this.siInfo.EndUpdate();

            };

            this.MouseCaptureChanged += (s, e) =>
            {
                if(!string.IsNullOrEmpty(this.siInfo.Caption))
                    Clipboard.SetText(this.siInfo.Caption);
            };
        }
        void InitSkinGallery()
        {
            SkinHelper.InitSkinGallery(rgbiSkins, true);
        }

        void FileOpen()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "엑셀 파일 (*.xlsx)|*.xlsx|엑셀 파일 (*.xls)|*.xls";
            var result = openFileDialog.ShowDialog();

            if(result != DialogResult.OK)
                return;

            try
            {
                if(!File.Exists(openFileDialog.FileName))
                    throw new FileLoadException($"{openFileDialog.FileName} 경로 파일을 열 수 없습니다.");

                _currentFilePath = openFileDialog.FileName;

                spreadsheetControl.LoadDocument(_currentFilePath);
            }
            catch(Exception e)
            {
                DevExpress.XtraBars.Docking2010.Customization.FlyoutDialog.Show(this, e.Message, MessageBoxButtons.OK);
            }
            finally
            {

            }
        }

        void CreationScript()
        {
            const int HEADER_ROW_INDEX = 0;
            StringBuilder columnNames = new StringBuilder();
            StringBuilder rowStrings = new StringBuilder();
            var selectedSheet = spreadsheetControl.Document.Worksheets[spreadsheetControl.ActiveSheet.Name];
            Range usedRange = selectedSheet.GetUsedRange();
            for(int i = 0; i < usedRange.RowCount; i++)
            {
                if (HEADER_ROW_INDEX == i)
                {
                    foreach(Cell cell in usedRange[i])
                        columnNames.AppendLine(cell.Value.TextValue);
                    continue;
                }

                StringBuilder columnStrings = new StringBuilder();
                for(int j = 0; j < usedRange.ColumnCount; j++)
                {
                    Cell currentCell = usedRange[i, j];
                    // numeric values  
                    if(currentCell.Value.IsNumeric && !currentCell.Value.IsDateTime)
                    {
                        columnStrings.AppendLine(String.Format("Numeric Cell: {0}, Value: {1}\r\n", currentCell.GetReferenceA1(), currentCell.Value.NumericValue));
                    }
                    // date time values  
                    else if(currentCell.Value.IsDateTime)
                    {
                        columnStrings.AppendLine(String.Format("Date Time Cell: {0}, Value: {1}\r\n", currentCell.GetReferenceA1(), currentCell.Value.DateTimeValue));
                    }
                    // date time values  
                    else
                    {
                        columnStrings.AppendLine(String.Format("String Cell: {0}, Value: {1}\r\n", currentCell.GetReferenceA1(), currentCell.Value.TextValue));
                    }
                }


                rowStrings.AppendLine($"INSERT INTO {spreadsheetControl.ActiveSheet.Name}({rowStrings.ToString()}) VALUES { columnStrings.ToString()}");
            }

            richEditControlScriptText.Text = rowStrings.ToString();
        }

        /// <summary>
        /// Selection Range to DataTable and creation ad-hoc query 
        /// </summary>
        /// <returns></returns>
        string CreationDataTable()
        {
            var scriptText = string.Empty;

            try
            {
                Worksheet worksheet = spreadsheetControl.Document.Worksheets.ActiveWorksheet;
                Range range = worksheet.Selection;
                bool rangeHasHeaders = true;

                // Create a data table with column names obtained from the first row in a range if it has headers.
                // Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
                DataTable dataTable = worksheet.CreateDataTable(range, rangeHasHeaders);
                dataTable.TableName = worksheet.Name;

                const int DATA_ROW_INDEX = 1;
                //Validate cell value types. If cell value types in a column are different, the column values are exported as text.
                for(int col = 0; col < range.ColumnCount; col++)
                {
                    CellValueType cellType = range[DATA_ROW_INDEX, col].Value.Type;
                    for(int r = 1; r < range.RowCount; r++)
                    {
                        if(cellType != range[r, col].Value.Type)
                        {
                            dataTable.Columns[col].DataType = typeof(string);
                            break;
                        }
                    }
                }

                // Create the exporter that obtains data from the specified range, 
                // skips the header row (if required) and populates the previously created data table. 
                DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, rangeHasHeaders);
                // Handle value conversion errors.
                exporter.CellValueConversionError += (s,e) => {
                    MessageBox.Show("Error in cell " + e.Cell.GetReferenceA1());
                    e.DataTableValue = null;
                    e.Action = DataTableExporterAction.Continue;
                };

                // Perform the export.
                exporter.Export();

                // one Query
                //scriptText = SQLHelper.BuildInsertSQL(dataTable);
                // all Querys
                scriptText = SQLHelper.BuildInsertAdhocSQL(dataTable);

            }
            catch (Exception e)
            {
                DevExpress.XtraBars.Docking2010.Customization.FlyoutDialog.Show(this, e.Message, MessageBoxButtons.OK);
            }
            finally
            {

            }

            return scriptText;
        }

        string ConvertExecutesqlToQuery()
        {
            var scriptText = string.Empty;

            try
            {
                var originalText = string.Empty;
                if(richEditControlScriptText.Document.Selection.Length > 0)
                    originalText = richEditControlScriptText.Document.GetText(richEditControlScriptText.Document.Selection);

                if (string.IsNullOrEmpty(originalText))
                    originalText = Clipboard.GetText(TextDataFormat.Text);
                    

                scriptText = SQLHelper.ConvertExecuteSqlToNomarlSql(originalText);
            }
            catch (Exception e)
            {
                DevExpress.XtraBars.Docking2010.Customization.FlyoutDialog.Show(this, e.Message, MessageBoxButtons.OK);
            }
            finally
            {

            }

            return scriptText;
        }

    }
}