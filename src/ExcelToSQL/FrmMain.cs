using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.XtraBars.Helpers;
using DevExpress.XtraBars.Ribbon;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraBars.Docking2010.Views.WindowsUI;
using DevExpress.XtraBars.Docking2010.Customization;
using Presentation.Views.Upgrade;
using Presentation.Views.Excel;
using Presenter.Views.PDF;

namespace ExcelToSQL
{
    public partial class FrmMain : RibbonForm
    {
        const string CONNECTION_STRING = "";
        string _currentFilePath = string.Empty;
        enum BUTTON
        {
            SOME_STATUS = 31,
            SOME_INFO = 32,
            FILE_OPEN = 62,
            CREATION_SCRIPT = 63,
            FILE_RELOAD = 64,
            CONVERT_SQL = 65,
            LOAD_UPGRADE_FORM = 66,
            LOAD_ERBVIEWER_FORM = 67,
            LOAD_PDFVIEWER_FORM = 68,
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
                            dockPanel_script.ShowSliding();
                        break;
                    case (int)BUTTON.CONVERT_SQL:
                        richEditControlScriptText.Text = ConvertExecutesqlToQuery();
                        if(richEditControlScriptText.Text.Length > 0)
                        {
                            Clipboard.SetText(richEditControlScriptText.Text);
                            dockPanel_script.ShowSliding();
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
                    case (int)BUTTON.LOAD_UPGRADE_FORM:
                        Form form = new Form();
                        form.Controls.Add(new UpgradeForm() { Dock = DockStyle.Fill });
                        form.Show();
                        break;
                    case (int)BUTTON.LOAD_ERBVIEWER_FORM:
                        var uctrl = new EBRViewer() { Dock = DockStyle.Fill };
                        uctrl.ExceptionTrigger += (s1, e1)=> { MessageShowOK(e1.Message); };
                        Form formEBR = new Form();
                        formEBR.Size = new Size(1024, 768);
                        formEBR.Controls.Add(uctrl);
                        formEBR.Show();
                        break;
                    case (int)BUTTON.LOAD_PDFVIEWER_FORM:
                        var uctrlPdf = new Viewer() { Dock = DockStyle.Fill };
                        uctrlPdf.ExceptionTrigger += (s1, e1) => { MessageShowOK(e1.Message); };
                        Form formPDF = new Form();
                        formPDF.Size = new Size(1024, 768);
                        formPDF.Controls.Add(uctrlPdf);
                        formPDF.Show();
                        break;
                    default:
                        MessageShowOK(e.Item.Id.ToString());
                        break;
                }

            };

            this.Load += (s, e) =>
            {
                try
                {
                    SqlDependency.Start(CONNECTION_STRING);
                    this.GetData();
                }
                catch (Exception ex)
                {
                    MessageShowOK(ex.Message);
                }
                finally
                {

                }
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

            this.FormClosing += (s, e) =>
            {
                try
                { 
                    SqlDependency.Stop(CONNECTION_STRING);
                }
                catch (Exception ex)
                {
                    MessageShowOK(ex.Message);
                }
                finally
                {

                }
            };
        }

        void MessageShowOK(string message)
        {
            FlyoutAction action = new FlyoutAction() { Caption = "확인", Description = message };
            Predicate<DialogResult> predicate = canCloseFunc;
            FlyoutCommand command1 = new FlyoutCommand() { Text = "Close", Result = System.Windows.Forms.DialogResult.Yes };
            FlyoutCommand command2 = new FlyoutCommand() { Text = "Cancel", Result = System.Windows.Forms.DialogResult.No };
            action.Commands.Add(command1);
            action.Commands.Add(command2);
            FlyoutProperties properties = new FlyoutProperties();
            properties.ButtonSize = new Size(200, 40);
            properties.Style = FlyoutStyle.MessageBox;
            properties.AppearanceDescription.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;

            const int AUTO_CLOSE_INTERVAL = 10;
            // button 은 document usercontrol 로 구성
            var button = new Button() { Text = $"AUTO CLOSE ({AUTO_CLOSE_INTERVAL})" };
            button.Size = new Size(200, 40);

            Timer timer = new Timer();
            timer.Interval = 1000;

            FlyoutDialog flyoutDialog = new FlyoutDialog(this, "확인", button, MessageBoxButtons.OK);
            button.Click += (s, e) => {
                flyoutDialog.Close();
                timer.Stop();
            };


            int count = 0;
            int msec = 0;
            timer.Tick += (s, e) =>
            {
                msec += timer.Interval;
                button.Text = $"{message} ({AUTO_CLOSE_INTERVAL - count++})";
                if(msec > AUTO_CLOSE_INTERVAL * 1000)
                    flyoutDialog.Close();
            };
            timer.Start();

            flyoutDialog.Show(this);

            //var result = FlyoutDialog.Show(this, action, properties, predicate);
            //if(result != DialogResult.OK)
            //    return;





            //DevExpress.XtraBars.Docking2010.Customization.FlyoutDialog.Show(this, message, MessageBoxButtons.OK);
        }

        private bool canCloseFunc(DialogResult parameter)
        {
            return parameter != DialogResult.Cancel;
        }

        void InitSkinGallery()
        {
            SkinHelper.InitSkinGallery(rgbiSkins, true);
        }

        void GetData()
        {
            string sql = @"SELECT [MESSAGE_ID]
,[SENDER_ID]
,[RECEIVER_ID]
,[CONTENT]
,SEND_DT
,CONVERT(VARCHAR(10), SEND_DT, 120) as GROUP_DT
,CONFIRMED
,CONFIRM_DT
FROM[dbo].[SYS_MESSAGE]
ORDER BY SEND_DT DESC
";

            using(SqlConnection conn = new SqlConnection(CONNECTION_STRING))
            {
                conn.Open();
                var cmd = new SqlCommand(sql, conn);
                // (optional) 이전 Notification 삭제
                cmd.Notification = null;

                // SqlDependency객체 생성
                var sqlDependency = new SqlDependency(cmd);
                // SqlDependency.OnChange 이벤트 핸들러 지정
                sqlDependency.OnChange += new OnChangeEventHandler(Dependency_OnChange);

                // cmd 객체 SQL 실행, 데이타 Fetch
                SqlDataReader rdr = cmd.ExecuteReader();

                // 그리드에 결과 표시
                DataTable dt = new DataTable();
                dt.Load(rdr);
                gridControl1.DataSource = dt;
            }
        }

        private void Dependency_OnChange(object sender, SqlNotificationEventArgs e)
        {
            if(e.Info.ToString().StartsWith("Invalid")) // 에러
            {
                MessageShowOK(e.Info.ToString());
                return;
            }


            if(this.InvokeRequired)
            {
                this.Invoke((MethodInvoker)delegate 
                {
                    GetData();
                    dockPanel_messages.ShowSliding();
                });
            }
            else
            {
                GetData();
            }
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
                MessageShowOK(e.Message);
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
                MessageShowOK(e.Message);
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
                MessageShowOK(e.Message);
            }
            finally
            {

            }

            return scriptText;
        }

    }
}