namespace Presentation.Views.Excel
{
    partial class EBRViewer
    {
        /// <summary> 
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if(disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary> 
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            DevExpress.XtraBars.Docking.CustomHeaderButtonImageOptions customHeaderButtonImageOptions2 = new DevExpress.XtraBars.Docking.CustomHeaderButtonImageOptions();
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject2 = new DevExpress.Utils.SerializableAppearanceObject();
            DevExpress.XtraBars.Docking.CustomHeaderButtonImageOptions customHeaderButtonImageOptions1 = new DevExpress.XtraBars.Docking.CustomHeaderButtonImageOptions();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EBRViewer));
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject1 = new DevExpress.Utils.SerializableAppearanceObject();
            this.spreadsheetControl1 = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
            this.barManager1 = new DevExpress.XtraBars.BarManager(this.components);
            this.bar3 = new DevExpress.XtraBars.Bar();
            this.barHeaderItem1 = new DevExpress.XtraBars.BarHeaderItem();
            this.barStaticItem_performance = new DevExpress.XtraBars.BarStaticItem();
            this.barDockControlTop = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlBottom = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlLeft = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlRight = new DevExpress.XtraBars.BarDockControl();
            this.spreadsheetDockManager1 = new DevExpress.XtraSpreadsheet.SpreadsheetDockManager(this.components);
            this.hideContainerRight = new DevExpress.XtraBars.Docking.AutoHideContainer();
            this.dockPanel_pdfViewer = new DevExpress.XtraBars.Docking.DockPanel();
            this.controlContainer1 = new DevExpress.XtraBars.Docking.ControlContainer();
            this.pdfViewer1 = new DevExpress.XtraPdfViewer.PdfViewer();
            this.dockPanel_scripts = new DevExpress.XtraBars.Docking.DockPanel();
            this.dockPanel1_Container = new DevExpress.XtraBars.Docking.ControlContainer();
            this.memoEdit_sql = new DevExpress.XtraEditors.MemoEdit();
            this.dockPanel_excelList = new DevExpress.XtraBars.Docking.DockPanel();
            this.dockPanel2_Container = new DevExpress.XtraBars.Docking.ControlContainer();
            this.gridControl_excelList = new DevExpress.XtraGrid.GridControl();
            this.gridView_excelList = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.searchControl1 = new DevExpress.XtraEditors.SearchControl();
            this.barStaticItem_message = new DevExpress.XtraBars.BarStaticItem();
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.spreadsheetDockManager1)).BeginInit();
            this.hideContainerRight.SuspendLayout();
            this.dockPanel_pdfViewer.SuspendLayout();
            this.controlContainer1.SuspendLayout();
            this.dockPanel_scripts.SuspendLayout();
            this.dockPanel1_Container.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.memoEdit_sql.Properties)).BeginInit();
            this.dockPanel_excelList.SuspendLayout();
            this.dockPanel2_Container.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl_excelList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView_excelList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.searchControl1.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // spreadsheetControl1
            // 
            this.spreadsheetControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.spreadsheetControl1.Location = new System.Drawing.Point(200, 200);
            this.spreadsheetControl1.MenuManager = this.barManager1;
            this.spreadsheetControl1.Name = "spreadsheetControl1";
            this.spreadsheetControl1.Options.Import.Csv.Encoding = ((System.Text.Encoding)(resources.GetObject("spreadsheetControl1.Options.Import.Csv.Encoding")));
            this.spreadsheetControl1.Options.Import.Txt.Encoding = ((System.Text.Encoding)(resources.GetObject("spreadsheetControl1.Options.Import.Txt.Encoding")));
            this.spreadsheetControl1.Size = new System.Drawing.Size(580, 375);
            this.spreadsheetControl1.TabIndex = 0;
            this.spreadsheetControl1.Text = "spreadsheetControl1";
            // 
            // barManager1
            // 
            this.barManager1.Bars.AddRange(new DevExpress.XtraBars.Bar[] {
            this.bar3});
            this.barManager1.DockControls.Add(this.barDockControlTop);
            this.barManager1.DockControls.Add(this.barDockControlBottom);
            this.barManager1.DockControls.Add(this.barDockControlLeft);
            this.barManager1.DockControls.Add(this.barDockControlRight);
            this.barManager1.Form = this;
            this.barManager1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.barHeaderItem1,
            this.barStaticItem_performance,
            this.barStaticItem_message});
            this.barManager1.MaxItemId = 5;
            this.barManager1.StatusBar = this.bar3;
            // 
            // bar3
            // 
            this.bar3.BarName = "Status bar";
            this.bar3.CanDockStyle = DevExpress.XtraBars.BarCanDockStyle.Bottom;
            this.bar3.DockCol = 0;
            this.bar3.DockRow = 0;
            this.bar3.DockStyle = DevExpress.XtraBars.BarDockStyle.Bottom;
            this.bar3.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
            new DevExpress.XtraBars.LinkPersistInfo(this.barHeaderItem1),
            new DevExpress.XtraBars.LinkPersistInfo(this.barStaticItem_performance),
            new DevExpress.XtraBars.LinkPersistInfo(this.barStaticItem_message)});
            this.bar3.OptionsBar.AllowQuickCustomization = false;
            this.bar3.OptionsBar.DrawDragBorder = false;
            this.bar3.OptionsBar.UseWholeRow = true;
            this.bar3.Text = "Status bar";
            // 
            // barHeaderItem1
            // 
            this.barHeaderItem1.Caption = "STATUS";
            this.barHeaderItem1.Id = 2;
            this.barHeaderItem1.Name = "barHeaderItem1";
            // 
            // barStaticItem_performance
            // 
            this.barStaticItem_performance.Id = 3;
            this.barStaticItem_performance.Name = "barStaticItem_performance";
            // 
            // barDockControlTop
            // 
            this.barDockControlTop.CausesValidation = false;
            this.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.barDockControlTop.Location = new System.Drawing.Point(0, 0);
            this.barDockControlTop.Manager = this.barManager1;
            this.barDockControlTop.Size = new System.Drawing.Size(800, 0);
            // 
            // barDockControlBottom
            // 
            this.barDockControlBottom.CausesValidation = false;
            this.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.barDockControlBottom.Location = new System.Drawing.Point(0, 575);
            this.barDockControlBottom.Manager = this.barManager1;
            this.barDockControlBottom.Size = new System.Drawing.Size(800, 25);
            // 
            // barDockControlLeft
            // 
            this.barDockControlLeft.CausesValidation = false;
            this.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left;
            this.barDockControlLeft.Location = new System.Drawing.Point(0, 0);
            this.barDockControlLeft.Manager = this.barManager1;
            this.barDockControlLeft.Size = new System.Drawing.Size(0, 575);
            // 
            // barDockControlRight
            // 
            this.barDockControlRight.CausesValidation = false;
            this.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right;
            this.barDockControlRight.Location = new System.Drawing.Point(800, 0);
            this.barDockControlRight.Manager = this.barManager1;
            this.barDockControlRight.Size = new System.Drawing.Size(0, 575);
            // 
            // spreadsheetDockManager1
            // 
            this.spreadsheetDockManager1.AutoHideContainers.AddRange(new DevExpress.XtraBars.Docking.AutoHideContainer[] {
            this.hideContainerRight});
            this.spreadsheetDockManager1.Form = this;
            this.spreadsheetDockManager1.RootPanels.AddRange(new DevExpress.XtraBars.Docking.DockPanel[] {
            this.dockPanel_scripts,
            this.dockPanel_excelList});
            this.spreadsheetDockManager1.SpreadsheetControl = this.spreadsheetControl1;
            this.spreadsheetDockManager1.TopZIndexControls.AddRange(new string[] {
            "DevExpress.XtraBars.BarDockControl",
            "DevExpress.XtraBars.StandaloneBarDockControl",
            "System.Windows.Forms.StatusBar",
            "System.Windows.Forms.MenuStrip",
            "System.Windows.Forms.StatusStrip",
            "DevExpress.XtraBars.Ribbon.RibbonStatusBar",
            "DevExpress.XtraBars.Ribbon.RibbonControl",
            "DevExpress.XtraBars.Navigation.OfficeNavigationBar",
            "DevExpress.XtraBars.Navigation.TileNavPane",
            "DevExpress.XtraBars.TabFormControl"});
            // 
            // hideContainerRight
            // 
            this.hideContainerRight.BackColor = System.Drawing.SystemColors.Control;
            this.hideContainerRight.Controls.Add(this.dockPanel_pdfViewer);
            this.hideContainerRight.Dock = System.Windows.Forms.DockStyle.Right;
            this.hideContainerRight.Location = new System.Drawing.Point(780, 0);
            this.hideContainerRight.Name = "hideContainerRight";
            this.hideContainerRight.Size = new System.Drawing.Size(20, 575);
            // 
            // dockPanel_pdfViewer
            // 
            this.dockPanel_pdfViewer.Controls.Add(this.controlContainer1);
            customHeaderButtonImageOptions2.Image = ((System.Drawing.Image)(resources.GetObject("customHeaderButtonImageOptions2.Image")));
            this.dockPanel_pdfViewer.CustomHeaderButtons.AddRange(new DevExpress.XtraBars.Docking2010.IButton[] {
            new DevExpress.XtraBars.Docking.CustomHeaderButton("Preview(F9)", true, customHeaderButtonImageOptions2, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "", -1, true, null, true, false, true, serializableAppearanceObject2, null, -1)});
            this.dockPanel_pdfViewer.Dock = DevExpress.XtraBars.Docking.DockingStyle.Right;
            this.dockPanel_pdfViewer.ID = new System.Guid("4a35d2ed-3d65-48eb-b350-3f99c17ccdde");
            this.dockPanel_pdfViewer.Location = new System.Drawing.Point(0, 0);
            this.dockPanel_pdfViewer.Name = "dockPanel_pdfViewer";
            this.dockPanel_pdfViewer.OriginalSize = new System.Drawing.Size(434, 200);
            this.dockPanel_pdfViewer.SavedDock = DevExpress.XtraBars.Docking.DockingStyle.Right;
            this.dockPanel_pdfViewer.SavedIndex = 2;
            this.dockPanel_pdfViewer.Size = new System.Drawing.Size(434, 575);
            this.dockPanel_pdfViewer.Text = "PDF Viewer";
            this.dockPanel_pdfViewer.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide;
            // 
            // controlContainer1
            // 
            this.controlContainer1.Controls.Add(this.pdfViewer1);
            this.controlContainer1.Location = new System.Drawing.Point(5, 28);
            this.controlContainer1.Name = "controlContainer1";
            this.controlContainer1.Size = new System.Drawing.Size(425, 543);
            this.controlContainer1.TabIndex = 0;
            // 
            // pdfViewer1
            // 
            this.pdfViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pdfViewer1.Location = new System.Drawing.Point(0, 0);
            this.pdfViewer1.Name = "pdfViewer1";
            this.pdfViewer1.Size = new System.Drawing.Size(425, 543);
            this.pdfViewer1.TabIndex = 0;
            // 
            // dockPanel_scripts
            // 
            this.dockPanel_scripts.Controls.Add(this.dockPanel1_Container);
            customHeaderButtonImageOptions1.Image = ((System.Drawing.Image)(resources.GetObject("customHeaderButtonImageOptions1.Image")));
            this.dockPanel_scripts.CustomHeaderButtons.AddRange(new DevExpress.XtraBars.Docking2010.IButton[] {
            new DevExpress.XtraBars.Docking.CustomHeaderButton("Run(F5)", true, customHeaderButtonImageOptions1, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "", -1, true, null, true, false, true, serializableAppearanceObject1, null, -1)});
            this.dockPanel_scripts.Dock = DevExpress.XtraBars.Docking.DockingStyle.Top;
            this.dockPanel_scripts.ID = new System.Guid("4eb975c8-411d-4bea-bfb5-8d504f451eb1");
            this.dockPanel_scripts.Location = new System.Drawing.Point(0, 0);
            this.dockPanel_scripts.Name = "dockPanel_scripts";
            this.dockPanel_scripts.OriginalSize = new System.Drawing.Size(200, 200);
            this.dockPanel_scripts.Size = new System.Drawing.Size(780, 200);
            this.dockPanel_scripts.Text = "Scripts";
            // 
            // dockPanel1_Container
            // 
            this.dockPanel1_Container.Controls.Add(this.memoEdit_sql);
            this.dockPanel1_Container.Location = new System.Drawing.Point(4, 44);
            this.dockPanel1_Container.Name = "dockPanel1_Container";
            this.dockPanel1_Container.Size = new System.Drawing.Size(772, 151);
            this.dockPanel1_Container.TabIndex = 0;
            // 
            // memoEdit_sql
            // 
            this.memoEdit_sql.Dock = System.Windows.Forms.DockStyle.Fill;
            this.memoEdit_sql.Location = new System.Drawing.Point(0, 0);
            this.memoEdit_sql.Name = "memoEdit_sql";
            this.memoEdit_sql.Size = new System.Drawing.Size(772, 151);
            this.memoEdit_sql.TabIndex = 0;
            // 
            // dockPanel_excelList
            // 
            this.dockPanel_excelList.Controls.Add(this.dockPanel2_Container);
            this.dockPanel_excelList.Dock = DevExpress.XtraBars.Docking.DockingStyle.Left;
            this.dockPanel_excelList.ID = new System.Guid("28fbc935-6904-421a-bc36-21811df83639");
            this.dockPanel_excelList.Location = new System.Drawing.Point(0, 200);
            this.dockPanel_excelList.Name = "dockPanel_excelList";
            this.dockPanel_excelList.OriginalSize = new System.Drawing.Size(200, 200);
            this.dockPanel_excelList.Size = new System.Drawing.Size(200, 375);
            this.dockPanel_excelList.Text = "Excel List";
            // 
            // dockPanel2_Container
            // 
            this.dockPanel2_Container.Controls.Add(this.gridControl_excelList);
            this.dockPanel2_Container.Controls.Add(this.searchControl1);
            this.dockPanel2_Container.Location = new System.Drawing.Point(4, 23);
            this.dockPanel2_Container.Name = "dockPanel2_Container";
            this.dockPanel2_Container.Size = new System.Drawing.Size(191, 348);
            this.dockPanel2_Container.TabIndex = 0;
            // 
            // gridControl_excelList
            // 
            this.gridControl_excelList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridControl_excelList.Location = new System.Drawing.Point(0, 20);
            this.gridControl_excelList.MainView = this.gridView_excelList;
            this.gridControl_excelList.Name = "gridControl_excelList";
            this.gridControl_excelList.Size = new System.Drawing.Size(191, 328);
            this.gridControl_excelList.TabIndex = 0;
            this.gridControl_excelList.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView_excelList});
            // 
            // gridView_excelList
            // 
            this.gridView_excelList.GridControl = this.gridControl_excelList;
            this.gridView_excelList.Name = "gridView_excelList";
            this.gridView_excelList.OptionsView.ShowGroupPanel = false;
            // 
            // searchControl1
            // 
            this.searchControl1.Client = this.gridControl_excelList;
            this.searchControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.searchControl1.Location = new System.Drawing.Point(0, 0);
            this.searchControl1.Name = "searchControl1";
            this.searchControl1.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Repository.ClearButton(),
            new DevExpress.XtraEditors.Repository.SearchButton()});
            this.searchControl1.Properties.Client = this.gridControl_excelList;
            this.searchControl1.Size = new System.Drawing.Size(191, 20);
            this.searchControl1.TabIndex = 1;
            // 
            // barStaticItem_message
            // 
            this.barStaticItem_message.Id = 4;
            this.barStaticItem_message.Name = "barStaticItem_message";
            // 
            // EBRViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.spreadsheetControl1);
            this.Controls.Add(this.dockPanel_excelList);
            this.Controls.Add(this.dockPanel_scripts);
            this.Controls.Add(this.hideContainerRight);
            this.Controls.Add(this.barDockControlLeft);
            this.Controls.Add(this.barDockControlRight);
            this.Controls.Add(this.barDockControlBottom);
            this.Controls.Add(this.barDockControlTop);
            this.Name = "EBRViewer";
            this.Size = new System.Drawing.Size(800, 600);
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.spreadsheetDockManager1)).EndInit();
            this.hideContainerRight.ResumeLayout(false);
            this.dockPanel_pdfViewer.ResumeLayout(false);
            this.controlContainer1.ResumeLayout(false);
            this.dockPanel_scripts.ResumeLayout(false);
            this.dockPanel1_Container.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.memoEdit_sql.Properties)).EndInit();
            this.dockPanel_excelList.ResumeLayout(false);
            this.dockPanel2_Container.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridControl_excelList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView_excelList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.searchControl1.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraSpreadsheet.SpreadsheetControl spreadsheetControl1;
        private DevExpress.XtraSpreadsheet.SpreadsheetDockManager spreadsheetDockManager1;
        private DevExpress.XtraBars.Docking.DockPanel dockPanel_excelList;
        private DevExpress.XtraBars.Docking.ControlContainer dockPanel2_Container;
        private DevExpress.XtraGrid.GridControl gridControl_excelList;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView_excelList;
        private DevExpress.XtraEditors.SearchControl searchControl1;
        private DevExpress.XtraBars.Docking.DockPanel dockPanel_scripts;
        private DevExpress.XtraBars.Docking.ControlContainer dockPanel1_Container;
        private DevExpress.XtraEditors.MemoEdit memoEdit_sql;
        private DevExpress.XtraBars.Docking.AutoHideContainer hideContainerRight;
        private DevExpress.XtraBars.Docking.DockPanel dockPanel_pdfViewer;
        private DevExpress.XtraBars.Docking.ControlContainer controlContainer1;
        private DevExpress.XtraPdfViewer.PdfViewer pdfViewer1;
        private DevExpress.XtraBars.BarManager barManager1;
        private DevExpress.XtraBars.Bar bar3;
        private DevExpress.XtraBars.BarHeaderItem barHeaderItem1;
        private DevExpress.XtraBars.BarStaticItem barStaticItem_performance;
        private DevExpress.XtraBars.BarDockControl barDockControlTop;
        private DevExpress.XtraBars.BarDockControl barDockControlBottom;
        private DevExpress.XtraBars.BarDockControl barDockControlLeft;
        private DevExpress.XtraBars.BarDockControl barDockControlRight;
        private DevExpress.XtraBars.BarStaticItem barStaticItem_message;
    }
}
