namespace ExcelToSQL
{
    partial class FrmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if(disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            this.ribbonControl = new DevExpress.XtraBars.Ribbon.RibbonControl();
            this.appMenu = new DevExpress.XtraBars.Ribbon.ApplicationMenu(this.components);
            this.popupControlContainer2 = new DevExpress.XtraBars.PopupControlContainer(this.components);
            this.buttonEdit = new DevExpress.XtraEditors.ButtonEdit();
            this.iExit = new DevExpress.XtraBars.BarButtonItem();
            this.popupControlContainer1 = new DevExpress.XtraBars.PopupControlContainer(this.components);
            this.someLabelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.someLabelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.ribbonImageCollection = new DevExpress.Utils.ImageCollection(this.components);
            this.iHelp = new DevExpress.XtraBars.BarButtonItem();
            this.iAbout = new DevExpress.XtraBars.BarButtonItem();
            this.siStatus = new DevExpress.XtraBars.BarStaticItem();
            this.siInfo = new DevExpress.XtraBars.BarStaticItem();
            this.rgbiSkins = new DevExpress.XtraBars.RibbonGalleryBarItem();
            this.iOpenFile = new DevExpress.XtraBars.BarButtonItem();
            this.ribbonImageCollectionLarge = new DevExpress.Utils.ImageCollection(this.components);
            this.ribbonPageSkins = new DevExpress.XtraBars.Ribbon.RibbonPage();
            this.skinsRibbonPageGroup = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
            this.helpRibbonPage = new DevExpress.XtraBars.Ribbon.RibbonPage();
            this.helpRibbonPageGroup = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
            this.FileRibbonPage = new DevExpress.XtraBars.Ribbon.RibbonPage();
            this.fileRibbonPageGroup = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
            this.ribbonStatusBar = new DevExpress.XtraBars.Ribbon.RibbonStatusBar();
            this.spreadsheetFormulaBarPanel = new System.Windows.Forms.Panel();
            this.spreadsheetControl = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
            this.splitterControl = new DevExpress.XtraEditors.SplitterControl();
            this.formulaBarNameBoxSplitContainerControl = new DevExpress.XtraEditors.SplitContainerControl();
            this.spreadsheetNameBoxControl = new DevExpress.XtraSpreadsheet.SpreadsheetNameBoxControl();
            this.spreadsheetFormulaBarControl1 = new DevExpress.XtraSpreadsheet.SpreadsheetFormulaBarControl();
            this.ScriptRibbonPageGroup = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
            this.iCreationScript = new DevExpress.XtraBars.BarButtonItem();
            this.richEditControlScriptText = new DevExpress.XtraRichEdit.RichEditControl();
            this.dockManager1 = new DevExpress.XtraBars.Docking.DockManager(this.components);
            this.dockPanel1 = new DevExpress.XtraBars.Docking.DockPanel();
            this.dockPanel1_Container = new DevExpress.XtraBars.Docking.ControlContainer();
            ((System.ComponentModel.ISupportInitialize)(this.ribbonControl)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.appMenu)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.popupControlContainer2)).BeginInit();
            this.popupControlContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.buttonEdit.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.popupControlContainer1)).BeginInit();
            this.popupControlContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ribbonImageCollection)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ribbonImageCollectionLarge)).BeginInit();
            this.spreadsheetFormulaBarPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.formulaBarNameBoxSplitContainerControl)).BeginInit();
            this.formulaBarNameBoxSplitContainerControl.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.spreadsheetNameBoxControl.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dockManager1)).BeginInit();
            this.dockPanel1.SuspendLayout();
            this.dockPanel1_Container.SuspendLayout();
            this.SuspendLayout();
            // 
            // ribbonControl
            // 
            this.ribbonControl.ApplicationButtonDropDownControl = this.appMenu;
            this.ribbonControl.ApplicationButtonText = null;
            this.ribbonControl.ExpandCollapseItem.Id = 0;
            this.ribbonControl.Images = this.ribbonImageCollection;
            this.ribbonControl.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.ribbonControl.ExpandCollapseItem,
            this.iExit,
            this.iHelp,
            this.iAbout,
            this.siStatus,
            this.siInfo,
            this.rgbiSkins,
            this.iOpenFile,
            this.iCreationScript});
            this.ribbonControl.LargeImages = this.ribbonImageCollectionLarge;
            this.ribbonControl.Location = new System.Drawing.Point(0, 0);
            this.ribbonControl.MaxItemId = 64;
            this.ribbonControl.Name = "ribbonControl";
            this.ribbonControl.PageHeaderItemLinks.Add(this.iAbout);
            this.ribbonControl.Pages.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPage[] {
            this.ribbonPageSkins,
            this.helpRibbonPage,
            this.FileRibbonPage});
            this.ribbonControl.QuickToolbarItemLinks.Add(this.iHelp);
            this.ribbonControl.RibbonStyle = DevExpress.XtraBars.Ribbon.RibbonControlStyle.Office2010;
            this.ribbonControl.Size = new System.Drawing.Size(1100, 147);
            this.ribbonControl.StatusBar = this.ribbonStatusBar;
            // 
            // appMenu
            // 
            this.appMenu.BottomPaneControlContainer = this.popupControlContainer2;
            this.appMenu.ItemLinks.Add(this.iExit);
            this.appMenu.Name = "appMenu";
            this.appMenu.Ribbon = this.ribbonControl;
            this.appMenu.RightPaneControlContainer = this.popupControlContainer1;
            this.appMenu.ShowRightPane = true;
            // 
            // popupControlContainer2
            // 
            this.popupControlContainer2.Appearance.BackColor = System.Drawing.Color.Transparent;
            this.popupControlContainer2.Appearance.Options.UseBackColor = true;
            this.popupControlContainer2.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.popupControlContainer2.Controls.Add(this.buttonEdit);
            this.popupControlContainer2.Location = new System.Drawing.Point(238, 289);
            this.popupControlContainer2.Name = "popupControlContainer2";
            this.popupControlContainer2.Ribbon = this.ribbonControl;
            this.popupControlContainer2.Size = new System.Drawing.Size(118, 28);
            this.popupControlContainer2.TabIndex = 3;
            this.popupControlContainer2.Visible = false;
            // 
            // buttonEdit
            // 
            this.buttonEdit.EditValue = "Some Text";
            this.buttonEdit.Location = new System.Drawing.Point(3, 5);
            this.buttonEdit.MenuManager = this.ribbonControl;
            this.buttonEdit.Name = "buttonEdit";
            this.buttonEdit.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
            this.buttonEdit.Size = new System.Drawing.Size(100, 20);
            this.buttonEdit.TabIndex = 0;
            // 
            // iExit
            // 
            this.iExit.Caption = "Exit";
            this.iExit.Description = "Closes this program after prompting you to save unsaved data.";
            this.iExit.Hint = "Closes this program after prompting you to save unsaved data";
            this.iExit.Id = 20;
            this.iExit.ImageOptions.ImageIndex = 6;
            this.iExit.ImageOptions.LargeImageIndex = 6;
            this.iExit.Name = "iExit";
            // 
            // popupControlContainer1
            // 
            this.popupControlContainer1.Appearance.BackColor = System.Drawing.Color.Transparent;
            this.popupControlContainer1.Appearance.Options.UseBackColor = true;
            this.popupControlContainer1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.popupControlContainer1.Controls.Add(this.someLabelControl2);
            this.popupControlContainer1.Controls.Add(this.someLabelControl1);
            this.popupControlContainer1.Location = new System.Drawing.Point(111, 197);
            this.popupControlContainer1.Name = "popupControlContainer1";
            this.popupControlContainer1.Ribbon = this.ribbonControl;
            this.popupControlContainer1.Size = new System.Drawing.Size(76, 70);
            this.popupControlContainer1.TabIndex = 2;
            this.popupControlContainer1.Visible = false;
            // 
            // someLabelControl2
            // 
            this.someLabelControl2.Location = new System.Drawing.Point(3, 57);
            this.someLabelControl2.Name = "someLabelControl2";
            this.someLabelControl2.Size = new System.Drawing.Size(57, 14);
            this.someLabelControl2.TabIndex = 0;
            this.someLabelControl2.Text = "Some Info";
            // 
            // someLabelControl1
            // 
            this.someLabelControl1.Location = new System.Drawing.Point(3, 3);
            this.someLabelControl1.Name = "someLabelControl1";
            this.someLabelControl1.Size = new System.Drawing.Size(57, 14);
            this.someLabelControl1.TabIndex = 0;
            this.someLabelControl1.Text = "Some Info";
            // 
            // ribbonImageCollection
            // 
            this.ribbonImageCollection.ImageStream = ((DevExpress.Utils.ImageCollectionStreamer)(resources.GetObject("ribbonImageCollection.ImageStream")));
            this.ribbonImageCollection.Images.SetKeyName(0, "Ribbon_Exit_16x16.png");
            this.ribbonImageCollection.Images.SetKeyName(1, "Ribbon_Content_16x16.png");
            this.ribbonImageCollection.Images.SetKeyName(2, "Ribbon_Info_16x16.png");
            // 
            // iHelp
            // 
            this.iHelp.Caption = "Help";
            this.iHelp.Description = "Start the program help system.";
            this.iHelp.Hint = "Start the program help system";
            this.iHelp.Id = 22;
            this.iHelp.ImageOptions.ImageIndex = 7;
            this.iHelp.ImageOptions.LargeImageIndex = 7;
            this.iHelp.Name = "iHelp";
            // 
            // iAbout
            // 
            this.iAbout.Caption = "About";
            this.iAbout.Description = "Displays general program information.";
            this.iAbout.Hint = "Displays general program information";
            this.iAbout.Id = 24;
            this.iAbout.ImageOptions.ImageIndex = 8;
            this.iAbout.ImageOptions.LargeImageIndex = 8;
            this.iAbout.Name = "iAbout";
            // 
            // siStatus
            // 
            this.siStatus.Caption = "Some Status Info";
            this.siStatus.Id = 31;
            this.siStatus.Name = "siStatus";
            // 
            // siInfo
            // 
            this.siInfo.Caption = "Some Info";
            this.siInfo.Id = 32;
            this.siInfo.Name = "siInfo";
            // 
            // rgbiSkins
            // 
            this.rgbiSkins.Caption = "Skins";
            // 
            // 
            // 
            this.rgbiSkins.Gallery.AllowHoverImages = true;
            this.rgbiSkins.Gallery.Appearance.ItemCaptionAppearance.Normal.Options.UseFont = true;
            this.rgbiSkins.Gallery.Appearance.ItemCaptionAppearance.Normal.Options.UseTextOptions = true;
            this.rgbiSkins.Gallery.Appearance.ItemCaptionAppearance.Normal.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.rgbiSkins.Gallery.ColumnCount = 4;
            this.rgbiSkins.Gallery.FixedHoverImageSize = false;
            this.rgbiSkins.Gallery.ImageSize = new System.Drawing.Size(32, 17);
            this.rgbiSkins.Gallery.ItemImageLocation = DevExpress.Utils.Locations.Top;
            this.rgbiSkins.Gallery.RowCount = 4;
            this.rgbiSkins.Id = 60;
            this.rgbiSkins.Name = "rgbiSkins";
            // 
            // iOpenFile
            // 
            this.iOpenFile.Caption = "File Open";
            this.iOpenFile.Description = "Open Excel File";
            this.iOpenFile.Hint = "Open Excel File";
            this.iOpenFile.Id = 62;
            this.iOpenFile.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("iOpenFile.ImageOptions.Image")));
            this.iOpenFile.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("iOpenFile.ImageOptions.LargeImage")));
            this.iOpenFile.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O));
            this.iOpenFile.Name = "iOpenFile";
            this.iOpenFile.ShortcutKeyDisplayString = "Ctrl+O";
            // 
            // ribbonImageCollectionLarge
            // 
            this.ribbonImageCollectionLarge.ImageSize = new System.Drawing.Size(32, 32);
            this.ribbonImageCollectionLarge.ImageStream = ((DevExpress.Utils.ImageCollectionStreamer)(resources.GetObject("ribbonImageCollectionLarge.ImageStream")));
            this.ribbonImageCollectionLarge.Images.SetKeyName(0, "Ribbon_Exit_32x32.png");
            this.ribbonImageCollectionLarge.Images.SetKeyName(1, "Ribbon_Content_32x32.png");
            this.ribbonImageCollectionLarge.Images.SetKeyName(2, "Ribbon_Info_32x32.png");
            // 
            // ribbonPageSkins
            // 
            this.ribbonPageSkins.Groups.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageGroup[] {
            this.skinsRibbonPageGroup});
            this.ribbonPageSkins.Name = "ribbonPageSkins";
            this.ribbonPageSkins.Text = "Skins";
            // 
            // skinsRibbonPageGroup
            // 
            this.skinsRibbonPageGroup.ItemLinks.Add(this.rgbiSkins);
            this.skinsRibbonPageGroup.Name = "skinsRibbonPageGroup";
            this.skinsRibbonPageGroup.ShowCaptionButton = false;
            this.skinsRibbonPageGroup.Text = "Skins";
            // 
            // helpRibbonPage
            // 
            this.helpRibbonPage.Groups.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageGroup[] {
            this.helpRibbonPageGroup});
            this.helpRibbonPage.Name = "helpRibbonPage";
            this.helpRibbonPage.Text = "Help";
            // 
            // helpRibbonPageGroup
            // 
            this.helpRibbonPageGroup.ItemLinks.Add(this.iHelp);
            this.helpRibbonPageGroup.ItemLinks.Add(this.iAbout);
            this.helpRibbonPageGroup.Name = "helpRibbonPageGroup";
            this.helpRibbonPageGroup.Text = "Help";
            // 
            // FileRibbonPage
            // 
            this.FileRibbonPage.Groups.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageGroup[] {
            this.fileRibbonPageGroup,
            this.ScriptRibbonPageGroup});
            this.FileRibbonPage.Name = "FileRibbonPage";
            this.FileRibbonPage.Text = "File";
            // 
            // fileRibbonPageGroup
            // 
            this.fileRibbonPageGroup.ItemLinks.Add(this.iOpenFile);
            this.fileRibbonPageGroup.Name = "fileRibbonPageGroup";
            this.fileRibbonPageGroup.Text = "File";
            // 
            // ribbonStatusBar
            // 
            this.ribbonStatusBar.ItemLinks.Add(this.siStatus);
            this.ribbonStatusBar.ItemLinks.Add(this.siInfo);
            this.ribbonStatusBar.Location = new System.Drawing.Point(0, 669);
            this.ribbonStatusBar.Name = "ribbonStatusBar";
            this.ribbonStatusBar.Ribbon = this.ribbonControl;
            this.ribbonStatusBar.Size = new System.Drawing.Size(1100, 31);
            // 
            // spreadsheetFormulaBarPanel
            // 
            this.spreadsheetFormulaBarPanel.Controls.Add(this.spreadsheetControl);
            this.spreadsheetFormulaBarPanel.Controls.Add(this.splitterControl);
            this.spreadsheetFormulaBarPanel.Controls.Add(this.formulaBarNameBoxSplitContainerControl);
            this.spreadsheetFormulaBarPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.spreadsheetFormulaBarPanel.Location = new System.Drawing.Point(0, 147);
            this.spreadsheetFormulaBarPanel.Name = "spreadsheetFormulaBarPanel";
            this.spreadsheetFormulaBarPanel.Size = new System.Drawing.Size(633, 522);
            this.spreadsheetFormulaBarPanel.TabIndex = 3;
            // 
            // spreadsheetControl
            // 
            this.spreadsheetControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.spreadsheetControl.Location = new System.Drawing.Point(0, 27);
            this.spreadsheetControl.MenuManager = this.ribbonControl;
            this.spreadsheetControl.Name = "spreadsheetControl";
            this.spreadsheetControl.Options.Import.Csv.Encoding = ((System.Text.Encoding)(resources.GetObject("spreadsheetControl.Options.Import.Csv.Encoding")));
            this.spreadsheetControl.Options.Import.Txt.Encoding = ((System.Text.Encoding)(resources.GetObject("spreadsheetControl.Options.Import.Txt.Encoding")));
            this.spreadsheetControl.Size = new System.Drawing.Size(633, 495);
            this.spreadsheetControl.TabIndex = 1;
            this.spreadsheetControl.Text = "spreadsheetControl1";
            // 
            // splitterControl
            // 
            this.splitterControl.Dock = System.Windows.Forms.DockStyle.Top;
            this.splitterControl.Location = new System.Drawing.Point(0, 22);
            this.splitterControl.MinSize = 20;
            this.splitterControl.Name = "splitterControl";
            this.splitterControl.Size = new System.Drawing.Size(633, 5);
            this.splitterControl.TabIndex = 2;
            this.splitterControl.TabStop = false;
            // 
            // formulaBarNameBoxSplitContainerControl
            // 
            this.formulaBarNameBoxSplitContainerControl.Dock = System.Windows.Forms.DockStyle.Top;
            this.formulaBarNameBoxSplitContainerControl.Location = new System.Drawing.Point(0, 0);
            this.formulaBarNameBoxSplitContainerControl.MinimumSize = new System.Drawing.Size(0, 22);
            this.formulaBarNameBoxSplitContainerControl.Name = "formulaBarNameBoxSplitContainerControl";
            this.formulaBarNameBoxSplitContainerControl.Panel1.Controls.Add(this.spreadsheetNameBoxControl);
            this.formulaBarNameBoxSplitContainerControl.Panel2.Controls.Add(this.spreadsheetFormulaBarControl1);
            this.formulaBarNameBoxSplitContainerControl.Size = new System.Drawing.Size(633, 22);
            this.formulaBarNameBoxSplitContainerControl.SplitterPosition = 145;
            this.formulaBarNameBoxSplitContainerControl.TabIndex = 3;
            // 
            // spreadsheetNameBoxControl
            // 
            this.spreadsheetNameBoxControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.spreadsheetNameBoxControl.EditValue = "A1";
            this.spreadsheetNameBoxControl.Location = new System.Drawing.Point(0, 0);
            this.spreadsheetNameBoxControl.Name = "spreadsheetNameBoxControl";
            this.spreadsheetNameBoxControl.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.spreadsheetNameBoxControl.Size = new System.Drawing.Size(145, 20);
            this.spreadsheetNameBoxControl.SpreadsheetControl = this.spreadsheetControl;
            this.spreadsheetNameBoxControl.TabIndex = 0;
            // 
            // spreadsheetFormulaBarControl1
            // 
            this.spreadsheetFormulaBarControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.spreadsheetFormulaBarControl1.Location = new System.Drawing.Point(0, 0);
            this.spreadsheetFormulaBarControl1.MinimumSize = new System.Drawing.Size(0, 20);
            this.spreadsheetFormulaBarControl1.Name = "spreadsheetFormulaBarControl1";
            this.spreadsheetFormulaBarControl1.Size = new System.Drawing.Size(483, 22);
            this.spreadsheetFormulaBarControl1.SpreadsheetControl = this.spreadsheetControl;
            this.spreadsheetFormulaBarControl1.TabIndex = 0;
            // 
            // ScriptRibbonPageGroup
            // 
            this.ScriptRibbonPageGroup.ItemLinks.Add(this.iCreationScript);
            this.ScriptRibbonPageGroup.Name = "ScriptRibbonPageGroup";
            this.ScriptRibbonPageGroup.Text = "Script";
            // 
            // iCreationScript
            // 
            this.iCreationScript.Caption = "SQL Creation";
            this.iCreationScript.Description = "Creation SQL Script.";
            this.iCreationScript.Hint = "Creation SQL Script.";
            this.iCreationScript.Id = 63;
            this.iCreationScript.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("iCreationScript.ImageOptions.Image")));
            this.iCreationScript.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("iCreationScript.ImageOptions.LargeImage")));
            this.iCreationScript.ItemShortcut = new DevExpress.XtraBars.BarShortcut(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
                | System.Windows.Forms.Keys.C));
            this.iCreationScript.Name = "iCreationScript";
            this.iCreationScript.ShortcutKeyDisplayString = "Ctrl+Shift+C";
            this.iCreationScript.ShowItemShortcut = DevExpress.Utils.DefaultBoolean.True;
            // 
            // richEditControlScriptText
            // 
            this.richEditControlScriptText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richEditControlScriptText.Location = new System.Drawing.Point(0, 0);
            this.richEditControlScriptText.MenuManager = this.ribbonControl;
            this.richEditControlScriptText.Name = "richEditControlScriptText";
            this.richEditControlScriptText.Size = new System.Drawing.Size(458, 495);
            this.richEditControlScriptText.TabIndex = 4;
            // 
            // dockManager1
            // 
            this.dockManager1.Form = this;
            this.dockManager1.RootPanels.AddRange(new DevExpress.XtraBars.Docking.DockPanel[] {
            this.dockPanel1});
            this.dockManager1.TopZIndexControls.AddRange(new string[] {
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
            // dockPanel1
            // 
            this.dockPanel1.Controls.Add(this.dockPanel1_Container);
            this.dockPanel1.Dock = DevExpress.XtraBars.Docking.DockingStyle.Right;
            this.dockPanel1.ID = new System.Guid("cbcac618-c473-4e34-b233-920f012d70b9");
            this.dockPanel1.Location = new System.Drawing.Point(633, 147);
            this.dockPanel1.Name = "dockPanel1";
            this.dockPanel1.OriginalSize = new System.Drawing.Size(467, 200);
            this.dockPanel1.Size = new System.Drawing.Size(467, 522);
            this.dockPanel1.Text = "Scripts";
            // 
            // dockPanel1_Container
            // 
            this.dockPanel1_Container.Controls.Add(this.richEditControlScriptText);
            this.dockPanel1_Container.Location = new System.Drawing.Point(5, 23);
            this.dockPanel1_Container.Name = "dockPanel1_Container";
            this.dockPanel1_Container.Size = new System.Drawing.Size(458, 495);
            this.dockPanel1_Container.TabIndex = 0;
            // 
            // Form1
            // 
            this.AllowFormGlass = DevExpress.Utils.DefaultBoolean.False;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1100, 700);
            this.Controls.Add(this.spreadsheetFormulaBarPanel);
            this.Controls.Add(this.popupControlContainer1);
            this.Controls.Add(this.popupControlContainer2);
            this.Controls.Add(this.dockPanel1);
            this.Controls.Add(this.ribbonStatusBar);
            this.Controls.Add(this.ribbonControl);
            this.Name = "Form1";
            this.Ribbon = this.ribbonControl;
            this.StatusBar = this.ribbonStatusBar;
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.ribbonControl)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.appMenu)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.popupControlContainer2)).EndInit();
            this.popupControlContainer2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.buttonEdit.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.popupControlContainer1)).EndInit();
            this.popupControlContainer1.ResumeLayout(false);
            this.popupControlContainer1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ribbonImageCollection)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ribbonImageCollectionLarge)).EndInit();
            this.spreadsheetFormulaBarPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.formulaBarNameBoxSplitContainerControl)).EndInit();
            this.formulaBarNameBoxSplitContainerControl.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.spreadsheetNameBoxControl.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dockManager1)).EndInit();
            this.dockPanel1.ResumeLayout(false);
            this.dockPanel1_Container.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraBars.Ribbon.RibbonControl ribbonControl;
        private DevExpress.XtraBars.BarButtonItem iExit;
        private DevExpress.XtraBars.BarButtonItem iHelp;
        private DevExpress.XtraBars.BarButtonItem iAbout;
        private DevExpress.XtraBars.BarStaticItem siStatus;
        private DevExpress.XtraBars.BarStaticItem siInfo;
        private DevExpress.XtraBars.RibbonGalleryBarItem rgbiSkins;
        private DevExpress.XtraBars.Ribbon.RibbonPage ribbonPageSkins;
        private DevExpress.XtraBars.Ribbon.RibbonPageGroup skinsRibbonPageGroup;
        private DevExpress.XtraBars.Ribbon.RibbonPage helpRibbonPage;
        private DevExpress.XtraBars.Ribbon.RibbonPageGroup helpRibbonPageGroup;
        private DevExpress.XtraBars.Ribbon.ApplicationMenu appMenu;
        private DevExpress.XtraBars.PopupControlContainer popupControlContainer1;
        private DevExpress.XtraEditors.LabelControl someLabelControl2;
        private DevExpress.XtraEditors.LabelControl someLabelControl1;
        private DevExpress.XtraBars.PopupControlContainer popupControlContainer2;
        private DevExpress.XtraEditors.ButtonEdit buttonEdit;
        private DevExpress.XtraBars.Ribbon.RibbonStatusBar ribbonStatusBar;
        private DevExpress.Utils.ImageCollection ribbonImageCollection;
        private DevExpress.Utils.ImageCollection ribbonImageCollectionLarge;
        private System.Windows.Forms.Panel spreadsheetFormulaBarPanel;
        private DevExpress.XtraSpreadsheet.SpreadsheetControl spreadsheetControl;
        private DevExpress.XtraEditors.SplitterControl splitterControl;
        private DevExpress.XtraEditors.SplitContainerControl formulaBarNameBoxSplitContainerControl;
        private DevExpress.XtraSpreadsheet.SpreadsheetNameBoxControl spreadsheetNameBoxControl;
        private DevExpress.XtraSpreadsheet.SpreadsheetFormulaBarControl spreadsheetFormulaBarControl1;
        private DevExpress.XtraBars.BarButtonItem iOpenFile;
        private DevExpress.XtraBars.Ribbon.RibbonPage FileRibbonPage;
        private DevExpress.XtraBars.Ribbon.RibbonPageGroup fileRibbonPageGroup;
        private DevExpress.XtraBars.BarButtonItem iCreationScript;
        private DevExpress.XtraBars.Ribbon.RibbonPageGroup ScriptRibbonPageGroup;
        private DevExpress.XtraRichEdit.RichEditControl richEditControlScriptText;
        private DevExpress.XtraBars.Docking.DockManager dockManager1;
        private DevExpress.XtraBars.Docking.DockPanel dockPanel1;
        private DevExpress.XtraBars.Docking.ControlContainer dockPanel1_Container;
    }
}
