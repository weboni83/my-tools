using DevExpress.XtraBars.Docking2010.Customization;
using DevExpress.XtraBars.Docking2010.Views.WindowsUI;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Presenter.Views.Dialogs
{
    public partial class AutoClosingDailog : Form
    {
        public AutoClosingDailog()
        {
            InitializeComponent();
        }

        public DialogResult ShowForm(string message)
        {

            FlyoutAction action = new FlyoutAction() { Caption = "확인", Description = message };
            Predicate<DialogResult> predicate = canCloseFunc;
            FlyoutCommand command1 = new FlyoutCommand() { Text = "Close", Result = System.Windows.Forms.DialogResult.OK };
            FlyoutCommand command2 = new FlyoutCommand() { Text = "Cancel", Result = System.Windows.Forms.DialogResult.Cancel };
            action.Commands.Add(command1);
            action.Commands.Add(command2);
            FlyoutProperties properties = new FlyoutProperties();
            properties.ButtonSize = new Size(200, 40);
            properties.Style = FlyoutStyle.MessageBox;
            properties.AppearanceDescription.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;

            const int AUTO_CLOSE_INTERVAL = 10;
            // button 은 document usercontrol 로 구성

            Timer timer = new Timer();
            timer.Interval = 1000;

            int count = 0;
            int msec = 0;
            timer.Tick += (s, e) =>
            {
                msec += timer.Interval;
                action.Commands.Remove(command1);
                action.Commands.Remove(command2);
                var caption = $"{message} ({AUTO_CLOSE_INTERVAL - count++})";
                command1 = new FlyoutCommand() { Text = caption, Result = DialogResult.OK };
                action.Commands.Add(command1);
                action.Commands.Add(command2);
                if(msec > AUTO_CLOSE_INTERVAL * 1000)
                    command1.Result = DialogResult.Yes;
            };
            timer.Start();

            var result = FlyoutDialog.Show(this, action, properties, predicate);
            if(result != DialogResult.OK)
                return result;

            return DialogResult = DialogResult.OK;
        }

        private bool canCloseFunc(DialogResult parameter)
        {
            return parameter != DialogResult.Cancel;
        }
    }
}
