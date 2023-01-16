using DevExpress.XtraSplashScreen;
using System;

namespace Presentation.Views.Excel
{
    public partial class SplashScreenForm : SplashScreen
    {
        public SplashScreenForm()
        {
            InitializeComponent();
        }

        #region Overrides

        public override void ProcessCommand(Enum cmd, object arg)
        {
            base.ProcessCommand(cmd, arg);
            SplashScreenCommand command = (SplashScreenCommand)cmd;
            if(command == SplashScreenCommand.SetProgress)
            {
                int pos = (int)arg;
                progressBarControl1.Position = pos;
            }
            if(command == SplashScreenCommand.SetMessage)
            {
                labelControl_message.Text = (string)arg;
            }
        }

        #endregion

        public enum SplashScreenCommand
        {
            SetProgress,
            SetMessage,
            Command3
        }
    }
}
