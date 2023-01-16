using DevExpress.XtraPdfViewer;
using DevExpress.XtraPdfViewer.Commands;
using System.Windows.Forms;


namespace Presenter.Views.PDF
{
    public class CustomNextPageCommand : PdfNextPageCommand
    {
        public CustomNextPageCommand(PdfViewer control)
            : base(control)
        {
        }

        public override void Execute()
        {
            //MessageBox.Show("Custom command executed");
            base.Execute();
        }
    }
}
