using Core.Exceptions;
using DevExpress.XtraPdfViewer;
using System;
using System.Windows.Forms;

namespace Presenter.Views.PDF
{
    public partial class Viewer : UserControl
    {
        public EventHandler<FormExceptionEventArgs> ExceptionTrigger;
        public Viewer()
        {
            InitializeComponent();
            // 기본 PDF 버튼 동작을 재정의 
            ReplacePdfViewerCommandFactoryService(this.pdfViewer);

            this.ribbonControl1.ItemClick += (s, e) =>
            {
                System.Console.WriteLine($"{e.Item.Id} Click");

                if(e.Item.Id == 30)
                {
                    ICommand command = new CertificateCommand(this.pdfViewer);
                    command.Execute();
                    return;
                }

            };
        }

        void ReplacePdfViewerCommandFactoryService(PdfViewer control)
        {
            IPdfViewerCommandFactoryService service = control.GetService<IPdfViewerCommandFactoryService>();
            if(service == null)
                return;
            control.RemoveService(typeof(IPdfViewerCommandFactoryService));
            control.AddService(typeof(IPdfViewerCommandFactoryService), new CustomPdfViewerCommandFactoryService(control, service));
        }

    }
}