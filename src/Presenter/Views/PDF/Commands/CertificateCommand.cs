using DevExpress.XtraPdfViewer;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using DevExpress.Pdf;
using System;

namespace Presenter.Views.PDF
{
    public class CertificateCommand : ICommand
    {
        PdfDocumentProcessor _pdfDocumentProcessor;
        PdfViewer _prfViewer;
        public CertificateCommand(PdfViewer control)
        {
            _prfViewer = control;
        }

        public void Execute()
        {
            MemoryStream reloadMS = new MemoryStream();
            // Create a PDF document processor.
            using(PdfDocumentProcessor documentProcessor = new PdfDocumentProcessor())
            {

                MemoryStream ms = new MemoryStream();
                _prfViewer.SaveDocument(ms);
                // Load a PDF document. 
                documentProcessor.LoadDocument(ms);
                
                // Load a certificate from a file.
                X509Certificate2 cert = new X509Certificate2(@"..\..\..\..\certificate.pfx", "qlalfqjsgh");

                // Create a PDF signature and specify signing location, contact info and reason.
                PdfSignature signature = new PdfSignature(cert)
                {
                    Location = "Seoul",
                    ContactInfo = "dev@acme.kr",
                    Reason = "Approved"
                };
                // 관리자/사용자 패스워드 설정, 출력 권한 설정, 수정 권한 설정
                PdfEncryptionOptions enc = new PdfEncryptionOptions()
                {
                    OwnerPasswordString = "all",
                    UserPasswordString = "1234"
                    , PrintingPermissions = PdfDocumentPrintingPermissions.NotAllowed
                    , ModificationPermissions = PdfDocumentModificationPermissions.NotAllowed
                };


                // Save the signed document.
                documentProcessor.SaveDocument(reloadMS, new PdfSaveOptions() { Signature = signature, EncryptionOptions = enc});
                ms.Dispose();
            }

            _prfViewer.LoadDocument(reloadMS);
        }

        //public override void Execute()
        //{
        //    base.Execute();
        //}
    }
}
