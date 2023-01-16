using System;
using DevExpress.Utils;

namespace Commands
{
    public struct PdfCustomCommandId : IEquatable<PdfCustomCommandId>
    {

        //
        // 요약:
        //     An undefined command.
        public static readonly PdfCustomCommandId None;
        public static readonly PdfCustomCommandId CertificateCommand;


        public bool Equals(PdfCustomCommandId other)
        {
            throw new NotImplementedException();
        }
    }
}
