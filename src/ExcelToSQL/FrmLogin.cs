using DevExpress.XtraEditors;
using System;
using System.Windows.Forms;
using Windows.Security.Credentials;

namespace ExcelToSQL
{
    public partial class FrmLogin : XtraForm
    {
        public FrmLogin()
        {
            InitializeComponent();

            this.LoginButton.Click += LoginButton_Click;
            this.HelloButton.Click += HelloButton_Click;
            this.Shown += (s, e) =>
            {
                GetLocalStorageID();
                HelloButton_Click(s, e);
            };
        }

        private void SetLocalStorageID()
        {
            Console.WriteLine($"Saved ID={textEdit_email.Text} to LocalStorage");
        }

        private void GetLocalStorageID()
        {
            Console.WriteLine($"Get ID={textEdit_email.Text} from LocalStorage");
        }

        private void LoginButton_Click(object sender, EventArgs e)
        {
            if(!VerifyIdentity())
                return;

            if(!Authentification())
                return;

            if(checkEdit_keepme.Checked)
                SetLocalStorageID();

            MessageBox.Show($"{textEdit_email.Text} {textEdit_password.Text}");

            // AccessToken을 사용하여 작업을 수행합니다.
            DialogResult = DialogResult.OK;
            this.Close();
        }

        private bool Authentification()
        {
            Console.WriteLine($"Authentification ID={textEdit_email.Text} to Auth Server");
            return true;
        }

        private bool VerifyIdentity()
        {
            Console.WriteLine($"Validation ID={textEdit_email.Text}");
            return true;
        }

        private async void HelloButton_Click(object sender, EventArgs e)
        {
            try
            {
                bool supported = await KeyCredentialManager.IsSupportedAsync();
                if(supported)
                {
                    KeyCredentialRetrievalResult result =
                        await KeyCredentialManager.RequestCreateAsync("login",
                        KeyCredentialCreationOption.ReplaceExisting);
                    if(result.Status != KeyCredentialStatus.Success)
                        throw new NotImplementedException();
                    // AccessToken을 사용하여 작업을 수행합니다.
                    DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
            catch(Exception ex)
            {
                // 인증 실패시 예외 처리

                MessageBox.Show(ex.Message);
            }
        }
    }
}
