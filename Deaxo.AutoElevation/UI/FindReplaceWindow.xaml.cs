using System.Security.Cryptography.X509Certificates;
using System.Windows;

namespace Deaxo.AutoElevation.UI
{
    public partial class FindReplaceWindow : Window
    {
        public string FindText => TxtFind.Text;
        public string ReplaceText => TxtReplace.Text;
        public string Prefix => TxtPrefix.Text;
        public string Suffix => TxtSuffix.Text;

        public FindReplaceWindow(string title = "Find and Replace")
        {
            InitializeComponent();

            this.Title = title;
            BtnClose.Click += (s, e) => { this.DialogResult = false; this.Close(); };
            BtnRename.Click += (s, e) =>
            {
                this.DialogResult = true;
                this.Close();
            };
        }
    }
}
