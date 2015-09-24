using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PowerPointQuestionnaire.Interfaces;
using PowerPointQuestionnaire.Services;

namespace PowerPointQuestionnaire.Controls
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        private IAuthService _authService;
        private bool _isBusing;

        public LoginWindow()
        {
            InitializeComponent();

            _authService = new AuthService();;
        }

        private async void Ok_Click(object sender, RoutedEventArgs e)
        {
            //TODO validation 

            var username = this.UsernameBox.Text.Trim();
            var password = this.PasswordBox.Password.Trim();

            if (username == string.Empty || password == string.Empty)
            {
                ErrorTextBox.Text = "请正确输入认证信息";
                return;
            }

            if (_isBusing) return;
            _isBusing = true;

            if (await _authService.Authenticate(this.UsernameBox.Text, this.PasswordBox.Password))
            {
                this.DialogResult = true;
                _isBusing = false;
                this.Close();
            }
            else
            {
                ErrorTextBox.Text = "登陆失败了";
                _isBusing = false;
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
