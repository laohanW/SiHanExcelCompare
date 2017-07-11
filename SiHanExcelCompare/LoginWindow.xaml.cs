using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
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

namespace SiHanExcelCompare
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();
        }

        private void Login_btn_Click(object sender, RoutedEventArgs e)
        {
            var dir = AppDomain.CurrentDomain.BaseDirectory;
            var hostName = Dns.GetHostName();
            string hostFile = dir + "host.l";
            string pwdFile = dir + "pwd.l";
            string targetHost = "";
            string targetPwd = "";
            using (var fileStream=new FileStream(hostFile,FileMode.Open))
            {
                byte[] bytes = new byte[fileStream.Length];
                fileStream.Read(bytes, 0, bytes.Length);
                targetHost = System.Text.Encoding.UTF8.GetString(bytes);
            }
            using (var fileStream = new FileStream(pwdFile, FileMode.Open))
            {
                byte[] bytes = new byte[fileStream.Length];
                fileStream.Read(bytes, 0, bytes.Length);
                targetPwd = System.Text.Encoding.UTF8.GetString(bytes);
            }
            if (hostName.Equals(targetHost))
            {
                var localStr = account_text.Text + password_text.Password;
                targetPwd = targetPwd.Replace("\n", String.Empty).Replace("\r",string.Empty);
                if (targetPwd.Equals(localStr))
                {
                    MainWindow mainWindow = new MainWindow();
                    mainWindow.Show();
                    Close();
                }
                else
                {
                    if (MessageBox.Show("账号密码错误", "错误", MessageBoxButton.OK) == MessageBoxResult.OK)
                    {

                    }
                }
            }
            else
            {
                if(MessageBox.Show("电脑未注册！", "错误", MessageBoxButton.OK)==MessageBoxResult.OK)
                {
                    Close();
                }
            }
        }
    }
}
