using System.IO;
using System.Text.RegularExpressions;
using System.Windows;

namespace AnalysisOfTextFiles;

public partial class AdminChangePassWindow
{
  public AdminChangePassWindow()
  {
    InitializeComponent();
  }

  private void BtnSubmit_Click(object sender, RoutedEventArgs e)
  {
    var password = txtPassword.Password;
    var repeatPass = repeatPassword.Password;

    if (!Equals(password, repeatPass))
    {
      MessageBox.Show("Passwords didn't match!");
      return;
    }

    if (!ValidateCredentials(password))
    {
      MessageBox.Show("Password should contain at least 1 special character, 1 numeric character, 1 uppercase letter");
      return;
    }

    HashAndSavePassword(password);
    MessageBox.Show("Password updated successful!");
    Hide();
  }

  private void BtnClose_Click(object sender, RoutedEventArgs e)
  {
    Hide();
  }

  private bool ValidateCredentials(string password)
  {
    if (password.Length < 8)
      return false;

    if (!Regex.IsMatch(password, @"[!@#$%^&*(),.?""':{}|<>]"))
      return false;

    if (!Regex.IsMatch(password, @"\d"))
      return false;

    if (!Regex.IsMatch(password, @"[A-Z]"))
      return false;

    return true;
  }

  public static void HashAndSavePassword(string password)
  {
    var encodedPass = AdminSettings.EncodeDataToHash(password);
    File.WriteAllText("sec.ini", encodedPass);
  }
}