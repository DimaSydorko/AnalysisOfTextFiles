using System;
using System.IO;
using System.Windows;

namespace AnalysisOfTextFiles;

public partial class AdminAuthWindow
{
  public delegate void VisibilityChangedEventHandler(bool visibility);

  public AdminAuthWindow()
  {
    InitializeComponent();
  }

  public event VisibilityChangedEventHandler IsAdminEditBtn,
    IsAdminChangePassBtn,
    IsAdminAuthBtn,
    IsAdminGetDocSttingsBtn;

  private void BtnLogin_Click(object sender, RoutedEventArgs e)
  {
    var password = txtPassword.Password;
    var isVerify = VerifyPassword(password);

    if (isVerify)
    {
      State.IsAdminAuth = true;

      IsAdminAuthBtn?.Invoke(false);
      IsAdminEditBtn?.Invoke(true);
      IsAdminChangePassBtn?.Invoke(true);
      IsAdminGetDocSttingsBtn?.Invoke(true);

      MessageBox.Show("Login successful!");
      Close();
    }
    else
    {
      MessageBox.Show("Invalid password.");
    }
  }

  private void BtnClose_Click(object sender, RoutedEventArgs e)
  {
    Hide();
  }

  public static bool VerifyPassword(string password)
  {
    string storedPass = null;
    try
    {
      storedPass = File.ReadAllText("sec.ini");
    }
    catch (Exception e)
    {
    }

    var encodedPass = AdminSettings.EncodeDataToHash(password);

    if (!string.IsNullOrEmpty(storedPass))
    {
      return string.Equals(encodedPass, storedPass);
    }
    
    return string.Equals(password, "admin");
  }
}