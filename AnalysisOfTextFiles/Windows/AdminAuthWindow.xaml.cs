using System;
using System.IO;
using System.Windows;


namespace AnalysisOfTextFiles;

public partial class AdminAuthWindow
{
  public AdminAuthWindow()
  {
    InitializeComponent();
  }
  
  public delegate void VisibilityChangedEventHandler(bool visibility);
  public event VisibilityChangedEventHandler IsAdminEditBtn, IsAdminChangePassBtn, IsAdminAuthBtn;
  
  private void BtnLogin_Click(object sender, RoutedEventArgs e)
  {
    string password = txtPassword.Password;
    bool isVerify = VerifyPassword(password);

    if (isVerify)
    {
      State.IsAdminAuth = true;

      IsAdminAuthBtn?.Invoke(false);
      IsAdminEditBtn?.Invoke(true);
      IsAdminChangePassBtn?.Invoke(true);
      
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
    Close();
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

    string encodedPass = AdminSettings.EncodeDataToHash(password);

    if (!string.IsNullOrEmpty(storedPass))
    {
      if (string.Equals(encodedPass, storedPass))
        return true;
      return false;
    }

    {
      if (string.Equals(password, "admin"))
        return true;
      return false;
    }
  }
}