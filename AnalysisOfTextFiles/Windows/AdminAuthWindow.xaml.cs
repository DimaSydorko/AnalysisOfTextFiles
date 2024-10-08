﻿using System;
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

  public event VisibilityChangedEventHandler IsAdminEditBtn, IsAdminChangePassBtn, IsAdminAuthBtn;

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

    var encodedPass = AdminSettings.EncodeDataToHash(password);

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