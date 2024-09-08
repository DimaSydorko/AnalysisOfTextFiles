using System;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Windows;

public class AdminSettings
{
  private static readonly string settingsFilePath = "styleSettings.ini";
  private static readonly string hashFilePath = "hash.ini";

  public static string GetStyleData()
  {
    string iniData = null;
    string iniDataHash = null;

    try
    {
      iniData = File.ReadAllText(settingsFilePath);
      iniDataHash = File.ReadAllText(hashFilePath);
    }
    catch (Exception ex)
    {
      MessageBox.Show(ex.Message, "Error");
      return "";
    }

    var isDataDamaged = !string.Equals(iniDataHash, EncodeDataToHash(iniData));

    if (isDataDamaged)
    {
      MessageBox.Show("Settings file was damaged", "Error");

      var isAdmin = IsUserAdmin();
      if (isAdmin) return "";
    }

    return iniData;
  }

  public static string GetStyleSettings(string settingsData)
  {
    var index = settingsData.IndexOf("\n");
    var settings = settingsData.Substring(index + 1);
    return settings;
  }

  public static string GetStyleKeyWord(string settingsData)
  {
    var index = settingsData.IndexOf('\n');
    var keyWord = settingsData.Substring(0, index);
    return keyWord;
  }

  public static void SetStyleSettings(string styleSettings, string keyWord)
  {
    if (!IsUserAdmin())
    {
      MessageBox.Show("Access denied. Administrator privileges required.", "Warning");
      return;
    }

    var decData = $"{keyWord}\n{styleSettings}";

    var encData = EncodeDataToHash(decData);
    if (encData != null)
    {
      File.WriteAllText(settingsFilePath, decData);
      File.WriteAllText(hashFilePath, encData);

      MessageBox.Show("INI file updated successfully.", "Success");
    }
  }

  public static string EncodeDataToHash(string dataToEncode)
  {
    var dataBytes = Encoding.UTF8.GetBytes(dataToEncode);
    using (var sha256 = SHA256.Create())
    {
      var hashBytes = sha256.ComputeHash(dataBytes);
      var hashedData = Convert.ToBase64String(hashBytes);
      return hashedData;
    }
  }

  public static bool IsUserAdmin()
  {
    var identity = WindowsIdentity.GetCurrent();
    var principal = new WindowsPrincipal(identity);
    var isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);

    if (Debugger.IsAttached) return true;
    return isAdmin;
  }
}