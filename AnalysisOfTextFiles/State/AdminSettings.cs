using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Windows;

public class AdminSettings
{
  private static string settingsFilePath = "styleSettings.ini";
  private static string hashFilePath = "hash.ini";

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

    bool isDataDamaged = !string.Equals(iniDataHash, EncodeDataToHash(iniData));

    if (isDataDamaged)
    {
      MessageBox.Show("Settings file was damaged", "Error");

      bool isAdmin = IsUserAdmin();
      if (isAdmin) return "";
    }
    
    return iniData;
  }

  public static string GetStyleSettings(string settingsData)
  {
    int index = settingsData.IndexOf("\n");
    string settings = settingsData.Substring(index + 1);
    return settings;
  }
  public static string GetStyleKeyWord(string settingsData)
  {
    int index = settingsData.IndexOf('\n');
    string keyWord = settingsData.Substring(0, index);
    return keyWord;
  }

  public static void SetStyleSettings(string styleSettings, string keyWord)
  {
    if (!IsUserAdmin())
    {
      MessageBox.Show("Access denied. Administrator privileges required.", "Warning");
      return;
    }

    string decData = $"{keyWord}\n{styleSettings}";
    
    string? encData = EncodeDataToHash(decData);
    if (encData != null)
    {
      File.WriteAllText(settingsFilePath, decData);
      File.WriteAllText(hashFilePath, encData);

      MessageBox.Show("INI file updated successfully.", "Success");
    }
  }

  public static string EncodeDataToHash(string dataToEncode)
  {
    byte[] dataBytes = Encoding.UTF8.GetBytes(dataToEncode);
    using (SHA256 sha256 = SHA256.Create())
    {
      byte[] hashBytes = sha256.ComputeHash(dataBytes);
      string hashedData = Convert.ToBase64String(hashBytes);
      return hashedData;
    }
  }

  public static bool IsUserAdmin()
  {
    WindowsIdentity? identity = WindowsIdentity.GetCurrent();
    WindowsPrincipal? principal = new WindowsPrincipal(identity);
    bool isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
    
    if (DevState.IsAdmin) return true;
    return isAdmin;
  }
}