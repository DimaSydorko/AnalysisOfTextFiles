using System;
using System.IO;
using System.Security.Principal;
using System.Text;
using System.Windows;

public class AdminSettings
{
  private static string iniFilePath = "styleSettings.ini";
  
  public static string GetStyleSettings()
  {
    string iniDataHash = null;

    try
    {
      iniDataHash = File.ReadAllText(iniFilePath);
    }
    catch (Exception ex)
    {
      MessageBox.Show(ex.Message, "Error");
      return "";
    }

    if (iniDataHash != null)
    {
      return DecodeHash(iniDataHash);
    }
    return "";
  }

  public static void SetStyleSettings(string decData)
  {
    if (!IsUserAdmin())
    {
      MessageBox.Show("Access denied. Administrator privileges required.", "Warning");
      return;
    }

    string? encData = EncodeDataToHash(decData);
    if (encData != null)
    {
      File.WriteAllText(iniFilePath, encData);

      MessageBox.Show("INI file updated successfully.", "Success");
    }
  }
  
  public static string EncodeDataToHash(string dataToEncode)
  {
    var plainTextBytes = Encoding.UTF8.GetBytes(dataToEncode);
    return Convert.ToBase64String(plainTextBytes);
  }

  public static string DecodeHash(string hashedData)
  {
    try
    {
    byte[] base64EncodedBytes = Convert.FromBase64String(hashedData);
    string decodedData = Encoding.UTF8.GetString(base64EncodedBytes);
    return decodedData;
    }
    catch (Exception ex)
    {
      MessageBox.Show($"Encoding: {ex.Message}", "Error");
      return null;
    }
  }

  public static bool IsUserAdmin()
  {
    WindowsIdentity? identity = WindowsIdentity.GetCurrent();
    WindowsPrincipal? principal = new WindowsPrincipal(identity);
    bool isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
    return !isAdmin;
  }
}