using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using AnalysisOfTextFiles.Objects;

public class PageProperties
{
    public static WPage GetPageSettings()
    {
        var cleanedLines = State.Content
            .Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Split(new[] { "//" }, StringSplitOptions.None)[0].Trim()) // Remove comments and trim
            .Where(line => !string.IsNullOrWhiteSpace(line)) // Remove empty lines
            .ToList();

        List<string> size = new List<string>();
        List<string> orientation = new List<string>();
        float marginTop = 0, marginBottom = 0, marginLeft = 0, marginRight = 0, marginHeader = 0, marginFooter = 0;

        // Parse each line and assign values to variables
        foreach (var line in cleanedLines)
        {
            if (line.StartsWith("pageSize"))
            {
                // Split by comma and take valid values
                size = line.Split('=')[1]
                    .Split(',')
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();
            }
            else if (line.StartsWith("orientation"))
            {
                // Split by equals and get orientation list
                orientation = line.Split('=')[1]
                    .Split(',')
                    .Select(s => s.Trim())
                    .ToList();
            }
            else if (line.StartsWith("marginTop"))
            {
                marginTop = ParseSm(line.Split('=')[1].Trim());
            }
            else if (line.StartsWith("marginBottom"))
            {
                marginBottom = ParseSm(line.Split('=')[1].Trim());
            }
            else if (line.StartsWith("marginLeft"))
            {
                marginLeft = ParseSm(line.Split('=')[1].Trim());
            }
            else if (line.StartsWith("marginRight"))
            {
                marginRight = ParseSm(line.Split('=')[1].Trim());
            }
            else if (line.StartsWith("marginHeader"))
            {
                marginHeader = ParseSm(line.Split('=')[1].Trim());
            }
            else if (line.StartsWith("marginFooter"))
            {
                marginFooter = ParseSm(line.Split('=')[1].Trim());
            }
        }

        return new WPage(size, orientation, marginTop, marginBottom, marginLeft, marginRight, marginHeader, marginFooter);
    }

    private static float ParseSm(string value)
    {
        return float.Parse(value.Replace("sm", "").Trim(), CultureInfo.InvariantCulture);
    }
}