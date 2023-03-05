using System.Collections.Generic;
using System.Linq;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;
using DocumentFormat.OpenXml;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
public class WParse
{
  public static void Content()
  {
    // Assign a reference to the appropriate part to the stylesPart variable.
    MainDocumentPart mainPart = State.WDocument.MainDocumentPart;
    Body body = mainPart.Document.Body;

    State.Styles = Styles();
    WReport.CreateReportFile();

    //--------header-------- 
    foreach (HeaderPart headerPart in mainPart.HeaderParts.ToList())
    {
      Header headers = headerPart.Header;
      List<Paragraph> paragraphs = headers.Descendants<Paragraph>().ToList();
      foreach (var paragraph in paragraphs)
      {
        int idx = paragraphs.IndexOf(paragraph);
        Analis.ParagraphCheck(paragraph, idx, Analis.ContentType.Header);
      }
    }

    //--------footer-------- 
    foreach (FooterPart footerPart in mainPart.FooterParts.ToList())
    {
      Footer footers = footerPart.Footer;
      List<Paragraph> paragraphs = footers.Descendants<Paragraph>().ToList();
      foreach (var paragraph in paragraphs)
      {
        int idx = paragraphs.IndexOf(paragraph);
        Analis.ParagraphCheck(paragraph, idx, Analis.ContentType.Footer);
      }
    }
    //--------paragraph-------- 
    // List<Paragraph> bodyParagraphs = body.Elements<Paragraph>().ToList();
    // foreach (Paragraph paragraph in bodyParagraphs)
    // {
    //   int idx = bodyParagraphs.IndexOf(paragraph);
    //   Analis.ParagraphCheck(paragraph, mainPart, idx, Analis.ContentType.Paragraph);
    // }

    // //--------tables-------- 
    // List<Table> tables = body.Descendants<Table>().ToList();
    // foreach (Table table in tables)
    // {
    //   int idx = tables.IndexOf(table);
    //   IEnumerable<TableCell> cells = table.Descendants<TableCell>();
    //   foreach (TableCell cell in cells)
    //   {
    //     IEnumerable<Paragraph> paragraphs = cell.Descendants<Paragraph>();
    //     foreach (Paragraph paragraph in paragraphs)
    //     {
    //       Analis.ParagraphCheck(paragraph, mainPart, idx, Analis.ContentType.Table);
    //     }
    //   }
    // }

    //--------File parsing-------- 
    List<Paragraph> descendants = body.Descendants<Paragraph>().ToList();
    foreach (Paragraph parDesc in descendants)
    {
      int idx = descendants.IndexOf(parDesc);
      if (((string)parDesc.ParagraphProperties?.ParagraphStyleId?.Val)?.StartsWith("TOC") == true)
      {
        // This is a TOC entry
        Analis.ParagraphCheck(parDesc, idx, Analis.ContentType.TOC);
      }
      else
      {
        // Check if the paragraph is part of a table
        if (parDesc.Parent != null && parDesc.Parent.LocalName == "tc")
        {
          // This is a table row
          Analis.ParagraphCheck(parDesc, idx, Analis.ContentType.Table);
        }
        else
        {
          // This is an ordinary paragraph
          Analis.ParagraphCheck(parDesc, idx, Analis.ContentType.Paragraph);
        }
      }
    }
  }

  public static List<WStyle> Styles()
  {
    List<WStyle> styles = new List<WStyle>();
    Styles stylesXml = State.WDocument.MainDocumentPart.StyleDefinitionsPart.Styles;
    
    //Map to get Encoded and Decoded StyleNames
    foreach (var styleXml in stylesXml.ChildElements)
    {
      if (styleXml.ChildElements.Count >= 3)
      {
        OpenXmlElement styleDec = styleXml.ChildElements[0];
        OpenXmlElement styleEnc = styleXml.ChildElements[2];

        WStyle style = new WStyle();

        void GetName (PropertyInfo property, OpenXmlElement styleXml, bool isDec)
        {
          if (property != null)
          {
            var styleNameObj = property.GetValue(styleXml);
            if (styleNameObj != null)
            {
              PropertyInfo propertyName = styleNameObj.GetType().GetProperty("Value");
              if (propertyName != null)
              {
                string name = propertyName.GetValue(styleNameObj)?.ToString();
                if (!string.IsNullOrEmpty(name))
                {
                  if (isDec) style.SetDec(name);
                  else style.SetEnc(name);
                }
              }
            }
          }
        }
        
        PropertyInfo propertyDec = styleDec.GetType().GetProperty("Val");
        GetName(propertyDec, styleDec, true);

        PropertyInfo propertyEnc = styleEnc.GetType().GetProperty("Val");
        GetName(propertyEnc, styleEnc, false);

        //Rewrite TOC style names
        string[] tocStyles = { "toc 1", "toc 2", "toc 3", "TOC Heading" };
        if (tocStyles.Contains(style.decoded))
        {
          string newEncoded = style.decoded.Replace(" ", "");
          string Upper = newEncoded.Substring(0, 3).ToUpper() + newEncoded.Substring(3);
          style.encoded = Upper;
        }

        //Save only styles which exist   
        bool isNotAllowedStyle = style.encoded != null && style.encoded != "CommentText";

        if (style.decoded != null && isNotAllowedStyle) styles.Add(style);
        else if (style.decoded == "Normal" && isNotAllowedStyle) styles.Add(style);
      }
    }

    return styles;
  }
}