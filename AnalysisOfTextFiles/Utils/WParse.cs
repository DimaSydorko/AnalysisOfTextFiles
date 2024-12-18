﻿using System;
using System.Linq;
using System.Threading.Tasks;
using AnalysisOfTextFiles.Objects;
using DocumentFormat.OpenXml.Wordprocessing;

public class WParse
{
  public async static Task Content()
  {
    WReport.CreateReportFile();
    State.Styles = WStyles.GetDocStyles();
    WStyles.Review();

    _PageFormat();
    _Header();
    _Footer();
    await Task.Run(() => _Body());

    State.WDocument.Close();
  }

  public static void StyleSettings()
  {
    WReport.CreateStylesFile();
    State.Styles = WStyles.GetDocStyles();
    CheckPage.AnalisePageSettings();
    WStyles.AnaliseStylesSettings();
    State.WDocument.Close();
  }

  private static async Task _Body()
  {
    WReport.Write("----Body----");

    var body = State.WDocument.MainDocumentPart.Document.Body;
    var descendants = body.Descendants<Paragraph>().ToList();

    async Task AnalyzeParagraph(Paragraph parDesc)
    {
      var idx = descendants.IndexOf(parDesc);
      State.NextParagraphName =
        WDecoding.RemoveSuffixIfExists(
          CheckParagraph.GetParagraphStyle(idx == descendants.Count - 1 ? null : descendants[idx + 1]));
      State.PrevParagraphName =
        WDecoding.RemoveSuffixIfExists(CheckParagraph.GetParagraphStyle(idx == 0 ? null : descendants[idx - 1]));

      if (parDesc.Parent.LocalName == "sdtContent")
      {
       await CheckParagraph.ParagraphCheck(parDesc, idx, CheckParagraph.ContentType.TOC);
      }
      else if (parDesc.Parent != null && parDesc.Parent is TableCell)
      {
        var cell = (TableCell)parDesc.Parent;
        var row = (TableRow)parDesc.Parent.Parent;
        var table = (Table)parDesc.Parent.Parent.Parent;

        var parIdx = cell.Descendants<Paragraph>().ToList().IndexOf(parDesc);
        var cellIdx = row.Descendants<TableCell>().ToList().IndexOf(cell);
        var rowIdx = table.Descendants<TableRow>().ToList().IndexOf(row);
        var tableIdx = body.Descendants<Table>().ToList().IndexOf(table);

        var Wtable = new WTable(tableIdx, rowIdx, cellIdx, parIdx);

        await CheckParagraph.ParagraphCheck(parDesc, idx, CheckParagraph.ContentType.Table, Wtable);
      }
      else
      {
        await CheckParagraph.ParagraphCheck(parDesc, idx, CheckParagraph.ContentType.Paragraph);
      }
    }

    foreach (var parDesc in descendants)
    {
      await AnalyzeParagraph(parDesc);

      Console.WriteLine($"Processing at {DateTime.Now:HH:mm:ss}");
    }
  }

  private static void _Header()
  {
    WReport.Write("----Header & Footer----");

    var headerParts = State.WDocument.MainDocumentPart.HeaderParts.ToList();

    foreach (var headerPart in headerParts)
    {
      var headers = headerPart.Header;
      var paragraphs = headers.Descendants<Paragraph>().ToList();
      foreach (var paragraph in paragraphs)
      {
        var idx = paragraphs.IndexOf(paragraph);
        CheckParagraph.ParagraphCheck(paragraph, idx, CheckParagraph.ContentType.Header);
      }
    }
  }

  private static void _Footer()
  {
    var footerParts = State.WDocument.MainDocumentPart.FooterParts.ToList();

    foreach (var footerPart in footerParts)
    {
      var footers = footerPart.Footer;
      var paragraphs = footers.Descendants<Paragraph>().ToList();
      foreach (var paragraph in paragraphs)
      {
        var idx = paragraphs.IndexOf(paragraph);
        CheckParagraph.ParagraphCheck(paragraph, idx, CheckParagraph.ContentType.Footer);
      }
    }
  }

  private static void _PageFormat()
  {
    WReport.Write("----Page format----");

    var body = State.WDocument.MainDocumentPart.Document.Body;

    foreach (var section in body.Elements<SectionProperties>())
    {
      CheckPage.CheckDimensions(section);
      CheckPage.CheckPageMargin(section);
    }
  }
}