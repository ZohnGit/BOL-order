using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

namespace BolOrderExporter;

internal static class XlsxWriter
{
    private static readonly string[] Headers =
    {
        "orderId", "buyerName", "buyerEmail", "paymentTime", "countryCode", "trackAndTrace",
        "amount", "sku", "quantity", "quantityOriginal", "quantityCancelled", "shipmentDateTime"
    };

    public static byte[] Create(IReadOnlyList<ExportRow> rows)
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            WriteText(archive, "[Content_Types].xml", ContentTypes());
            WriteText(archive, "_rels/.rels", RootRelationships());
            WriteText(archive, "xl/workbook.xml", Workbook());
            WriteText(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationships());
            WriteText(archive, "xl/styles.xml", Styles());
            WriteWorksheet(archive, rows);
        }
        return stream.ToArray();
    }

    private static void WriteWorksheet(ZipArchive archive, IReadOnlyList<ExportRow> rows)
    {
        var entry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, Settings());
        writer.WriteStartDocument(true);
        writer.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.WriteStartElement("sheetViews");
        writer.WriteStartElement("sheetView");
        writer.WriteAttributeString("workbookViewId", "0");
        writer.WriteStartElement("pane");
        writer.WriteAttributeString("ySplit", "1");
        writer.WriteAttributeString("topLeftCell", "A2");
        writer.WriteAttributeString("activePane", "bottomLeft");
        writer.WriteAttributeString("state", "frozen");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteStartElement("cols");
        var widths = new[] { 18, 20, 32, 26, 13, 24, 13, 22, 12, 18, 20, 26 };
        for (var i = 0; i < widths.Length; i++)
        {
            writer.WriteStartElement("col");
            writer.WriteAttributeString("min", (i + 1).ToString());
            writer.WriteAttributeString("max", (i + 1).ToString());
            writer.WriteAttributeString("width", widths[i].ToString());
            writer.WriteAttributeString("customWidth", "1");
            writer.WriteEndElement();
        }
        writer.WriteEndElement();

        writer.WriteStartElement("sheetData");
        writer.WriteStartElement("row");
        writer.WriteAttributeString("r", "1");
        for (var col = 0; col < Headers.Length; col++) WriteInlineCell(writer, CellRef(col, 1), Headers[col], 1);
        writer.WriteEndElement();

        for (var index = 0; index < rows.Count; index++)
        {
            var rowNumber = index + 2;
            var row = rows[index];
            writer.WriteStartElement("row");
            writer.WriteAttributeString("r", rowNumber.ToString());
            WriteInlineCell(writer, CellRef(0, rowNumber), row.OrderId);
            WriteInlineCell(writer, CellRef(1, rowNumber), row.BuyerName);
            WriteInlineCell(writer, CellRef(2, rowNumber), row.BuyerEmail);
            WriteInlineCell(writer, CellRef(3, rowNumber), row.PaymentTime);
            WriteInlineCell(writer, CellRef(4, rowNumber), row.CountryCode);
            WriteInlineCell(writer, CellRef(5, rowNumber), row.TrackAndTrace);
            WriteNumberCell(writer, CellRef(6, rowNumber), row.Amount);
            WriteInlineCell(writer, CellRef(7, rowNumber), row.Sku);
            WriteNumberCell(writer, CellRef(8, rowNumber), row.Quantity);
            WriteNumberCell(writer, CellRef(9, rowNumber), row.QuantityOriginal);
            WriteNumberCell(writer, CellRef(10, rowNumber), row.QuantityCancelled);
            WriteInlineCell(writer, CellRef(11, rowNumber), row.ShipmentDateTime);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();

        writer.WriteStartElement("autoFilter");
        writer.WriteAttributeString("ref", $"A1:L{Math.Max(1, rows.Count + 1)}");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void WriteInlineCell(XmlWriter writer, string reference, string? value, int style = 0)
    {
        writer.WriteStartElement("c");
        writer.WriteAttributeString("r", reference);
        writer.WriteAttributeString("t", "inlineStr");
        if (style > 0) writer.WriteAttributeString("s", style.ToString());
        writer.WriteStartElement("is");
        writer.WriteElementString("t", Sanitize(value));
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteNumberCell(XmlWriter writer, string reference, decimal? value)
    {
        if (value is null)
        {
            WriteInlineCell(writer, reference, null);
            return;
        }
        writer.WriteStartElement("c");
        writer.WriteAttributeString("r", reference);
        writer.WriteElementString("v", value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        writer.WriteEndElement();
    }

    private static string Sanitize(string? value)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        return new string(value.Where(ch => ch == '\t' || ch == '\n' || ch == '\r' || ch >= ' ').ToArray());
    }

    private static string CellRef(int zeroBasedColumn, int row)
    {
        var column = string.Empty;
        for (var n = zeroBasedColumn + 1; n > 0; n = (n - 1) / 26)
            column = (char)('A' + (n - 1) % 26) + column;
        return column + row;
    }

    private static void WriteText(ZipArchive archive, string path, string text)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(text);
    }

    private static XmlWriterSettings Settings() => new()
    {
        Encoding = new UTF8Encoding(false),
        Indent = false,
        CloseOutput = false
    };

    private static string ContentTypes() => """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
        </Types>
        """;

    private static string RootRelationships() => """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private static string Workbook() => """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets><sheet name="Orders" sheetId="1" r:id="rId1"/></sheets>
        </workbook>
        """;

    private static string WorkbookRelationships() => """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
        """;

    private static string Styles() => """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="2"><font/><font><b/><color rgb="FFFFFFFF"/></font></fonts>
          <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2563EB"/><bgColor indexed="64"/></patternFill></fill></fills>
          <borders count="1"><border/></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>
        </styleSheet>
        """;
}
