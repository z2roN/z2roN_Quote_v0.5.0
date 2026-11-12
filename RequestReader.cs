using System.Globalization;
using System.Text;
using ExcelDataReader;

namespace ZQuote;

public class ColMap
{
    public int HeaderRow = -1, Name = -1, Spec = -1, Unit = -1, Qty = -1, Price = -1, Note = -1, Barcode = -1;
    public bool Ok => Name >= 0 && (Qty >= 0 || Unit >= 0);
}

/// 견적의뢰 엑셀(.xls/.xlsx/.csv)의 첫 시트를 읽어 품목 목록으로 바꾼다.
public static class RequestReader
{
    /// 파일 종류를 내용으로 판별해서 읽는다. (확장자가 .xls 여도 진짜 엑셀 / HTML 표 / XML 스프레드시트 / 텍스트(CSV·탭) 모두 지원)
    public static List<string[]> ReadRows(string file)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        string kind = Sniff(file);
        if (kind == "html") return ReadHtml(file);
        if (kind == "xml") return ReadXmlSheet(file);
        if (kind == "text") return ReadText(file);
        try { return ReadBinary(file); }
        catch (Exception first)
        {
            // 엑셀로 안 읽히면 내용이 텍스트 계열일 수 있으니 한 번 더 시도
            try { var r = ReadHtml(file); if (r.Count > 0) return r; } catch { }
            try { var r = ReadXmlSheet(file); if (r.Count > 0) return r; } catch { }
            try { var r = ReadText(file); if (r.Count > 0) return r; } catch { }
            throw new Exception("엑셀 형식을 알 수 없습니다. (" + first.Message + ")");
        }
    }

    static List<string[]> ReadBinary(string file)
    {
        var rows = new List<string[]>();
        using var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var cfg = new ExcelReaderConfiguration { FallbackEncoding = Encoding.GetEncoding(949) };
        using IExcelDataReader r = file.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) ? ExcelReaderFactory.CreateCsvReader(fs, cfg) : ExcelReaderFactory.CreateReader(fs, cfg);
        while (r.Read() && rows.Count < 20000)
        {
            var a = new string[r.FieldCount];
            for (int i = 0; i < a.Length; i++) a[i] = Str(r.GetValue(i));
            rows.Add(a);
        }
        return rows;
    }

    /// binary(.xlsx / 진짜 .xls) | html | xml | text
    static string Sniff(string file)
    {
        byte[] h = new byte[2048]; int n;
        using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) n = fs.Read(h, 0, h.Length);
        if (n < 4) return "binary";
        if (h[0] == 'P' && h[1] == 'K') return "binary";                                   // xlsx (zip)
        if (h[0] == 0xD0 && h[1] == 0xCF && h[2] == 0x11 && h[3] == 0xE0) return "binary";   // xls (OLE)
        if (h[0] == 0x09 && (h[1] == 0x00 || h[1] == 0x02 || h[1] == 0x04 || h[1] == 0x08) && h[3] == 0x00) return "binary"; // 아주 옛 xls (BIFF2~5)
        int bad = 0; for (int i = 0; i < n; i++) if (h[i] == 0) bad++;
        if (bad > n / 20) return "binary";                                                   // 0 바이트가 많으면 이진 파일
        string t = DecodeText(h.Take(n).ToArray()).TrimStart('\uFEFF', ' ', '\r', '\n', '\t').ToLowerInvariant();
        if (t.StartsWith("<?xml") || t.StartsWith("<workbook") || t.Contains("urn:schemas-microsoft-com:office:spreadsheet")) return "xml";
        if (t.StartsWith("<")) return "html";
        return "text";
    }

    static string DecodeText(byte[] b)
    {
        if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return Encoding.UTF8.GetString(b, 3, b.Length - 3);
        if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE) return Encoding.Unicode.GetString(b, 2, b.Length - 2);
        if (b.Length >= 2 && b[0] == 0xFE && b[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(b, 2, b.Length - 2);
        try { return new UTF8Encoding(false, true).GetString(b); }
        catch { return Encoding.GetEncoding(949).GetString(b); }   // 한글 윈도우 기본(EUC-KR/CP949)
    }

    static string ReadAll(string file)
    {
        byte[] b;
        using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) { b = new byte[fs.Length]; fs.ReadExactly(b, 0, b.Length); }
        return DecodeText(b);
    }

    // 웹 시스템에서 내려받은 '엑셀' (사실은 HTML 표)
    static List<string[]> ReadHtml(string file)
    {
        string html = ReadAll(file);
        var rows = new List<string[]>();
        foreach (System.Text.RegularExpressions.Match tr in System.Text.RegularExpressions.Regex.Matches(html, "<tr[^>]*>(.*?)</tr>", System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase))
        {
            var cells = new List<string>();
            foreach (System.Text.RegularExpressions.Match td in System.Text.RegularExpressions.Regex.Matches(tr.Groups[1].Value, "<t[dh]([^>]*)>(.*?)</t[dh]>", System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            {
                string txt = System.Text.RegularExpressions.Regex.Replace(td.Groups[2].Value, "<br\\s*/?>", " ", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                txt = System.Text.RegularExpressions.Regex.Replace(txt, "<[^>]+>", "");
                txt = System.Text.RegularExpressions.Regex.Replace(System.Net.WebUtility.HtmlDecode(txt).Replace('\u00A0', ' '), "\\s+", " ").Trim();
                cells.Add(txt);
                var cs = System.Text.RegularExpressions.Regex.Match(td.Groups[1].Value, "colspan\\s*=\\s*[\"']?(\\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (cs.Success && int.TryParse(cs.Groups[1].Value, out int span)) for (int k = 1; k < Math.Min(span, 50); k++) cells.Add("");
            }
            if (cells.Count > 0) rows.Add(cells.ToArray());
            if (rows.Count >= 20000) break;
        }
        return rows;
    }

    // XML 스프레드시트 2003 (.xml / 이름만 .xls)
    static List<string[]> ReadXmlSheet(string file)
    {
        var doc = System.Xml.Linq.XDocument.Parse(ReadAll(file).TrimStart('\uFEFF'));
        System.Xml.Linq.XNamespace ss = "urn:schemas-microsoft-com:office:spreadsheet";
        var rows = new List<string[]>();
        var ws = doc.Descendants(ss + "Worksheet").FirstOrDefault();
        if (ws == null) return rows;
        foreach (var row in ws.Descendants(ss + "Row"))
        {
            var list = new List<string>();
            foreach (var cell in row.Elements(ss + "Cell"))
            {
                var ix = cell.Attribute(ss + "Index");
                if (ix != null && int.TryParse(ix.Value, out int idx)) while (list.Count < idx - 1) list.Add("");
                list.Add((cell.Element(ss + "Data")?.Value ?? "").Trim());
            }
            rows.Add(list.ToArray());
            if (rows.Count >= 20000) break;
        }
        return rows;
    }

    // 탭/쉼표/세미콜론으로 나뉜 텍스트
    static List<string[]> ReadText(string file)
    {
        var lines = ReadAll(file).Replace("\r\n", "\n").Split('\n');
        var rows = new List<string[]>();
        foreach (var line in lines)
        {
            if (line.Trim() == "") continue;
            char d = line.Contains('\t') ? '\t' : line.Contains(',') ? ',' : line.Contains(';') ? ';' : '\t';
            var cells = new List<string>(); var sb = new StringBuilder(); bool q = false;
            foreach (char ch in line)
            {
                if (ch == '"') q = !q;
                else if (ch == d && !q) { cells.Add(sb.ToString().Trim()); sb.Clear(); }
                else sb.Append(ch);
            }
            cells.Add(sb.ToString().Trim());
            rows.Add(cells.ToArray());
            if (rows.Count >= 20000) break;
        }
        return rows;
    }

    static string Str(object v)
    {
        if (v == null) return "";
        if (v is double d) return d == Math.Floor(d) && Math.Abs(d) < 1e15 ? ((long)d).ToString(CultureInfo.InvariantCulture) : d.ToString("0.####", CultureInfo.InvariantCulture);
        if (v is DateTime t) return t.ToString("yyyy-MM-dd");
        return Convert.ToString(v, CultureInfo.InvariantCulture).Trim();
    }

    public static string Norm(string s) => (s ?? "").Replace(" ", "").Replace("　", "").Replace("\n", "").ToLowerInvariant();

    static readonly string[] NameKeys = { "품목명", "품명", "상품명", "제품명", "물품명", "자재명", "품목", "상품", "제품", "이름" };
    static readonly string[] UnitKeys = { "단위" };
    static readonly string[] QtyKeys = { "주문수량", "수량" };
    static readonly string[] PriceKeys = { "우리단가", "견적단가", "판매단가", "판매가", "단가" };
    static readonly string[] NoteKeys = { "비고", "메모" };
    static readonly string[] BarKeys = { "바코드", "barcode", "barcd" };
    static readonly string[] SpecKeys = { "규격", "사양" };

    static int Find(string[] cells, string[] keys, HashSet<int> used)
    {
        foreach (var k in keys)
            for (int i = 0; i < cells.Length; i++)
                if (!used.Contains(i) && Norm(cells[i]) == k) return i;
        foreach (var k in keys)
            for (int i = 0; i < cells.Length; i++)
                if (!used.Contains(i) && Norm(cells[i]).Contains(k)) return i;
        return -1;
    }

    /// 머리글 행과 열 위치를 자동으로 찾는다. (위쪽 40행 안에서)
    public static ColMap Detect(List<string[]> rows)
    {
        for (int r = 0; r < Math.Min(rows.Count, 40); r++)
        {
            var c = rows[r];
            var used = new HashSet<int>();
            var m = new ColMap { HeaderRow = r };
            m.Barcode = Find(c, BarKeys, used); if (m.Barcode >= 0) used.Add(m.Barcode);
            m.Name = Find(c, NameKeys, used); if (m.Name >= 0) used.Add(m.Name);
            m.Spec = Find(c, SpecKeys, used); if (m.Spec >= 0) used.Add(m.Spec);
            m.Unit = Find(c, UnitKeys, used); if (m.Unit >= 0) used.Add(m.Unit);
            m.Qty = Find(c, QtyKeys, used); if (m.Qty >= 0) used.Add(m.Qty);
            m.Price = Find(c, PriceKeys, used); if (m.Price >= 0) used.Add(m.Price);
            m.Note = Find(c, NoteKeys, used);
            if (m.Ok) return m;
        }
        return new ColMap();
    }

    static string Cell(string[] r, int i) => i >= 0 && i < r.Length ? r[i] : "";

    public static double Num(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return 0;
        var t = new string(s.Where(ch => char.IsDigit(ch) || ch == '.' || ch == '-').ToArray());
        return double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : 0;
    }

    static bool IsNoise(string[] r, int nameCol)
    {
        for (int i = 0; i < Math.Min(r.Length, 4); i++)
        {
            var n = Norm(r[i]);
            if (n == "합계" || n == "소계" || n == "총계" || n == "합계금액" || n.Contains("見積") || n == "견적서") return true;
        }
        var nm = Norm(Cell(r, nameCol));
        return nm == "품목" || nm == "품명" || nm == "상품명" || nm == "품목명";   // 반복되는 머리글
    }

    public static List<QItem> Parse(List<string[]> rows, ColMap m)
    {
        var list = new List<QItem>();
        for (int r = m.HeaderRow + 1; r < rows.Count; r++)
        {
            var row = rows[r];
            string name = Cell(row, m.Name).Trim();
            if (name == "" || name == "0") continue;
            if (IsNoise(row, m.Name)) continue;
            string spec = Cell(row, m.Spec).Trim();
            if (spec != "" && spec != "0" && !name.Contains(spec)) name += " " + spec;
            string unit = Cell(row, m.Unit).Trim(); if (unit == "0") unit = "";
            string note = Cell(row, m.Note).Trim(); if (note == "0") note = "";
            string bc = Cell(row, m.Barcode).Trim(); if (bc == "0") bc = "";
            list.Add(new QItem { Name = name, Unit = unit, Qty = Num(Cell(row, m.Qty)), OurPrice = Num(Cell(row, m.Price)), Note = note, Barcode = bc });
        }
        return list;
    }
}
