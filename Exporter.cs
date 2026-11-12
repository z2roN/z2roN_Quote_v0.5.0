using ClosedXML.Excel;
using System.Text;

namespace ZQuote;

/// 견적서 모양(양식). 0번이 우리 기본 양식(사용자가 준 견적서 이미지와 같은 모양), 1~3번은 타업체용으로 조금씩 다르게 만든 양식.
class FormStyle
{
    public string Label, Font, Accent, Title, LabelFill, VoucherLabel, Greeting, DateFmt, TotalLabel;
    public bool Voucher, Barcode;
    public int TotalStyle;                 // 0: 一金…원整(₩n)  1: 일금 …원정  2: ₩ n (부가세 포함)
    public string[] Footer, Heads;

    public static readonly FormStyle[] All =
    {
        new FormStyle { Label = "우리 기본 양식", Font = "맑은 고딕", Accent = "14286E", Title = "견   적   서", LabelFill = null, Voucher = true, VoucherLabel = "전표번호: ", Barcode = true,
            Greeting = "아래와 같이 견적합니다.", DateFmt = "yyyy 年 MM 月 dd 日", TotalLabel = "합계금액", TotalStyle = 0,
            Footer = new[] { "공 급 가", "세    액", "합계금액" }, Heads = new[] { "No.", "품   명", "단위", "수량", "단  가", "금  액", "비고" } },
        new FormStyle { Label = "양식 1 (검정·회색 라벨)", Font = "돋움", Accent = "000000", Title = "견  적  서", LabelFill = "F2F2F2", Voucher = false, VoucherLabel = "", Barcode = false,
            Greeting = "귀하의 요청에 따라 아래와 같이 견적합니다.", DateFmt = "yyyy년 M월 d일", TotalLabel = "총 합계", TotalStyle = 1,
            Footer = new[] { "공급가액", "부 가 세", "합    계" }, Heads = new[] { "번호", "품   목", "단위", "수량", "단가", "금액", "비고" } },
        new FormStyle { Label = "양식 2 (녹색·바탕체)", Font = "바탕", Accent = "1E5C3A", Title = "견  적  서", LabelFill = "E8F2EC", Voucher = true, VoucherLabel = "No. ", Barcode = false,
            Greeting = "다음과 같이 견적서를 제출합니다.", DateFmt = "yyyy. MM. dd.", TotalLabel = "합계금액", TotalStyle = 2,
            Footer = new[] { "공급가", "세  액", "총 금액" }, Heads = new[] { "No", "품명 / 규격", "단위", "수량", "단가", "금액", "비고" } },
        new FormStyle { Label = "양식 3 (적갈색·바코드 표시)", Font = "맑은 고딕", Accent = "7A1F1F", Title = "견 적 서", LabelFill = "F7ECEC", Voucher = false, VoucherLabel = "", Barcode = true,
            Greeting = "아래와 같이 견적을 제출합니다.", DateFmt = "yyyy-MM-dd", TotalLabel = "견적금액", TotalStyle = 0,
            Footer = new[] { "공급가액", "세  액", "합계금액" }, Heads = new[] { "순번", "품   명", "단위", "수량", "단  가", "금  액", "비고" } },
    };

    public static FormStyle Of(Company c) => c.Kind == "our" ? All[0] : All[Math.Clamp(c.Form, 0, All.Length - 1)];
}

/// 견적서 엑셀 만들기. 업체마다 한 시트(또는 한 파일).
public static class Exporter
{
    static string SafeSheet(string n, HashSet<string> used)
    {
        var bad = new[] { ':', '\\', '/', '?', '*', '[', ']' };
        string s = new string((n ?? "").Select(ch => bad.Contains(ch) ? '_' : ch).ToArray()).Trim();
        if (s == "") s = "견적";
        if (s.Length > 28) s = s.Substring(0, 28);
        string r = s; int i = 2;
        while (!used.Add(r.ToLowerInvariant())) r = s + "_" + i++;
        return r;
    }

    public static string SafeFile(string n)
    {
        var bad = Path.GetInvalidFileNameChars();
        var s = new string((n ?? "").Select(ch => bad.Contains(ch) ? '_' : ch).ToArray()).Trim().TrimEnd('.');
        return s == "" ? "견적" : s;
    }

    internal static Func<QItem, double> PriceOf(Company c) =>
        c.Kind == "our" ? (it => it.OurPrice) : (it => it.Prices.TryGetValue(c.Id, out var p) ? p : 0);

    /// 저장 폴더: 기본폴더\년도\견적의뢰업체명
    public static string OutputDir(string root, Quote q) => Path.Combine(root, q.Date.Year.ToString(), SafeFile(q.Recipient));

    /// 우리 견적 파일 1개 + 타업체 견적 파일(1개에 업체별 시트, 또는 업체마다 따로). 만든 파일 경로를 돌려줌.
    public static List<string> ExportAll(Db db, Quote q, bool includeOur, List<Company> others, bool separateOthers, string root)
    {
        string dir = OutputDir(root, q);
        Directory.CreateDirectory(dir);
        string d = q.Date.ToString("yyyyMMdd"), rc = SafeFile(q.Recipient);
        var res = new List<string>();
        if (includeOur)
        {
            string ourLabel = string.IsNullOrWhiteSpace(q.Title) || q.Title.Trim() == "견적서" ? "우리견적" : "우리" + SafeFile(q.Title);
            string p = Path.Combine(dir, $"{rc}_{ourLabel}_{d}.xlsx");
            ExportWorkbook(db, q, new[] { q.Our }, p); res.Add(p);
        }
        if (others.Count > 0)
        {
            if (separateOthers)
                foreach (var o in others)
                {
                    string p = Path.Combine(dir, $"{rc}_타업체_{SafeFile(o.Name)}_{d}.xlsx");
                    ExportWorkbook(db, q, new[] { o }, p); res.Add(p);
                }
            else
            {
                string p = Path.Combine(dir, $"{rc}_타업체_{d}.xlsx");
                ExportWorkbook(db, q, others, p); res.Add(p);
            }
        }
        return res;
    }

    /// who: 내보낼 업체들. 한 파일에 업체별 시트.
    public static void ExportWorkbook(Db db, Quote q, IEnumerable<Company> who, string path)
    {
        var keep = new List<IDisposable>();
        try
        {
            using var wb = new XLWorkbook();
            var used = new HashSet<string>();
            foreach (var c in who)
                BuildSheet(wb.Worksheets.Add(SafeSheet(c.Name, used)), q, c, db.PerPage, keep);
            wb.SaveAs(path);
        }
        finally { foreach (var d in keep) d.Dispose(); }
    }

    /// 1234567 → 일백이십삼만사천오백육십칠
    public static string KoreanAmount(long n)
    {
        if (n <= 0) return "영";
        string[] dg = { "", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구" };
        string[] un = { "", "십", "백", "천" };
        string[] big = { "", "만", "억", "조", "경" };
        var parts = new List<string>();
        int g = 0;
        while (n > 0 && g < big.Length)
        {
            int grp = (int)(n % 10000); n /= 10000;
            if (grp > 0)
            {
                var sb = new StringBuilder();
                for (int i = 3; i >= 0; i--)
                {
                    int pow = i == 3 ? 1000 : i == 2 ? 100 : i == 1 ? 10 : 1;
                    int digit = (grp / pow) % 10;
                    if (digit > 0) sb.Append(dg[digit]).Append(un[i]);
                }
                parts.Add(sb + big[g]);
            }
            g++;
        }
        parts.Reverse();
        return string.Concat(parts);
    }

    internal static string TotalText(FormStyle st, long total) => st.TotalStyle switch
    {
        0 => $"一金{KoreanAmount(total)}원整(₩{total:#,##0})",
        1 => $"일금 {KoreanAmount(total)}원정   (₩{total:#,##0})",
        _ => $"₩ {total:#,##0}   (부가세 포함)",
    };

    static XLColor Hex(string h) => XLColor.FromHtml("#" + h);

    static void Box(IXLWorksheet ws, int r1, int c1, int r2, int c2, XLBorderStyleValues st, XLColor col)
    {
        var rg = ws.Range(r1, c1, r2, c2);
        rg.Style.Border.OutsideBorder = st; rg.Style.Border.OutsideBorderColor = col;
    }

    static void BuildSheet(IXLWorksheet ws, Quote q, Company c, int perPage, List<IDisposable> keep)
    {
        var st = FormStyle.Of(c);
        var price = PriceOf(c);
        var accent = Hex(st.Accent);
        var black = XLColor.Black;
        var gray = XLColor.FromArgb(0x80, 0x80, 0x80);
        XLColor fill = st.LabelFill == null ? null : Hex(st.LabelFill);

        ws.Style.Font.FontName = st.Font;
        ws.Style.Font.FontSize = 11;
        // A 번호 / B·C 품명(앞) / D 품명(뒤)=공급자 세로글자 / E 단위=라벨 / F 수량 / G 단가 / H·I 금액 / J 비고
        double[] widths = { 5, 16, 21, 5, 8, 7, 10, 4, 8, 10 };
        for (int i = 0; i < widths.Length; i++) ws.Column(i + 1).Width = widths[i];

        int n = q.Items.Count;
        int pages = Math.Max(1, (n + perPage - 1) / perPage);
        int pageRows = 12 + perPage;
        double sum = 0; foreach (var it in q.Items) sum += it.Qty * price(it);
        long total = (long)Math.Round(sum, MidpointRounding.AwayFromZero);
        var rangeSums = new List<string>();
        for (int p = 0; p < pages; p++) rangeSums.Add($"SUM(H{p * pageRows + 9}:H{p * pageRows + 8 + perPage})");
        string grand = string.Join("+", rangeSums);

        void Put(int r, int c1, int c2, string text, double size = 11, bool bold = false, XLAlignmentHorizontalValues h = XLAlignmentHorizontalValues.Left,
                 bool wrap = false, XLAlignmentVerticalValues v = XLAlignmentVerticalValues.Center)
        {
            if (c2 > c1) ws.Range(r, c1, r, c2).Merge();
            var x = ws.Cell(r, c1);
            x.Value = text ?? "";
            x.Style.Font.FontSize = size; x.Style.Font.Bold = bold;
            x.Style.Alignment.Horizontal = h; x.Style.Alignment.Vertical = v; x.Style.Alignment.WrapText = wrap;
            if (h == XLAlignmentHorizontalValues.Left) x.Style.Alignment.Indent = 1;
        }
        var C = XLAlignmentHorizontalValues.Center; var Rt = XLAlignmentHorizontalValues.Right;

        for (int p = 0; p < pages; p++)
        {
            int r0 = p * pageRows + 1;
            int ra = r0 + 2, rb = r0 + 3, rc = r0 + 4, rd = r0 + 5, rt = r0 + 6, rh = r0 + 7, first = r0 + 8, last = r0 + 7 + perPage, blank = last + 1, f1 = blank + 1;

            // 제목 / 전표번호
            ws.Row(r0).Height = 46;
            string titleText = st.Title; double titleSize = 30;
            if (c.Kind == "our" && !string.IsNullOrWhiteSpace(q.Title))   // 우리 견적서만 제목 변경 가능 (거래명세서, 발주서 ...)
            {
                var tc = q.Title.Trim().Replace(" ", "").ToCharArray();
                titleText = tc.Length <= 3 ? string.Join("   ", tc) : string.Join(" ", tc);
                titleSize = tc.Length <= 3 ? 30 : tc.Length == 4 ? 26 : 22;
            }
            Put(r0, 2, 8, titleText, titleSize, true, C);
            ws.Range(r0, 2, r0, 8).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            ws.Range(r0, 2, r0, 8).Style.Border.BottomBorderColor = accent;
            ws.Row(r0 + 1).Height = 18;
            if (st.Voucher)
                Put(r0 + 1, 5, 10, $"{st.VoucherLabel}{q.Date:yyyyMMdd}-{q.Id:00000}" + (pages > 1 ? $"  ({p + 1}/{pages})" : ""), 10, false, Rt);

            // 머리 상자: 왼쪽 = 날짜 / 수신처 / 인사말, 오른쪽 = 공급자
            ws.Row(ra).Height = 34; ws.Row(rb).Height = 42; ws.Row(rc).Height = 36; ws.Row(rd).Height = 32;
            Put(ra, 1, 3, q.Date.ToString(st.DateFmt), 12, true);
            Put(rb, 1, 3, (q.Recipient ?? "") + " 귀하", 12);
            ws.Cell(rb, 1).Style.Font.Underline = XLFontUnderlineValues.Single;
            Put(rc, 1, 3, string.IsNullOrWhiteSpace(q.RecipientTel) ? "" : q.RecipientTel, 10);
            Put(rd, 1, 3, st.Greeting, 12, true);

            ws.Range(ra, 4, rd, 4).Merge();
            var gv = ws.Cell(ra, 4); gv.Value = "공 급 자"; gv.Style.Font.FontSize = 10; gv.Style.Alignment.Horizontal = C; gv.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            gv.Style.Alignment.TextRotation = 255; gv.Style.Alignment.WrapText = true;
            Put(ra, 5, 5, "사업자\n번호", 9, false, C, true);
            Put(rb, 5, 5, "상 호", 9, false, C);
            Put(rc, 5, 5, "주 소", 9, false, C);
            Put(rd, 5, 5, "업 태", 9, false, C);
            Put(ra, 6, 7, c.BizNo, 13, true);
            string tf = string.Join("\n", new[] { string.IsNullOrWhiteSpace(c.Tel) ? null : "Tel." + c.Tel, string.IsNullOrWhiteSpace(c.Fax) ? null : "Fax." + c.Fax }.Where(x => x != null));
            Put(ra, 8, 10, tf, 10, false, XLAlignmentHorizontalValues.Left, true);
            Put(rb, 6, 7, c.Name, 11, true);
            Put(rb, 8, 8, "성\n명", 9, false, C, true);
            Put(rb, 9, 10, c.Ceo, 11);                       // 대표자 이름 옆(오른쪽 J칸)에 도장
            Put(rc, 6, 10, c.Address, 10, false, XLAlignmentHorizontalValues.Left, true);
            Put(rd, 6, 7, c.BizType, 11);
            Put(rd, 8, 8, "종\n목", 9, false, C, true);
            Put(rd, 9, 10, c.BizItem, 10);
            if (fill != null)
                foreach (var (r1, c1, r2, c2) in new[] { (ra, 4, rd, 4), (ra, 5, rd, 5), (rb, 8, rb, 8), (rd, 8, rd, 8) })
                    ws.Range(r1, c1, r2, c2).Style.Fill.BackgroundColor = fill;
            // 선: 안쪽 가는 선, 바깥 굵은 선
            foreach (var (r1, c1, r2, c2) in new[] { (ra, 4, rd, 4), (ra, 5, ra, 5), (rb, 5, rb, 5), (rc, 5, rc, 5), (rd, 5, rd, 5), (ra, 6, ra, 7), (ra, 8, ra, 10), (rb, 6, rb, 7), (rb, 8, rb, 8), (rb, 9, rb, 10), (rc, 6, rc, 10), (rd, 6, rd, 7), (rd, 8, rd, 8), (rd, 9, rd, 10) })
                Box(ws, r1, c1, r2, c2, XLBorderStyleValues.Thin, black);
            Box(ws, ra, 1, rd, 10, XLBorderStyleValues.Medium, accent);

            // 합계금액 줄
            ws.Row(rt).Height = 34;
            Put(rt, 1, 2, st.TotalLabel, 13, true, C);
            Put(rt, 3, 10, TotalText(st, total), 14, true);
            Box(ws, rt, 1, rt, 2, XLBorderStyleValues.Thin, black); Box(ws, rt, 3, rt, 10, XLBorderStyleValues.Thin, black);
            Box(ws, rt, 1, rt, 10, XLBorderStyleValues.Medium, accent);

            // 도장 (대표자 이름 옆)
            if (c.Stamp != null && c.Stamp.Length > 0)
            {
                try
                {
                    var ms = new MemoryStream(c.Stamp); keep.Add(ms);
                    ws.AddPicture(ms).MoveTo(ws.Cell(rb, 10)).WithSize(46, 46);
                }
                catch { }
            }

            // 품목 표
            ws.Row(rh).Height = 22;
            var segs = new[] { (1, 1), (2, 4), (5, 5), (6, 6), (7, 7), (8, 9), (10, 10) };
            for (int i = 0; i < 7; i++)
            {
                Put(rh, segs[i].Item1, segs[i].Item2, st.Heads[i], 10, false, C);
                if (fill != null) ws.Range(rh, segs[i].Item1, rh, segs[i].Item2).Style.Fill.BackgroundColor = fill;
                Box(ws, rh, segs[i].Item1, rh, segs[i].Item2, XLBorderStyleValues.Thin, black);
            }
            Box(ws, rh, 1, rh, 10, XLBorderStyleValues.Medium, accent);
            for (int i = 0; i < perPage; i++)
            {
                int row = first + i, idx = p * perPage + i;
                ws.Row(row).Height = 20;
                foreach (var (a, b) in segs) { if (b > a) ws.Range(row, a, row, b).Merge(); Box(ws, row, a, row, b, XLBorderStyleValues.Thin, gray); }
                ws.Cell(row, 1).Style.Font.FontSize = 10; ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Cell(row, 5).Style.Alignment.Horizontal = C;
                foreach (int cc in new[] { 6, 7, 8 }) ws.Cell(row, cc).Style.Alignment.Horizontal = Rt;
                ws.Cell(row, 6).Style.NumberFormat.Format = "#,##0.##";
                ws.Cell(row, 7).Style.NumberFormat.Format = "#,##0";
                ws.Cell(row, 8).Style.NumberFormat.Format = "#,##0";
                if (idx < n)
                {
                    var it = q.Items[idx];
                    ws.Cell(row, 1).Value = idx + 1;
                    var nm = ws.Cell(row, 2);
                    nm.Value = (st.Barcode && !string.IsNullOrWhiteSpace(it.Barcode) ? it.Barcode.Trim() + "  " : "") + it.Name;
                    nm.Style.Alignment.ShrinkToFit = true;
                    ws.Cell(row, 5).Value = it.Unit;
                    ws.Cell(row, 6).Value = it.Qty;
                    ws.Cell(row, 7).Value = price(it);
                    ws.Cell(row, 8).FormulaA1 = $"F{row}*G{row}";
                    ws.Cell(row, 10).Value = it.Note;
                    foreach (int cc in new[] { 1, 2, 5, 6, 7, 8, 10 }) ws.Cell(row, cc).Style.Font.FontSize = 10;
                }
            }
            Box(ws, rh, 1, last, 10, XLBorderStyleValues.Medium, accent);
            ws.Row(blank).Height = 16; Box(ws, blank, 1, blank, 10, XLBorderStyleValues.Medium, accent);

            // 아래쪽: 메모(담당자·계좌) / 공급가·세액·합계
            // 우리 견적서만 하단 금액 표시를 고를 수 있음: auto(공급가·세액 자동) / manual(세액 직접 입력, 공급가 자동 맞춤) / hide(합계금액만)
            string tmode = c.Kind == "our" ? (q.TaxMode ?? "auto") : "auto";
            for (int k = 0; k < 3; k++)
            {
                int rr = f1 + k; ws.Row(rr).Height = 30;
                if (tmode == "hide" && k < 2) continue;   // 공급가·세액 칸을 비워 둠
                Put(rr, 7, 7, st.Footer[k], 10, false, XLAlignmentHorizontalValues.Left);
                ws.Range(rr, 8, rr, 10).Merge();
                var amt = ws.Cell(rr, 8); amt.Style.Font.FontSize = 13; amt.Style.Font.Bold = true; amt.Style.Alignment.Horizontal = Rt; amt.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                amt.Style.NumberFormat.Format = "#,##0"; amt.Style.Alignment.Indent = 1;
                if (k == 0) amt.FormulaA1 = tmode == "manual" ? $"H{f1 + 2}-H{f1 + 1}" : $"ROUND(H{f1 + 2}/1.1,0)";
                else if (k == 1) { if (tmode == "manual") amt.Value = Math.Round(q.TaxVal); else amt.FormulaA1 = $"H{f1 + 2}-H{f1}"; }
                else amt.FormulaA1 = grand;
                if (fill != null) ws.Cell(rr, 7).Style.Fill.BackgroundColor = fill;
                Box(ws, rr, 7, rr, 7, XLBorderStyleValues.Thin, black); Box(ws, rr, 8, rr, 10, XLBorderStyleValues.Thin, black);
            }
            ws.Range(f1, 1, f1 + 2, 6).Merge();
            var memo = ws.Cell(f1, 1); memo.Value = c.Memo ?? ""; memo.Style.Font.FontSize = 10;
            memo.Style.Alignment.WrapText = true; memo.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top; memo.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; memo.Style.Alignment.Indent = 1;
            Box(ws, f1, 1, f1 + 2, 6, XLBorderStyleValues.Thin, black);
            Box(ws, f1, 1, f1 + 2, 10, XLBorderStyleValues.Medium, accent);

            if (p < pages - 1) ws.PageSetup.AddHorizontalPageBreak(f1 + 2);
        }
        ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
        ws.PageSetup.PageOrientation = XLPageOrientation.Portrait;
        ws.PageSetup.Margins.Top = 0.5; ws.PageSetup.Margins.Bottom = 0.5; ws.PageSetup.Margins.Left = 0.5; ws.PageSetup.Margins.Right = 0.5;
        ws.PageSetup.CenterHorizontally = true;
        ws.PageSetup.Scale = 88;   // 수동 쪽 나눔이 유지되도록 '배율'만 사용
        ws.ShowGridLines = false;
    }
}
