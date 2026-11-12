using System.Net;
using System.Text;

namespace ZQuote;

/// 견적서 미리보기: 엑셀로 내려받을 모양을 HTML 로 비슷하게 그려 보여줍니다. (실제 엑셀과 글꼴·줄 높이가 조금 다를 수 있음)
static class Previewer
{
    static string E(string s) => WebUtility.HtmlEncode(s ?? "").Replace("\n", "<br>");

    public static string Build(Quote q, Company c, int perPage)
    {
        var st = FormStyle.Of(c);
        var price = Exporter.PriceOf(c);
        string ac = "#" + st.Accent, fill = st.LabelFill == null ? "#fff" : "#" + st.LabelFill;
        int n = q.Items.Count, pages = Math.Max(1, (n + perPage - 1) / perPage);
        double sum = 0; foreach (var it in q.Items) sum += it.Qty * price(it);
        long total = (long)Math.Round(sum, MidpointRounding.AwayFromZero);
        string tmode = c.Kind == "our" ? (q.TaxMode ?? "auto") : "auto";
        long supply = tmode == "manual" ? total - (long)Math.Round(q.TaxVal) : (long)Math.Round(total / 1.1, MidpointRounding.AwayFromZero);
        long tax = tmode == "manual" ? (long)Math.Round(q.TaxVal) : total - supply;
        string titleText = st.Title;
        if (c.Kind == "our" && !string.IsNullOrWhiteSpace(q.Title))
        {
            var tc = q.Title.Trim().Replace(" ", "").ToCharArray();
            titleText = tc.Length <= 3 ? string.Join("   ", tc) : string.Join(" ", tc);
        }
        string stamp = c.Stamp != null && c.Stamp.Length > 0 ? "<img class=st src='data:image/png;base64," + Convert.ToBase64String(c.Stamp) + "'>" : "";

        var sb = new StringBuilder();
        sb.Append("<html><head><meta http-equiv='X-UA-Compatible' content='IE=edge'><meta charset='utf-8'><style>");
        sb.Append($"body{{background:#dfe5f0;margin:0;font-family:'{st.Font}','Malgun Gothic',sans-serif;font-size:12px;color:#000}}");
        sb.Append($".pg{{width:680px;background:#fff;margin:14px auto;padding:22px 24px;border:1px solid #b8c4dc}}");
        sb.Append("table{border-collapse:collapse;width:100%;table-layout:fixed}td{border:1px solid #000;padding:3px 5px;overflow:hidden}");
        sb.Append($".t{{text-align:center;font-size:26px;font-weight:bold;border:0;border-bottom:3px solid {ac};padding:6px}}");
        sb.Append($".box{{border:2px solid {ac}}}.lb{{background:{fill};text-align:center;font-size:10px}}.r{{text-align:right}}.c{{text-align:center}}");
        sb.Append(".g{border-color:#808080}.st{position:absolute;right:2px;top:2px;width:44px;height:44px}.rel{position:relative}.nb{border:0}");
        sb.Append("@media print{body{background:#fff}.pg{border:0;margin:0 auto;page-break-after:always}}</style></head><body>");

        for (int p = 0; p < pages; p++)
        {
            sb.Append("<div class=pg>");
            sb.Append($"<table><tr><td class='t' style='color:{ac}'>{E(titleText)}</td></tr></table>");
            if (st.Voucher)
                sb.Append($"<div class=r style='font-size:10px;margin:2px 0'>{E(st.VoucherLabel)}{q.Date:yyyyMMdd}-{q.Id:00000}{(pages > 1 ? $"  ({p + 1}/{pages})" : "")}</div>");
            else sb.Append("<div style='height:6px'></div>");
            string tf = string.Join("<br>", new[] { string.IsNullOrWhiteSpace(c.Tel) ? null : "Tel." + E(c.Tel), string.IsNullOrWhiteSpace(c.Fax) ? null : "Fax." + E(c.Fax) }.Where(x => x != null));
            // 머리 상자
            sb.Append($"<table class=box><colgroup><col width=240><col width=26><col width=52><col width=50><col width=40><col width=40><col width=40><col width=40><col width=40></colgroup>");
            sb.Append($"<tr style='height:34px'><td rowspan=4 style='border:0;vertical-align:top;font-weight:bold'>{E(q.Date.ToString(st.DateFmt))}<br><br><div style='text-decoration:underline;font-weight:normal;font-size:13px'>{E(q.Recipient)} 귀하</div><div style='font-weight:normal;font-size:10px;margin-top:8px'>{E(q.RecipientTel)}</div><div style='margin-top:10px'>{E(st.Greeting)}</div></td>");
            sb.Append($"<td rowspan=4 class=lb style='width:26px'>공<br>급<br>자</td><td class=lb>사업자<br>번호</td><td colspan=2 style='font-weight:bold;font-size:13px'>{E(c.BizNo)}</td><td colspan=4 style='font-size:10px'>{tf}</td></tr>");
            sb.Append($"<tr style='height:40px'><td class=lb>상 호</td><td colspan=2 style='font-weight:bold'>{E(c.Name)}</td><td class=lb>성<br>명</td><td colspan=3 class=rel>{E(c.Ceo)}{stamp}</td></tr>");
            sb.Append($"<tr style='height:34px'><td class=lb>주 소</td><td colspan=6 style='font-size:10px'>{E(c.Address)}</td></tr>");
            sb.Append($"<tr style='height:30px'><td class=lb>업 태</td><td colspan=2>{E(c.BizType)}</td><td class=lb>종<br>목</td><td colspan=3 style='font-size:10px'>{E(c.BizItem)}</td></tr></table>");
            // 합계금액
            sb.Append($"<table class=box style='margin-top:-1px'><tr style='height:34px'><td class='lb' style='width:90px;font-size:13px;font-weight:bold'>{E(st.TotalLabel)}</td><td style='font-size:14px;font-weight:bold'>{E(Exporter.TotalText(st, total))}</td></tr></table>");
            // 품목
            sb.Append("<table class=box style='margin-top:-1px'><colgroup><col width=34><col width=215><col width=40><col width=48><col width=70><col width=80><col width=70></colgroup>");
            sb.Append("<tr style='height:22px'>");
            for (int i = 0; i < 7; i++) sb.Append($"<td class=lb style='font-size:10px'>{E(st.Heads[i])}</td>");
            sb.Append("</tr>");
            for (int i = 0; i < perPage; i++)
            {
                int idx = p * perPage + i;
                sb.Append("<tr style='height:20px'>");
                if (idx < n)
                {
                    var it = q.Items[idx];
                    string nm = (st.Barcode && !string.IsNullOrWhiteSpace(it.Barcode) ? it.Barcode.Trim() + "  " : "") + it.Name;
                    sb.Append($"<td class=g style='font-size:10px'>{idx + 1}</td><td class=g style='font-size:10px;white-space:nowrap'>{E(nm)}</td><td class='g c' style='font-size:10px'>{E(it.Unit)}</td>");
                    sb.Append($"<td class='g r' style='font-size:10px'>{it.Qty:#,##0.##}</td><td class='g r' style='font-size:10px'>{Ui.Won(price(it))}</td><td class='g r' style='font-size:10px'>{Ui.Won(it.Qty * price(it))}</td><td class=g style='font-size:10px'>{E(it.Note)}</td>");
                }
                else sb.Append("<td class=g>&nbsp;</td><td class=g></td><td class=g></td><td class=g></td><td class=g></td><td class=g></td><td class=g></td>");
                sb.Append("</tr>");
            }
            sb.Append("</table>");
            // 아래쪽
            sb.Append("<table class=box style='margin-top:-1px'><colgroup><col><col width=70><col width=130></colgroup>");
            sb.Append($"<tr style='height:30px'><td rowspan=3 style='vertical-align:top;font-size:10px'>{E(c.Memo)}</td>");
            for (int k = 0; k < 3; k++)
            {
                if (k > 0) sb.Append("<tr style='height:30px'>");
                if (tmode == "hide" && k < 2) { sb.Append("<td style='border:0'></td><td style='border:0'></td></tr>"); continue; }
                long v = k == 0 ? supply : k == 1 ? tax : total;
                sb.Append($"<td class=lb style='font-size:10px;text-align:left'>{E(st.Footer[k])}</td><td class=r style='font-size:13px;font-weight:bold'>{v:#,##0}</td></tr>");
            }
            sb.Append("</table></div>");
        }
        sb.Append("</body></html>");
        return sb.ToString();
    }
}

/// 미리보기 창: 우리 견적서 + 타업체 견적서를 탭으로 보여줌
class PreviewForm : Form
{
    readonly TabControl tabs = new() { Dock = DockStyle.Fill };

    public PreviewForm(Quote q, int perPage)
    {
        Text = "견적서 미리보기 — " + q.Recipient;
        Width = 820; Height = 900; StartPosition = FormStartPosition.CenterParent; Font = new Font("Malgun Gothic", 9.5f);
        var all = new List<Company> { q.Our }; all.AddRange(q.Others);
        foreach (var c in all)
        {
            var tp = new TabPage((c.Kind == "our" ? "[우리] " : "[타업체] ") + c.Name);
            var wb = new WebBrowser { Dock = DockStyle.Fill, ScriptErrorsSuppressed = true, AllowWebBrowserDrop = false, IsWebBrowserContextMenuEnabled = false };
            wb.DocumentText = Previewer.Build(q, c, perPage);
            tp.Controls.Add(wb); tabs.TabPages.Add(tp);
        }
        var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 48, Padding = new Padding(8), BackColor = Ui.BlueSoft };
        bar.Controls.Add(Ui.Btn("인쇄", (_, _) => { if (tabs.SelectedTab?.Controls[0] is WebBrowser w) w.ShowPrintDialog(); }, 1));
        bar.Controls.Add(Ui.Btn("닫기", (_, _) => Close()));
        bar.Controls.Add(new Label { Text = "화면용 미리보기입니다. 실제 엑셀과 글꼴·줄 높이가 조금 다를 수 있습니다.", AutoSize = true, ForeColor = Color.Gray, Margin = new Padding(10, 9, 0, 0) });
        Controls.Add(tabs); Controls.Add(bar);
    }
}
