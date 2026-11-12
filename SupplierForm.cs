using ClosedXML.Excel;

namespace ZQuote;

/// 견적업체(우리에게 견적을 주는 업체) 등록 · 수정 · 삭제 · 등록현황 조회 · 엑셀 내보내기
class SupplierForm : Form
{
    readonly Db db;
    Supplier cur = new();
    List<Supplier> rows = new();
    bool loading;
    public int Chosen;          // [이 업체로 새 견적]을 눌렀을 때 고른 업체 번호

    readonly DataGridView grid = new();
    readonly TextBox tFind = new() { Width = 260, PlaceholderText = "상호·사업자번호·담당자·취급품목 검색" };
    readonly Label lblCount = new() { AutoSize = true, ForeColor = Ui.BlueDark, Font = new Font("Malgun Gothic", 10f, FontStyle.Bold), Margin = new Padding(12, 9, 0, 0) };
    readonly Label lblMode = new() { Left = 12, Top = 8, Width = 350, Height = 24, ForeColor = Ui.BlueDark, Font = new Font("Malgun Gothic", 10f, FontStyle.Bold) };
    readonly TextBox tName = new(), tBiz = new(), tCeo = new(), tContact = new(), tTel = new(), tFax = new(), tEmail = new(), tAddr = new();
    readonly TextBox tItems = new() { Multiline = true, ScrollBars = ScrollBars.Vertical }, tMemo = new() { Multiline = true, ScrollBars = ScrollBars.Vertical };

    public SupplierForm(Db db)
    {
        this.db = db;
        Text = "견적업체 등록 · 현황 (우리에게 견적을 주는 업체)";
        Width = 1180; Height = 700; StartPosition = FormStartPosition.CenterParent; Font = new Font("Malgun Gothic", 10f); BackColor = Color.White;

        // 위쪽: 검색 / 버튼 / 개수
        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(8, 6, 8, 0), BackColor = Ui.BlueSoft, WrapContents = false };
        tFind.Margin = new Padding(3, 5, 8, 0);
        tFind.TextChanged += (_, _) => Fill(cur.Id);
        top.Controls.Add(tFind);
        top.Controls.Add(Ui.Btn("+ 새 업체", (_, _) => StartNew(), 1));
        top.Controls.Add(Ui.Btn("현황 엑셀로 내보내기", (_, _) => ExportList()));
        top.Controls.Add(lblCount);

        // 가운데: 등록현황 표
        Ui.StyleGrid(grid);
        grid.Dock = DockStyle.Fill; grid.ReadOnly = true; grid.AllowUserToAddRows = false; grid.RowHeadersVisible = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = false;
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        foreach (var (n, h, w) in new[] { ("no", "번호", 40), ("nm", "상호", 150), ("bz", "사업자번호", 100), ("ceo", "대표자", 70), ("ct", "담당자", 70), ("tel", "전화", 100), ("em", "이메일", 130), ("it", "취급품목", 170), ("dt", "등록일", 80) })
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = n, HeaderText = h, FillWeight = w });
        grid.SelectionChanged += (_, _) =>
        {
            if (loading || grid.CurrentRow == null || grid.CurrentRow.Index < 0 || grid.CurrentRow.Index >= rows.Count) return;
            LoadSup(rows[grid.CurrentRow.Index]);
        };

        // 오른쪽: 입력 칸
        var right = new Panel { Dock = DockStyle.Right, Width = 390, Padding = new Padding(10), BackColor = Color.White, AutoScroll = true };
        right.Controls.Add(lblMode);
        int y = 40;
        void Row(string label, TextBox t, int h = 0)
        {
            right.Controls.Add(new Label { Text = label, Left = 12, Top = y + 4, Width = 80 });
            t.Left = 96; t.Top = y; t.Width = 270; if (h > 0) t.Height = h;
            right.Controls.Add(t); y += (h > 0 ? h : 28) + 6;
        }
        Row("상호(업체명)", tName); Row("사업자번호", tBiz); Row("대표자", tCeo); Row("담당자", tContact); Row("전화", tTel); Row("팩스", tFax); Row("이메일", tEmail); Row("주소", tAddr);
        Row("취급품목", tItems, 56); Row("비고", tMemo, 70);
        var save = Ui.Btn("저장", (_, _) => Save(), 1); save.Left = 96; save.Top = y + 6;
        var del = Ui.Btn("삭제(관리자)", (_, _) => Delete(), 2); del.Left = 160; del.Top = y + 6;
        var use = Ui.Btn("이 업체로 새 견적 시작", (_, _) =>
        {
            if (cur == null || cur.Id == 0) { MessageBox.Show("표에서 견적업체를 먼저 고르세요."); return; }
            Chosen = cur.Id; Close();
        }); use.Left = 96; use.Top = y + 48;
        right.Controls.Add(save); right.Controls.Add(del); right.Controls.Add(use);
        right.Controls.Add(new Label { Left = 12, Top = y + 96, Width = 360, Height = 60, ForeColor = Color.Gray, Text = "우리에게 견적서를 주는 업체를 등록해 두는 곳입니다. 저장하면 왼쪽 표(등록현황)에 바로 나타나고, 검색·엑셀 내보내기를 할 수 있습니다." });

        Controls.Add(grid); Controls.Add(right); Controls.Add(top);
        Fill(0);
        StartNew();
    }

    void StartNew()
    {
        cur = new Supplier();
        Fill2(cur);
        lblMode.Text = "▶ 새 견적업체 입력 중";
        loading = true; grid.ClearSelection(); loading = false;
        tName.Focus();
    }

    void LoadSup(Supplier s)
    {
        cur = s; Fill2(s);
        lblMode.Text = $"▶ '{s.Name}' 수정 중";
    }

    void Fill2(Supplier s)
    {
        tName.Text = s.Name; tBiz.Text = s.BizNo; tCeo.Text = s.Ceo; tContact.Text = s.Contact; tTel.Text = s.Tel; tFax.Text = s.Fax; tEmail.Text = s.Email; tAddr.Text = s.Address;
        tItems.Text = (s.Items ?? "").Replace("\r\n", "\n").Replace("\n", "\r\n"); tMemo.Text = (s.Memo ?? "").Replace("\r\n", "\n").Replace("\n", "\r\n");
    }

    void Fill(int selectId)
    {
        loading = true;
        rows = db.ListSuppliers(tFind.Text.Trim());
        grid.Rows.Clear();
        int sel = -1;
        for (int i = 0; i < rows.Count; i++)
        {
            var r = rows[i];
            grid.Rows.Add(i + 1, r.Name, r.BizNo, r.Ceo, r.Contact, r.Tel, r.Email, (r.Items ?? "").Replace("\n", " ").Replace("\r", ""), r.Created);
            if (r.Id == selectId) sel = i;
        }
        int all = string.IsNullOrWhiteSpace(tFind.Text) ? rows.Count : db.ListSuppliers("").Count;
        lblCount.Text = string.IsNullOrWhiteSpace(tFind.Text) ? $"등록 {rows.Count}개" : $"검색 {rows.Count}개 / 전체 {all}개";
        if (sel >= 0) { grid.ClearSelection(); grid.Rows[sel].Selected = true; grid.CurrentCell = grid.Rows[sel].Cells["nm"]; }
        else grid.ClearSelection();
        loading = false;
    }

    void Save()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(tName.Text)) { MessageBox.Show("상호(업체명)를 입력하세요."); tName.Focus(); return; }
            var s = cur ?? new Supplier();
            s.Name = tName.Text.Trim(); s.BizNo = tBiz.Text.Trim(); s.Ceo = tCeo.Text.Trim(); s.Contact = tContact.Text.Trim(); s.Tel = tTel.Text.Trim(); s.Fax = tFax.Text.Trim();
            s.Email = tEmail.Text.Trim(); s.Address = tAddr.Text.Trim(); s.Items = tItems.Text.Replace("\r\n", "\n").Trim(); s.Memo = tMemo.Text.Replace("\r\n", "\n").Trim();
            var dup = db.ListSuppliers("").FirstOrDefault(x => x.Id != s.Id && (string.Equals(x.Name, s.Name, StringComparison.OrdinalIgnoreCase) || (s.BizNo != "" && x.BizNo == s.BizNo)));
            if (dup != null && MessageBox.Show($"'{dup.Name}' 업체가 이미 등록되어 있습니다(상호 또는 사업자번호 같음).\n그래도 저장할까요?", "중복 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            bool isNew = s.Id == 0;
            db.SaveSupplier(s);
            tFind.Text = "";
            Fill(s.Id);
            if (isNew) { StartNew(); lblMode.Text = $"✔ '{s.Name}' 등록했습니다. 다음 업체를 입력하세요."; }
            else LoadSup(rows.FirstOrDefault(x => x.Id == s.Id) ?? s);
        }
        catch (Exception e) { MessageBox.Show("저장하지 못했습니다.\n" + e.Message, "오류"); }
    }

    void Delete()
    {
        if (cur == null || cur.Id == 0) { MessageBox.Show("삭제할 업체를 표에서 고르세요."); return; }
        if (!Ui.AskAdmin(db, this, $"견적업체 '{cur.Name}' 을(를) 삭제합니다.")) return;
        try { db.DeleteSupplier(cur.Id); Fill(0); StartNew(); }
        catch (Exception e) { MessageBox.Show("삭제하지 못했습니다.\n" + e.Message, "오류"); }
    }

    void ExportList()
    {
        if (rows.Count == 0) { MessageBox.Show("내보낼 업체가 없습니다."); return; }
        using var d = new SaveFileDialog { Filter = "엑셀 (*.xlsx)|*.xlsx", FileName = $"견적업체_등록현황_{DateTime.Today:yyyyMMdd}.xlsx", Title = "견적업체 등록현황 저장" };
        if (d.ShowDialog() != DialogResult.OK) return;
        try
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("등록현황");
            string[] heads = { "번호", "상호", "사업자번호", "대표자", "담당자", "전화", "팩스", "이메일", "주소", "취급품목", "비고", "등록일" };
            for (int i = 0; i < heads.Length; i++)
            {
                var h = ws.Cell(1, i + 1); h.Value = heads[i]; h.Style.Font.Bold = true; h.Style.Fill.BackgroundColor = XLColor.FromArgb(0xEE, 0xF3, 0xFD);
            }
            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                string[] v = { (i + 1).ToString(), r.Name, r.BizNo, r.Ceo, r.Contact, r.Tel, r.Fax, r.Email, r.Address, r.Items, r.Memo, r.Created };
                for (int c = 0; c < v.Length; c++) ws.Cell(i + 2, c + 1).Value = v[c] ?? "";
            }
            ws.Cell(rows.Count + 3, 1).Value = $"총 {rows.Count}개 업체";
            ws.SheetView.FreezeRows(1);
            ws.Columns().AdjustToContents();
            wb.SaveAs(d.FileName);
            MessageBox.Show("저장했습니다.\n" + d.FileName);
        }
        catch (Exception e) { MessageBox.Show("저장하지 못했습니다. 같은 이름의 파일이 열려 있지 않은지 확인하세요.\n" + e.Message, "오류"); }
    }
}
