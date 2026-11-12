namespace ZQuote;

/// 설정: 우리 업체 / 타업체 / 기본·백업 / 바코드 상품 / 관리자
class SettingsForm : Form
{
    readonly Db db;
    public bool Restored;     // 백업에서 복원했으면 true (메인 화면을 새로 읽어야 함)
    readonly TabControl tabs = new() { Dock = DockStyle.Fill };
    Label otherMsg;           // 타업체 탭 아래쪽 안내 글
    Action otherNew;          // 타업체 탭을 '새 업체 입력' 상태로 만든다

    public SettingsForm(Db db)
    {
        this.db = db;
        Text = "설정"; Width = 960; Height = 780; StartPosition = FormStartPosition.CenterParent; Font = new Font("Malgun Gothic", 10f);
        var tOur = OurTab(); var tOther = OtherTab();
        tabs.TabPages.Add(tOur); tabs.TabPages.Add(tOther); tabs.TabPages.Add(BasicTab()); tabs.TabPages.Add(ProductTab()); tabs.TabPages.Add(AdminTab());
        Controls.Add(tabs);
        // 우리 업체가 이미 등록돼 있으면 곧바로 타업체 탭에서 시작
        if (!string.IsNullOrWhiteSpace(db.GetOur().Name)) tabs.SelectedTab = tOther;
    }

    // ---------- 우리 업체 ----------
    TabPage OurTab()
    {
        var tp = new TabPage("우리 업체");
        var ed = new CompanyEditor(false) { Left = 14, Top = 14 };
        var our = db.GetOur(); ed.LoadFrom(our);
        var msg = new Label { Left = 120, Top = 372, Width = 780, Height = 24, ForeColor = Ui.BlueDark };
        var save = Ui.Btn("저장하고 타업체 등록으로", (_, _) =>
        {
            try
            {
                var m = ed.Apply(our); if (m != null) { MessageBox.Show(m); return; }
                db.SaveCompany(our);
                otherMsg.Text = "우리 업체 정보를 저장했습니다. 이어서 타업체를 입력하고 [저장]을 누르세요. (최대 " + Db.MaxOthers + "개)";
                tabs.SelectedIndex = 1; otherNew?.Invoke();
            }
            catch (Exception e) { MessageBox.Show("저장하지 못했습니다.\n" + e.Message, "오류"); }
        }, 1);
        save.Left = 14; save.Top = 366;
        tp.Controls.Add(ed); tp.Controls.Add(save);
        tp.Controls.Add(new Label { Left = 14, Top = 410, Width = 880, Height = 44, ForeColor = Color.Gray, Text = "우리 업체는 위 견적서 양식(이미지와 같은 모양)으로 나옵니다. 도장은 대표자 이름 오른쪽 칸에 들어갑니다.\n이미 만든 견적서에는 영향이 없고, 새 견적·작성중 견적에 적용됩니다." });
        return tp;
    }

    // ---------- 타업체 ----------
    TabPage OtherTab()
    {
        var tp = new TabPage("타업체");
        var lb = new ListBox { Left = 10, Top = 10, Width = 200, Height = 270 };
        var ed = new CompanyEditor(true) { Left = 225, Top = 10 };
        var info = new Label { Left = 10, Top = 340, Width = 205, Height = 90, ForeColor = Color.Gray };
        var mode = new Label { Left = 10, Top = 296, Width = 205, Height = 40, ForeColor = Ui.BlueDark };
        otherMsg = new Label { Left = 225, Top = 468, Width = 690, Height = 40, ForeColor = Ui.BlueDark };
        Company cur = null;                 // null 이면 '새 타업체 입력 중'
        List<Company> list = null;
        Button save = null;

        int NextForm()
        {
            var used = new HashSet<int>(list.Select(c => c.Form));
            for (int i = 1; i < FormStyle.All.Length; i++) if (!used.Contains(i)) return i;
            return 1;
        }
        void Fill(int sel)
        {
            list = db.GetOthers();
            lb.Items.Clear();
            foreach (var c in list) lb.Items.Add(c.Name);
            info.Text = $"등록 {list.Count}개 / 최대 {Db.MaxOthers}개\n기본 인상률 {db.DefaultMarkup:0.##}%\n(우리 견적서 제외)";
            if (sel >= 0 && sel < list.Count) lb.SelectedIndex = sel;
        }
        void StartNew()
        {
            if (list.Count >= Db.MaxOthers)
            {
                cur = null; lb.SelectedIndex = list.Count > 0 ? 0 : -1;
                otherMsg.Text = $"타업체는 최대 {Db.MaxOthers}개까지 등록됩니다. 고치려면 왼쪽 목록에서 업체를 고르세요.";
                return;
            }
            cur = null; lb.SelectedIndex = -1;
            var d = new Company { Kind = "other", Form = NextForm() };
            ed.LoadFrom(d);
            save.Enabled = true;
            mode.Text = "▶ 새 타업체 입력 중\n입력한 뒤 [저장]";
        }
        lb.SelectedIndexChanged += (_, _) =>
        {
            if (lb.SelectedIndex < 0 || lb.SelectedIndex >= list.Count) return;
            cur = list[lb.SelectedIndex]; ed.LoadFrom(cur); save.Enabled = true;
            mode.Text = $"▶ '{cur.Name}' 수정 중\n고친 뒤 [저장]";
        };
        var add = Ui.Btn("+ 새 타업체", (_, _) => { otherMsg.Text = ""; StartNew(); });
        add.Left = 10; add.Top = 436;
        var del = Ui.Btn("삭제(관리자)", (_, _) =>
        {
            if (cur == null) { MessageBox.Show("삭제할 업체를 왼쪽 목록에서 고르세요."); return; }
            if (!Ui.AskAdmin(db, this, $"'{cur.Name}' 업체를 삭제합니다. 이미 만든 견적서에는 영향이 없습니다.")) return;
            try { db.DeleteCompany(cur.Id); Fill(-1); cur = null; StartNew(); otherMsg.Text = "삭제했습니다."; }
            catch (Exception e) { MessageBox.Show("삭제하지 못했습니다.\n" + e.Message, "오류"); }
        }, 2);
        del.Left = 110; del.Top = 436;
        save = Ui.Btn("저장", (_, _) =>
        {
            try
            {
                var c = cur ?? new Company { Kind = "other" };
                var m = ed.Apply(c); if (m != null) { MessageBox.Show(m); return; }
                bool isNew = c.Id == 0;
                db.SaveCompany(c);
                int idx = list.FindIndex(x => x.Id == c.Id);
                Fill(-1);
                otherMsg.Text = $"'{c.Name}' 저장했습니다. ({list.Count}/{Db.MaxOthers})" + (isNew && list.Count < Db.MaxOthers ? "  다음 타업체를 이어서 입력하세요." : "");
                if (isNew) StartNew(); else lb.SelectedIndex = Math.Max(0, list.FindIndex(x => x.Id == c.Id));
            }
            catch (Exception e) { MessageBox.Show("저장하지 못했습니다.\n" + e.Message, "오류"); }
        }, 1);
        save.Left = 225; save.Top = 430;
        tp.Controls.AddRange(new Control[] { lb, ed, info, mode, otherMsg, add, del, save });
        Fill(-1);
        otherNew = StartNew;
        StartNew();
        return tp;
    }

    // ---------- 기본 · 백업 ----------
    TabPage BasicTab()
    {
        var tp = new TabPage("기본·백업");
        int y = 14;
        Label L(string t) => new Label { Text = t, Left = 14, Top = y + 4, Width = 190 };
        var nMk = new TextBox { Left = 210, Top = y, Width = 70, Text = db.DefaultMarkup.ToString("0.##") };
        tp.Controls.Add(L("기본 인상률 (우리보다 %)")); tp.Controls.Add(nMk);
        tp.Controls.Add(new Label { Text = "타업체 단가 = 우리 단가 × (1 + 인상률). 업체별로 따로 정하지 않으면 이 값을 씁니다.", Left = 290, Top = y + 5, Width = 600, ForeColor = Color.Gray }); y += 38;
        var cUnit = new ComboBox { Left = 210, Top = y, Width = 80, DropDownStyle = ComboBoxStyle.DropDownList }; cUnit.Items.AddRange(new object[] { "1", "10", "100" }); cUnit.Text = db.RoundUnit.ToString();
        var cMode = new ComboBox { Left = 300, Top = y, Width = 90, DropDownStyle = ComboBoxStyle.DropDownList }; cMode.Items.AddRange(new object[] { "올림", "반올림", "버림" }); cMode.Text = db.RoundMode;
        tp.Controls.Add(L("단가 끝자리 처리 (원)")); tp.Controls.Add(cUnit); tp.Controls.Add(cMode); y += 38;
        var nPage = new TextBox { Left = 210, Top = y, Width = 70, Text = db.PerPage.ToString() };
        tp.Controls.Add(L("한 쪽에 들어갈 품목 수")); tp.Controls.Add(nPage);
        tp.Controls.Add(new Label { Text = "10~40 (기본 23, 우리 견적서 양식 기준)", Left = 290, Top = y + 5, Width = 300, ForeColor = Color.Gray }); y += 38;
        var cTax = new ComboBox { Left = 210, Top = y, Width = 330, DropDownStyle = ComboBoxStyle.DropDownList };
        cTax.Items.AddRange(new object[] { "공급가·세액·합계 표시 (세액 자동)", "세액 직접 수정 (공급가 자동 맞춤)", "합계금액만 표시" });
        cTax.SelectedIndex = db.DefaultTaxMode == "manual" ? 1 : db.DefaultTaxMode == "hide" ? 2 : 0;
        tp.Controls.Add(L("견적서 하단 금액 기본값")); tp.Controls.Add(cTax);
        tp.Controls.Add(new Label { Text = "우리 견적서만 적용. 견적마다 화면 아래에서 바꿀 수 있습니다.", Left = 550, Top = y + 5, Width = 360, ForeColor = Color.Gray }); y += 38;
        var cTitle = new ComboBox { Left = 210, Top = y, Width = 160, DropDownStyle = ComboBoxStyle.DropDown, Text = db.DefaultTitle };
        cTitle.Items.AddRange(new object[] { "견적서", "거래명세서", "발주서", "납품서", "청구서", "주문서" });
        tp.Controls.Add(L("우리 견적서 제목 기본값")); tp.Controls.Add(cTitle);
        tp.Controls.Add(new Label { Text = "우리 견적서만 제목을 바꿀 수 있습니다. 견적마다 화면 위쪽에서도 바꿀 수 있습니다.", Left = 380, Top = y + 5, Width = 520, ForeColor = Color.Gray }); y += 38;
        var tApp = new TextBox { Left = 210, Top = y, Width = 300, Text = db.AppName };
        tp.Controls.Add(L("프로그램 이름")); tp.Controls.Add(tApp);
        tp.Controls.Add(new Label { Text = "창 제목과 맨 위 띠에 표시됩니다. (버전은 자동 표시)", Left = 520, Top = y + 5, Width = 380, ForeColor = Color.Gray }); y += 38;
        var tExp = new TextBox { Left = 210, Top = y, Width = 480, Text = db.ExportDir };
        var bExp = Ui.Btn("폴더 선택", (_, _) =>
        {
            using var d = new FolderBrowserDialog { SelectedPath = tExp.Text, Description = "견적서 저장 기본 폴더 선택" };
            if (d.ShowDialog() == DialogResult.OK) tExp.Text = d.SelectedPath;
        });
        bExp.Left = 700; bExp.Top = y - 3;
        tp.Controls.Add(L("견적서 저장 폴더")); tp.Controls.Add(tExp); tp.Controls.Add(bExp);
        tp.Controls.Add(new Label { Text = "내려받으면 이 폴더 아래에 [년도]\\[견적의뢰업체명]\\ 폴더가 자동으로 만들어집니다.", Left = 210, Top = y + 32, Width = 700, ForeColor = Color.Gray }); y += 62;
        var cbAuto = new CheckBox { Left = 14, Top = y, Width = 400, Text = "견적 '완료' 할 때 자동으로 백업", Checked = db.AutoBackup }; tp.Controls.Add(cbAuto); y += 36;
        var tDir = new TextBox { Left = 210, Top = y, Width = 480, Text = db.BackupDir };
        var bDir = Ui.Btn("폴더 선택", (_, _) =>
        {
            using var d = new FolderBrowserDialog { SelectedPath = tDir.Text, Description = "백업 폴더 선택" };
            if (d.ShowDialog() == DialogResult.OK) { tDir.Text = d.SelectedPath; }
        });
        bDir.Left = 700; bDir.Top = y - 3;
        tp.Controls.Add(L("백업 폴더")); tp.Controls.Add(tDir); tp.Controls.Add(bDir); y += 42;

        var files = new ListBox { Left = 14, Top = y + 44, Width = 676, Height = 120 };
        void FillFiles() { files.Items.Clear(); foreach (var f in Db.BackupFiles(tDir.Text)) files.Items.Add($"{f.Name}    ({f.LastWriteTime:yyyy-MM-dd HH:mm}, {f.Length / 1024:N0} KB)"); }
        Func<string> SelFile = () => files.SelectedIndex >= 0 ? Db.BackupFiles(tDir.Text)[files.SelectedIndex].FullName : null;
        Action saveAll = () =>
        {
            if (!double.TryParse(nMk.Text, out var mk) || mk < -50 || mk > 500) throw new InvalidOperationException("기본 인상률을 숫자로 입력하세요.");
            if (!int.TryParse(nPage.Text, out var pg) || pg < 10 || pg > 40) throw new InvalidOperationException("한 쪽 품목 수는 10~40 으로 입력하세요.");
            db.DefaultMarkup = mk; db.RoundUnit = int.Parse(cUnit.Text); db.RoundMode = cMode.Text; db.PerPage = pg; db.DefaultTitle = cTitle.Text; db.DefaultTaxMode = cTax.SelectedIndex == 1 ? "manual" : cTax.SelectedIndex == 2 ? "hide" : "auto"; db.AppName = tApp.Text; db.ExportDir = tExp.Text.Trim(); db.AutoBackup = cbAuto.Checked; db.BackupDir = tDir.Text.Trim();
        };
        var bSave = Ui.Btn("설정 저장", (_, _) => { try { saveAll(); FillFiles(); MessageBox.Show("저장했습니다."); } catch (Exception e) { MessageBox.Show(e.Message); } }, 1);
        bSave.Left = 14; bSave.Top = y;
        var bNow = Ui.Btn("지금 백업", (_, _) => { try { saveAll(); var f = db.Backup(); FillFiles(); MessageBox.Show("백업했습니다.\n" + f); } catch (Exception e) { MessageBox.Show(e.Message); } });
        bNow.Left = 120; bNow.Top = y;
        var bRes = Ui.Btn("선택한 백업으로 복원(관리자)", (_, _) =>
        {
            var f = SelFile(); if (f == null) { MessageBox.Show("목록에서 백업 파일을 고르세요."); return; }
            if (!Ui.AskAdmin(db, this, "현재 자료가 선택한 백업 내용으로 바뀝니다. (현재 자료는 .before_restore 로 보관)")) return;
            try { db.Restore(f); Restored = true; MessageBox.Show("복원했습니다. 설정 창을 닫으면 화면이 새로 고쳐집니다."); Close(); } catch (Exception e) { MessageBox.Show(e.Message); }
        });
        bRes.Left = 14; bRes.Top = y + 212 + 44 - 44; bRes.Top = files.Bottom + 6;
        var bDel = Ui.Btn("선택한 백업 삭제(관리자)", (_, _) =>
        {
            var f = SelFile(); if (f == null) { MessageBox.Show("목록에서 백업 파일을 고르세요."); return; }
            if (!Ui.AskAdmin(db, this, "백업 파일을 삭제합니다.")) return;
            try { File.Delete(f); db.Audit("백업삭제", Path.GetFileName(f)); FillFiles(); } catch (Exception e) { MessageBox.Show(e.Message); }
        }, 2);
        bDel.Left = 250; bDel.Top = files.Bottom + 6;
        tp.Controls.Add(new Label { Text = "백업 목록", Left = 14, Top = y + 48 - 30, Width = 200 });
        tp.Controls.AddRange(new Control[] { bSave, bNow, files, bRes, bDel });
        FillFiles();
        return tp;
    }

    // ---------- 바코드 상품 ----------
    TabPage ProductTab()
    {
        var tp = new TabPage("바코드 상품");
        var tFind = new TextBox { Left = 14, Top = 12, Width = 240 };
        var grid = new DataGridView { Left = 14, Top = 48, Width = 890, Height = 440, AllowUserToAddRows = false, AllowUserToDeleteRows = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
        Ui.StyleGrid(grid); grid.RowHeadersVisible = false;
        grid.Columns.Add("bc", "바코드"); grid.Columns.Add("nm", "품목"); grid.Columns.Add("un", "단위"); grid.Columns.Add("pr", "우리 단가"); grid.Columns.Add("nt", "비고"); grid.Columns.Add("up", "수정일");
        grid.Columns["bc"].ReadOnly = true; grid.Columns["up"].ReadOnly = true;
        grid.Columns["nm"].FillWeight = 200; grid.Columns["un"].FillWeight = 40; grid.Columns["pr"].FillWeight = 70; grid.Columns["pr"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        bool loading = false;
        void Fill()
        {
            loading = true; grid.Rows.Clear();
            foreach (var p in db.ListProducts(tFind.Text.Trim())) grid.Rows.Add(p.Barcode, p.Name, p.Unit, p.Price.ToString("#,##0"), p.Note, p.Updated);
            loading = false;
        }
        tFind.TextChanged += (_, _) => Fill();
        grid.CellEndEdit += (_, e) =>
        {
            if (loading || e.RowIndex < 0) return;
            var r = grid.Rows[e.RowIndex];
            db.SaveProduct(new Product { Barcode = Convert.ToString(r.Cells["bc"].Value), Name = Convert.ToString(r.Cells["nm"].Value) ?? "", Unit = Convert.ToString(r.Cells["un"].Value) ?? "", Price = RequestReader.Num(Convert.ToString(r.Cells["pr"].Value)), Note = Convert.ToString(r.Cells["nt"].Value) ?? "" });
        };
        var add = Ui.Btn("+ 바코드 추가", (_, _) =>
        {
            var b = Ui.Ask(this, "바코드 추가", "바코드를 입력(스캔)하세요.");
            if (string.IsNullOrWhiteSpace(b)) return;
            b = b.Trim();
            if (db.GetProduct(b) == null) db.SaveProduct(new Product { Barcode = b });
            tFind.Text = b; Fill();
        });
        add.Left = 270; add.Top = 8;
        var del = Ui.Btn("선택 삭제(관리자)", (_, _) =>
        {
            if (grid.CurrentRow == null) return;
            string b = Convert.ToString(grid.CurrentRow.Cells["bc"].Value);
            if (!Ui.AskAdmin(db, this, $"바코드 {b} 를 상품 목록에서 삭제합니다.")) return;
            db.DeleteProduct(b); Fill();
        }, 2);
        del.Left = 410; del.Top = 8;
        tp.Controls.Add(new Label { Text = "찾기", Left = 14 + 250 + 4, Top = 0, Width = 1 });
        tp.Controls.AddRange(new Control[] { tFind, add, del, grid });
        tp.Controls.Add(new Label { Text = "견적 품목에 바코드가 있으면 저장할 때 자동으로 여기에 수집됩니다. 칸을 직접 고쳐도 됩니다.", Left = 560, Top = 18, Width = 350, Height = 40, ForeColor = Color.Gray });
        Fill();
        return tp;
    }

    // ---------- 관리자 ----------
    TabPage AdminTab()
    {
        var tp = new TabPage("관리자");
        var info = new Label { Left = 14, Top = 14, Width = 900, Height = 66, Text = "관리자 비밀번호는 선택 사항입니다. 만들어 두지 않으면 누구나 삭제·복원할 수 있고(확인창만 나옴),\n만들어 두면 완료된 견적 삭제, 타업체·백업·바코드 삭제, 백업 복원에 비밀번호가 필요합니다.\n화면 왼쪽 아래 [로그아웃]을 누르면 관리자 로그인이 풀립니다." };
        Button chg = null, rm = null;
        void RefreshAdmin() { chg.Text = db.HasAdmin() ? "관리자 비밀번호 변경" : "관리자 비밀번호 만들기"; rm.Visible = db.HasAdmin(); }
        chg = Ui.Btn("", (_, _) => { using var f = new AdminSetupForm(db, db.HasAdmin()); f.ShowDialog(this); Ui.AdminLoggedIn = false; RefreshAdmin(); }, 1);
        chg.Left = 14; chg.Top = 90;
        rm = Ui.Btn("관리자 비밀번호 없애기", (_, _) =>
        {
            Ui.AdminLoggedIn = false;
            if (!Ui.AskAdmin(db, this, "관리자 비밀번호를 없앱니다. 이후에는 확인창만 나옵니다.")) return;
            db.RemoveAdmin(); Ui.AdminLoggedIn = false; RefreshAdmin();
        }, 2);
        rm.Left = 200; rm.Top = 90;
        RefreshAdmin();
        var lb = new ListBox { Left = 14, Top = 160, Width = 900, Height = 380 };
        tp.Controls.Add(new Label { Text = "삭제·복원 등 기록 (최근 200건)", Left = 14, Top = 136, Width = 400 });
        using (var c = db.Conn().CreateCommand())
        {
            c.CommandText = "SELECT at,action,detail FROM audit ORDER BY id DESC LIMIT 200";
            using var r = c.ExecuteReader();
            while (r.Read()) lb.Items.Add($"{r.GetString(0)}   {r.GetString(1)}   {(r.IsDBNull(2) ? "" : r.GetString(2))}");
        }
        tp.Controls.AddRange(new Control[] { info, chg, rm, lb });
        return tp;
    }
}
