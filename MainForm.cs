using System.Globalization;

namespace ZQuote;

public class MainForm : Form
{
    readonly Db db;
    Quote cur;
    bool dirty, suppress, askedCompletedEdit;
    List<Company> regOthers = new();

    readonly TabControl tabs = new() { Dock = DockStyle.Fill };
    readonly DataGridView grid = new();
    readonly DataGridView listGrid = new();
    readonly TextBox txtRecipient = new() { Width = 260 };
    readonly TextBox txtRecipientTel = new() { Width = 220, PlaceholderText = "수신처 전화/팩스 (선택)" };
    readonly ComboBox cbView = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 230 };
    readonly Label lblAdmin = new() { AutoSize = true, ForeColor = Ui.BlueDark, Margin = new Padding(14, 12, 0, 0) };
    Panel head;
    bool viewing;
    readonly DateTimePicker dtp = new() { Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Width = 120 };
    readonly Label lblState = new() { AutoSize = true, Font = new Font("Malgun Gothic", 10f, FontStyle.Bold) };
    readonly Label lblTotals = new() { Dock = DockStyle.Bottom, Height = 54, Padding = new Padding(10, 6, 10, 0), ForeColor = Ui.BlueDark, BackColor = Ui.BlueSoft };
    readonly CheckedListBox clbOthers = new() { Dock = DockStyle.Fill, CheckOnClick = true, IntegralHeight = false };
    readonly ComboBox cbTax = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 330 };
    readonly TextBox txtTax = new() { Width = 110, TextAlign = HorizontalAlignment.Right, Enabled = false };
    readonly Label lblSupply = new() { AutoSize = true, ForeColor = Ui.BlueDark, Margin = new Padding(12, 9, 0, 0) };
    static readonly string[] TaxModes = { "auto", "manual", "hide" };
    double ourTotal;
    readonly ComboBox cbSup = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 250 };
    readonly ComboBox cbTitle = new() { DropDownStyle = ComboBoxStyle.DropDown, Width = 150 };
    List<Supplier> suppliers = new();
    readonly TextBox txtScan = new() { Width = 230, PlaceholderText = "여기에 바코드 스캔 후 Enter" };
    readonly TextBox txtListFind = new() { Width = 220, PlaceholderText = "수신처 / 상태(작성중·완료) 검색" };
    readonly Label lblStatus = new() { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Ui.BlueDark, Padding = new Padding(10, 0, 0, 0) };

    public MainForm(Db db)
    {
        this.db = db;
        ApplyTitle();
        try { Icon = Icon.ExtractAssociatedIcon(Environment.ProcessPath); } catch { }
        Width = 1360; Height = 820; StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Malgun Gothic", 10f); BackColor = Color.White;
        BuildUi();
        UpdateAdminLabel();
        HookDrop(this);   // 창 어디에 끌어다 놓아도 견적의뢰 엑셀이 올라가도록 모든 칸에 연결
        Load += (_, _) => { NewQuote(false); CheckOur(); };
        var tm = new System.Windows.Forms.Timer { Interval = 800 };
        tm.Tick += (_, _) => UpdateAdminLabel();
        tm.Start();
        FormClosing += (_, e) =>
        {
            if (dirty && MessageBox.Show("저장하지 않은 내용이 있습니다. 그래도 닫을까요?", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) { e.Cancel = true; return; }
            db.Dispose();
        };
        FormClosed += (_, _) => Environment.Exit(0);
    }

    // ---------------- 우리 견적서 하단 금액 설정 ----------------
    void LoadTaxUi(string mode, double val)
    {
        suppress = true;
        cbTax.SelectedIndex = Math.Max(0, Array.IndexOf(TaxModes, mode ?? "auto"));
        txtTax.Text = val > 0 ? Ui.Won(val) : "";
        txtTax.Enabled = cbTax.SelectedIndex == 1;
        suppress = false;
        UpdateTaxLabel();
    }

    // 견적업체 선택 목록
    void LoadSuppliers(int selectId)
    {
        suppliers = db.ListSuppliers("");
        bool old = suppress; suppress = true;
        cbSup.Items.Clear(); cbSup.Items.Add("(지정 안 함)");
        foreach (var s in suppliers) cbSup.Items.Add(s.Name);
        int i = suppliers.FindIndex(x => x.Id == selectId);
        cbSup.SelectedIndex = i >= 0 ? i + 1 : 0;
        suppress = old;
    }

    void ApplySupplier(Supplier s)
    {
        txtRecipient.Text = s.Name;
        var tf = new List<string>();
        if (!string.IsNullOrWhiteSpace(s.Tel)) tf.Add("Tel." + s.Tel);
        if (!string.IsNullOrWhiteSpace(s.Fax)) tf.Add("Fax." + s.Fax);
        txtRecipientTel.Text = string.Join(" / ", tf);
        lblStatus.Text = $"견적업체 '{s.Name}' 로 진행합니다. 견적의뢰 엑셀을 올려 주세요.";
    }

    int SelectedSupplierId() => cbSup.SelectedIndex > 0 && cbSup.SelectedIndex <= suppliers.Count ? suppliers[cbSup.SelectedIndex - 1].Id : 0;

    void SyncTax() { cur.SupplierId = SelectedSupplierId(); cur.Title = cbTitle.Text.Trim(); cur.TaxMode = TaxModes[Math.Max(0, cbTax.SelectedIndex)]; cur.TaxVal = RequestReader.Num(txtTax.Text); }

    void UpdateTaxLabel()
    {
        if (cbTax.SelectedIndex < 0) return;
        double total = Math.Round(ourTotal);
        if (cbTax.SelectedIndex == 2) lblSupply.Text = "견적서에는 합계금액 " + Ui.Won(total) + "원만 나옵니다.";
        else if (cbTax.SelectedIndex == 0) { double sup = Math.Round(total / 1.1); lblSupply.Text = $"공급가 {Ui.Won(sup)}  /  세액 {Ui.Won(total - sup)}  /  합계 {Ui.Won(total)}"; }
        else { double tax = RequestReader.Num(txtTax.Text); lblSupply.Text = $"공급가 {Ui.Won(total - tax)} (= 합계 − 세액)  /  세액 {Ui.Won(tax)}  /  합계 {Ui.Won(total)}"; }
    }

    void OpenSuppliers()
    {
        int keep = SelectedSupplierId();
        using var f = new SupplierForm(db);
        f.ShowDialog(this);
        LoadSuppliers(keep);
        if (f.Chosen > 0)
        {
            // [이 업체로 새 견적 시작]: 새 견적을 열고 그 업체를 지정
            if (dirty && MessageBox.Show("저장하지 않은 내용이 있습니다. 버리고 새 견적을 시작할까요?", "새 견적", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            NewQuote(false);
            LoadSuppliers(f.Chosen);
            var s = suppliers.FirstOrDefault(x => x.Id == f.Chosen);
            if (s != null) { ApplySupplier(s); dirty = false; UpdateState(); tabs.SelectedIndex = 0; }
        }
    }

    void ApplyTitle() { Text = $"{db.AppName} v{Program.Version}"; head?.Invalidate(); }

    // ---------------- 끌어다 놓기 ----------------
    void HookDrop(Control c)
    {
        c.AllowDrop = true;
        c.DragEnter += (_, e) => { if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
        c.DragDrop += (_, e) =>
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is string[] f && f.Length > 0)
            {
                string path = f[0];
                BeginInvoke(new Action(() => { tabs.SelectedIndex = 0; ImportFile(path); }));   // 탐색기가 멈추지 않도록 끌어놓기가 끝난 뒤 처리
            }
        };
        foreach (Control ch in c.Controls) HookDrop(ch);
    }

    void UpdateAdminLabel()
    {
        lblAdmin.Text = !db.HasAdmin() ? "관리자: 설정 안 함" : Ui.AdminLoggedIn ? "관리자: 로그인 중" : "관리자: 로그아웃";
    }

    void Logout()
    {
        if (!db.HasAdmin()) { MessageBox.Show("관리자 비밀번호를 만들어 두지 않아 로그인 상태가 없습니다.\n(필요하면 설정 > 관리자에서 만들 수 있습니다.)", "로그아웃"); return; }
        if (!Ui.AdminLoggedIn) { MessageBox.Show("관리자로 로그인한 상태가 아닙니다. (삭제 같은 관리자 작업을 할 때 비밀번호를 묻습니다.)", "로그아웃"); return; }
        Ui.AdminLoggedIn = false; UpdateAdminLabel();
        MessageBox.Show("관리자 로그아웃했습니다.", "로그아웃");
    }

    // ---------------- 화면 구성 ----------------
    void BuildUi()
    {
        var menu = new MenuStrip { BackColor = Color.White };
        var file = new ToolStripMenuItem("파일(&F)");
        file.DropDownItems.Add(new ToolStripMenuItem("새 견적", null, (_, _) => NewQuote(true)) { ShortcutKeys = Keys.Control | Keys.N });
        file.DropDownItems.Add(new ToolStripMenuItem("견적의뢰 엑셀 불러오기...", null, (_, _) => PickAndImport()) { ShortcutKeys = Keys.Control | Keys.O });
        file.DropDownItems.Add(new ToolStripMenuItem("저장", null, (_, _) => Save(false)) { ShortcutKeys = Keys.Control | Keys.S });
        file.DropDownItems.Add(new ToolStripMenuItem("견적 완료", null, (_, _) => Save(true)));
        file.DropDownItems.Add(new ToolStripMenuItem("완료 해제 (관리자)", null, (_, _) => Uncomplete()));
        file.DropDownItems.Add(new ToolStripMenuItem("엑셀로 내려받기...", null, (_, _) => ExportCurrent()));
        file.DropDownItems.Add(new ToolStripSeparator());
        file.DropDownItems.Add(new ToolStripMenuItem("관리자 로그아웃", null, (_, _) => Logout()));
        file.DropDownItems.Add(new ToolStripMenuItem("프로그램 끝내기", null, (_, _) => Close()));
        var set = new ToolStripMenuItem("설정(&S)");
        set.DropDownItems.Add(new ToolStripMenuItem("업체 · 도장 · 인상률 · 백업 · 관리자...", null, (_, _) => OpenSettings()));
        set.DropDownItems.Add(new ToolStripMenuItem("견적업체(우리에게 견적 주는 업체) 등록·현황...", null, (_, _) => OpenSuppliers()));
        var help = new ToolStripMenuItem("도움말(&H)");
        help.DropDownItems.Add(new ToolStripMenuItem("사용 방법", null, (_, _) => MessageBox.Show(
            "1) 설정에서 우리 업체와 타업체(최대 3개), 도장, 인상률을 미리 등록합니다.\n" +
            "2) [엑셀 불러오기]로 견적의뢰 엑셀을 고르거나 창에 끌어다 놓습니다.\n" +
            "3) 왼쪽 목록에서 이번 견적에 쓸 타업체를 체크하면 우리 단가 × 인상률로 업체별 단가가 채워집니다. 표에서 업체별 단가를 직접 고칠 수 있습니다.\n" +
            "4) 바코드 칸에 스캔하면 품목이 추가되고, 새 바코드는 이름·단가를 입력해 수집됩니다.\n" +
            "5) [저장] → [견적 완료] 후 [엑셀 내려받기]: 저장 폴더\\년도\\견적의뢰업체명\\ 에 우리견적/타업체 엑셀이 만들어집니다.\n" +
            "6) 관리자 비밀번호는 선택 사항입니다. 만들어 두면 완료된 견적 삭제에 비밀번호가 필요합니다.", "사용 방법")));
        menu.Items.AddRange(new ToolStripItem[] { file, set, help });

        // 화면 왼쪽 아래: 로그아웃 / 프로그램 끝내기, 오른쪽: 안내 글
        var bottom = new Panel { Dock = DockStyle.Bottom, Height = 48, BackColor = Ui.BlueSoft };
        var bl = new FlowLayoutPanel { Dock = DockStyle.Left, AutoSize = true, WrapContents = false, Padding = new Padding(8, 6, 0, 0) };
        var bOut = Ui.Btn("로그아웃", (_, _) => Logout()); var bQuit = Ui.Btn("프로그램 끝내기", (_, _) => Close(), 2);
        bl.Controls.Add(bOut); bl.Controls.Add(bQuit); bl.Controls.Add(lblAdmin);
        bottom.Controls.Add(lblStatus); bottom.Controls.Add(bl);
        var status = bottom;

        head = new Panel { Dock = DockStyle.Top, Height = 54 };
        head.Paint += (_, e) =>
        {
            using var br = new System.Drawing.Drawing2D.LinearGradientBrush(head.ClientRectangle, Ui.BlueDark, Color.FromArgb(59, 139, 255), 0f);
            e.Graphics.FillRectangle(br, head.ClientRectangle);
            int x = 16;
            try { if (Icon != null) { using var ic = new Icon(Icon, 34, 34); e.Graphics.DrawIcon(ic, new Rectangle(x, 10, 34, 34)); x += 44; } } catch { }
            using var f1 = new Font("Malgun Gothic", 14f, FontStyle.Bold);
            using var f2 = new Font("Malgun Gothic", 9f);
            TextRenderer.DrawText(e.Graphics, db.AppName, f1, new Point(x, 5), Color.White);
            var sz = TextRenderer.MeasureText(db.AppName, f1);
            TextRenderer.DrawText(e.Graphics, "v" + Program.Version, f2, new Point(x + sz.Width - 4, 14), Color.FromArgb(255, 236, 160));
            TextRenderer.DrawText(e.Graphics, "견적의뢰 엑셀 → 우리 견적 + 타업체 견적 만들기", f2, new Point(x + 1, 31), Color.FromArgb(214, 228, 255));
        };

        // ----- 견적 작성 탭 -----
        var tpEdit = new TabPage("견적 작성");
        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 180, Padding = new Padding(8, 6, 8, 0), BackColor = Ui.BlueSoft, WrapContents = true };
        Label L(string t) => new Label { Text = t, AutoSize = true, ForeColor = Ui.BlueDark, Margin = new Padding(6, 8, 2, 0) };
        txtRecipient.Margin = new Padding(3, 5, 3, 0); txtRecipientTel.Margin = new Padding(3, 5, 3, 0); dtp.Margin = new Padding(3, 5, 3, 0); txtScan.Margin = new Padding(3, 5, 3, 0); cbView.Margin = new Padding(3, 5, 3, 0);
        txtRecipient.TextChanged += (_, _) => MarkDirty();
        txtRecipientTel.TextChanged += (_, _) => MarkDirty();
        dtp.ValueChanged += (_, _) => MarkDirty();
        top.Controls.Add(L("수신처(견적의뢰 업체)")); top.Controls.Add(txtRecipient); top.Controls.Add(txtRecipientTel); top.Controls.Add(L("견적일")); top.Controls.Add(dtp);
        top.Controls.Add(lblState); lblState.Margin = new Padding(14, 8, 0, 0);
        top.SetFlowBreak(lblState, true);
        cbSup.Margin = new Padding(3, 5, 3, 0); cbTitle.Margin = new Padding(3, 5, 3, 0);
        cbTitle.Items.AddRange(new object[] { "견적서", "거래명세서", "발주서", "납품서", "청구서", "주문서" });
        cbTitle.TextChanged += (_, _) => MarkDirty();
        cbSup.SelectedIndexChanged += (_, _) =>
        {
            if (suppress || cbSup.SelectedIndex <= 0 || cbSup.SelectedIndex > suppliers.Count) return;
            ApplySupplier(suppliers[cbSup.SelectedIndex - 1]);
        };
        top.Controls.Add(L("견적업체 선택")); top.Controls.Add(cbSup);
        top.Controls.Add(L("문서 제목(우리 견적서만)")); top.Controls.Add(cbTitle);
        top.Controls.Add(new Label { Text = "견적업체를 고르면 수신처가 그 업체로 채워집니다. 고르지 않으면 직접 입력/파일명 기준.", AutoSize = true, ForeColor = Color.Gray, Margin = new Padding(10, 9, 0, 0) });
        top.SetFlowBreak(top.Controls[top.Controls.Count - 1], true);
        top.Controls.Add(Ui.Btn("견적의뢰 엑셀 올리기", (_, _) => PickAndImport(), 1));
        top.Controls.Add(Ui.Btn("새 견적", (_, _) => NewQuote(true)));
        top.Controls.Add(Ui.Btn("저장", (_, _) => Save(false)));
        top.Controls.Add(Ui.Btn("견적 완료", (_, _) => Save(true)));
        top.Controls.Add(Ui.Btn("미리보기", (_, _) => PreviewCurrent()));
        var bExp = Ui.Btn("엑셀 내려받기", (_, _) => ExportCurrent(), 1);
        top.Controls.Add(bExp); top.SetFlowBreak(bExp, true);
        top.Controls.Add(Ui.Btn("+ 행 추가", (_, _) => AddRow()));
        top.Controls.Add(Ui.Btn("선택 행 삭제", (_, _) => DeleteRows(), 2));
        top.Controls.Add(Ui.Btn("견적업체 현황", (_, _) => OpenSuppliers()));
        top.Controls.Add(L("보기")); top.Controls.Add(cbView);
        top.Controls.Add(L("바코드")); top.Controls.Add(txtScan);
        cbView.SelectedIndexChanged += (_, _) => { if (!viewing) ApplyView(); };
        txtScan.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; var c = txtScan.Text.Trim(); txtScan.Clear(); if (c != "") ScanBarcode(c); } };

        var left = new Panel { Dock = DockStyle.Left, Width = 250, Padding = new Padding(8), BackColor = Color.White };
        var lblL = new Label { Dock = DockStyle.Top, Height = 40, Text = "이번 견적에 쓸 타업체\n(체크하면 단가 열이 생깁니다)", ForeColor = Ui.BlueDark };
        var btnEdit = Ui.Btn("업체 정보 수정 (이 견적만)", (_, _) => EditCompanyMenu()); btnEdit.Dock = DockStyle.Bottom; btnEdit.AutoSize = false; btnEdit.Height = 36;
        var btnRe = Ui.Btn("인상률 다시 적용 (직접 고친 값도 초기화)", (_, _) => ReapplyMarkup()); btnRe.Dock = DockStyle.Bottom; btnRe.AutoSize = false; btnRe.Height = 50;
        var btnFresh = Ui.Btn("설정의 최신 업체 정보로 갱신", (_, _) => RefreshSnapshots()); btnFresh.Dock = DockStyle.Bottom; btnFresh.AutoSize = false; btnFresh.Height = 36;
        left.Controls.Add(clbOthers); left.Controls.Add(lblL); left.Controls.Add(btnFresh); left.Controls.Add(btnEdit); left.Controls.Add(btnRe);
        clbOthers.ItemCheck += (_, _) => { if (!suppress) BeginInvoke(new Action(ApplyOtherSelection)); };

        Ui.StyleGrid(grid);
        grid.Dock = DockStyle.Fill; grid.AllowUserToAddRows = true; grid.AllowUserToDeleteRows = true; grid.SelectionMode = DataGridViewSelectionMode.CellSelect; grid.MultiSelect = true;
        grid.RowHeadersSelectionModeHack();
        grid.CellEndEdit += Grid_CellEndEdit;
        grid.DataError += (_, e) => { e.ThrowException = false; e.Cancel = false; };
        grid.UserDeletedRow += (_, _) => { MarkDirty(); Renumber(); UpdateTotals(); };
        grid.RowsAdded += (_, _) => { if (!suppress) Renumber(); };
        // 우리 견적서 하단 금액 표시 설정 (우리 견적서에만 적용)
        var taxPanel = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(8, 5, 8, 0), BackColor = Color.White, WrapContents = false };
        cbTax.Items.AddRange(new object[] { "공급가·세액·합계 표시 (세액 자동 계산)", "세액 직접 수정 (공급가 자동 맞춤)", "합계금액만 표시 (공급가·세액 숨김)" });
        cbTax.Margin = new Padding(3, 3, 3, 0); txtTax.Margin = new Padding(3, 3, 3, 0);
        taxPanel.Controls.Add(new Label { Text = "우리 견적서 하단 금액", AutoSize = true, ForeColor = Ui.BlueDark, Margin = new Padding(3, 9, 3, 0) });
        taxPanel.Controls.Add(cbTax);
        taxPanel.Controls.Add(new Label { Text = "세액", AutoSize = true, ForeColor = Ui.BlueDark, Margin = new Padding(8, 9, 2, 0) });
        taxPanel.Controls.Add(txtTax); taxPanel.Controls.Add(lblSupply);
        cbTax.SelectedIndexChanged += (_, _) =>
        {
            if (suppress || cbTax.SelectedIndex < 0) return;
            txtTax.Enabled = cbTax.SelectedIndex == 1;
            if (cbTax.SelectedIndex == 1 && RequestReader.Num(txtTax.Text) <= 0) { suppress = true; txtTax.Text = Ui.Won(Math.Round(ourTotal) - Math.Round(Math.Round(ourTotal) / 1.1)); suppress = false; }
            MarkDirty(); UpdateTaxLabel();
        };
        txtTax.TextChanged += (_, _) => { if (!suppress) MarkDirty(); UpdateTaxLabel(); };
        tpEdit.Controls.Add(grid); tpEdit.Controls.Add(left); tpEdit.Controls.Add(taxPanel); tpEdit.Controls.Add(lblTotals); tpEdit.Controls.Add(top);

        // ----- 견적 목록 탭 -----
        var tpList = new TabPage("견적 목록 (조회·수정·삭제)");
        var ltop = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(8, 6, 8, 0), BackColor = Ui.BlueSoft };
        txtListFind.Margin = new Padding(3, 5, 8, 0);
        txtListFind.TextChanged += (_, _) => FillList();
        ltop.Controls.Add(txtListFind);
        ltop.Controls.Add(Ui.Btn("열기 / 수정", (_, _) => OpenSelected(), 1));
        ltop.Controls.Add(Ui.Btn("엑셀 내려받기", (_, _) => ExportSelected()));
        ltop.Controls.Add(Ui.Btn("삭제", (_, _) => DeleteSelected(), 2));
        ltop.Controls.Add(Ui.Btn("새로고침", (_, _) => FillList()));
        Ui.StyleGrid(listGrid);
        listGrid.Dock = DockStyle.Fill; listGrid.ReadOnly = true; listGrid.AllowUserToAddRows = false; listGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; listGrid.MultiSelect = false; listGrid.RowHeadersVisible = false;
        listGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        listGrid.Columns.Add("id", "번호"); listGrid.Columns.Add("rc", "수신처"); listGrid.Columns.Add("dt", "견적일"); listGrid.Columns.Add("st", "상태"); listGrid.Columns.Add("it", "품목 수"); listGrid.Columns.Add("tt", "우리 합계"); listGrid.Columns.Add("up", "수정일");
        listGrid.Columns["id"].FillWeight = 40; listGrid.Columns["rc"].FillWeight = 200; listGrid.Columns["it"].FillWeight = 50; listGrid.Columns["tt"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        listGrid.CellDoubleClick += (_, e) => { if (e.RowIndex >= 0) OpenSelected(); };
        tpList.Controls.Add(listGrid); tpList.Controls.Add(ltop);

        tabs.TabPages.Add(tpEdit); tabs.TabPages.Add(tpList);
        tabs.SelectedIndexChanged += (_, _) => { if (tabs.SelectedIndex == 1) FillList(); };

        Controls.Add(tabs); Controls.Add(head); Controls.Add(menu); Controls.Add(status);
        MainMenuStrip = menu;
    }

    // ---------------- 상태 ----------------
    void MarkDirty() { if (suppress) return; dirty = true; UpdateState(); }

    void UpdateState()
    {
        if (cur == null) return;
        lblState.Text = (cur.Id == 0 ? "새 견적" : $"견적 #{cur.Id}") + " · " + cur.Status + (dirty ? " · ● 저장 안 됨" : "");
        lblState.ForeColor = cur.Completed ? Color.FromArgb(30, 130, 60) : Ui.BlueDark;
    }

    void CheckOur()
    {
        if (string.IsNullOrWhiteSpace(db.GetOur().Name) &&
            MessageBox.Show("우리 업체 정보(상호·사업자번호 등)가 아직 없습니다.\n지금 설정 화면을 열까요?", "처음 설정", MessageBoxButtons.YesNo) == DialogResult.Yes) OpenSettings();
    }

    void OpenSettings()
    {
        using var f = new SettingsForm(db);
        f.ShowDialog(this);
        ApplyTitle();
        if (f.Restored) { dirty = false; NewQuote(false); return; }
        LoadRegistered();
    }

    void LoadRegistered()
    {
        regOthers = db.GetOthers();
        suppress = true;
        clbOthers.Items.Clear();
        foreach (var c in regOthers) clbOthers.Items.Add($"{c.Name}  (+{db.MarkupOf(c):0.##}%)", cur != null && cur.Others.Any(o => o.Id == c.Id));
        suppress = false;
    }

    // ---------------- 새 견적 / 불러오기 ----------------
    void NewQuote(bool ask)
    {
        if (ask && dirty && MessageBox.Show("저장하지 않은 내용이 있습니다. 버리고 새 견적을 시작할까요?", "새 견적", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
        cur = new Quote { Our = db.GetOur() };
        askedCompletedEdit = false;
        suppress = true; txtRecipient.Text = ""; txtRecipientTel.Text = ""; dtp.Value = DateTime.Today; suppress = false;
        LoadRegistered();
        LoadTaxUi(db.DefaultTaxMode, 0);
        LoadSuppliers(0); cbTitle.Text = db.DefaultTitle;
        BuildGrid(new List<QItem>());
        dirty = false; UpdateState(); UpdateTotals();
        lblStatus.Text = "엑셀 파일을 끌어다 놓거나 [엑셀 불러오기]를 누르세요.";
    }

    void LoadQuoteToUi(Quote q)
    {
        cur = q; askedCompletedEdit = false;
        suppress = true; txtRecipient.Text = q.Recipient; txtRecipientTel.Text = q.RecipientTel ?? ""; dtp.Value = q.Date < dtp.MinDate ? DateTime.Today : q.Date; suppress = false;
        LoadRegistered();
        LoadTaxUi(q.TaxMode, q.TaxVal);
        LoadSuppliers(q.SupplierId); suppress = true; cbTitle.Text = string.IsNullOrWhiteSpace(q.Title) ? db.DefaultTitle : q.Title; suppress = false;
        BuildGrid(q.Items);
        dirty = false; UpdateState(); UpdateTotals();
        lblStatus.Text = $"견적 #{q.Id} 을(를) 열었습니다.";
    }

    // ---------------- 표 ----------------
    void BuildGrid(List<QItem> items)
    {
        suppress = true;
        grid.Columns.Clear();
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "bc", HeaderText = "바코드", Width = 120 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "nm", HeaderText = "품목", Width = 330 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "un", HeaderText = "단위", Width = 60 });
        grid.Columns.Add(NumCol("qty", "수량", "#,##0.##", 70));
        grid.Columns.Add(NumCol("our", "우리 단가", "#,##0", 100));
        foreach (var o in cur.Others)
        {
            var col = NumCol("p" + o.Id, $"{o.Name} (+{db.MarkupOf(o):0.##}%)", "#,##0", 120);
            col.Tag = o;
            col.HeaderCell.Style.BackColor = Color.FromArgb(70, 120, 200);
            grid.Columns.Add(col);
        }
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "nt", HeaderText = "비고", Width = 140 });
        foreach (var it in items)
        {
            var vals = new List<object> { it.Barcode, it.Name, it.Unit, it.Qty, it.OurPrice };
            foreach (var o in cur.Others) vals.Add(it.Prices.TryGetValue(o.Id, out var p) ? p : 0d);
            vals.Add(it.Note);
            int r = grid.Rows.Add(vals.ToArray());
            foreach (var o in cur.Others) if (it.Manual.Contains(o.Id)) grid.Rows[r].Cells["p" + o.Id].Tag = true;
        }
        suppress = false;
        Renumber(); UpdateTotals(); RefreshView();
    }

    // 보기: 전체 비교 / 우리 견적만 / 업체별 견적만 (보기만 바뀌고 자료는 그대로)
    void RefreshView()
    {
        int sel = cbView.SelectedIndex;
        viewing = true;
        cbView.Items.Clear();
        cbView.Items.Add("전체 비교 (우리 + 선택한 타업체)");
        cbView.Items.Add("우리 견적만");
        foreach (var o in cur.Others) cbView.Items.Add(o.Name + " 견적만");
        cbView.SelectedIndex = sel >= 0 && sel < cbView.Items.Count ? sel : 0;
        viewing = false;
        ApplyView();
    }

    void ApplyView()
    {
        if (cur == null || !grid.Columns.Contains("our")) return;
        int s = Math.Max(0, cbView.SelectedIndex);
        grid.Columns["our"].Visible = s <= 1;
        for (int k = 0; k < cur.Others.Count; k++)
            if (grid.Columns.Contains("p" + cur.Others[k].Id)) grid.Columns["p" + cur.Others[k].Id].Visible = s == 0 || s == k + 2;
    }

    // 행 추가 / 삭제
    void AddRow()
    {
        var vals = new List<object> { "", "", "", 1d, 0d };
        for (int k = 0; k < cur.Others.Count; k++) vals.Add(0d);
        vals.Add("");
        suppress = true; int idx = grid.Rows.Add(vals.ToArray()); suppress = false;
        Renumber(); MarkDirty(); UpdateTotals();
        grid.CurrentCell = grid.Rows[idx].Cells["nm"]; grid.BeginEdit(true);
    }

    void DeleteRows()
    {
        var rows = grid.SelectedCells.Cast<DataGridViewCell>().Select(c => c.RowIndex).Distinct().Where(i => i >= 0 && i < grid.Rows.Count && !grid.Rows[i].IsNewRow).OrderByDescending(i => i).ToList();
        if (rows.Count == 0) { MessageBox.Show("지울 행을 고르세요. (맨 왼쪽 번호 칸을 누르면 줄 전체가 선택됩니다)"); return; }
        if (MessageBox.Show($"선택한 {rows.Count}개 행을 삭제할까요?", "행 삭제", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
        suppress = true; foreach (int i in rows) grid.Rows.RemoveAt(i); suppress = false;
        MarkDirty(); Renumber(); UpdateTotals();
    }

    static DataGridViewTextBoxColumn NumCol(string name, string header, string fmt, int w) =>
        new DataGridViewTextBoxColumn { Name = name, HeaderText = header, Width = w, ValueType = typeof(double), DefaultCellStyle = new DataGridViewCellStyle { Format = fmt, Alignment = DataGridViewContentAlignment.MiddleRight } };

    void Renumber()
    {
        for (int i = 0; i < grid.Rows.Count; i++) grid.Rows[i].HeaderCell.Value = grid.Rows[i].IsNewRow ? "" : (i + 1).ToString();
    }

    static double D(object v) => v == null ? 0 : v is double d ? d : RequestReader.Num(Convert.ToString(v, CultureInfo.CurrentCulture));

    List<QItem> ReadGrid()
    {
        var list = new List<QItem>();
        foreach (DataGridViewRow r in grid.Rows)
        {
            if (r.IsNewRow) continue;
            string name = Convert.ToString(r.Cells["nm"].Value)?.Trim() ?? "";
            if (name == "") continue;
            var it = new QItem
            {
                Barcode = Convert.ToString(r.Cells["bc"].Value)?.Trim() ?? "", Name = name, Unit = Convert.ToString(r.Cells["un"].Value)?.Trim() ?? "",
                Qty = D(r.Cells["qty"].Value), OurPrice = D(r.Cells["our"].Value), Note = Convert.ToString(r.Cells["nt"].Value)?.Trim() ?? ""
            };
            foreach (var o in cur.Others)
            {
                var c = r.Cells["p" + o.Id];
                it.Prices[o.Id] = D(c.Value);
                if (c.Tag is bool b && b) it.Manual.Add(o.Id);
            }
            list.Add(it);
        }
        return list;
    }

    void Grid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || suppress) return;
        var row = grid.Rows[e.RowIndex];
        string col = grid.Columns[e.ColumnIndex].Name;
        if (col == "our")
        {
            double our = D(row.Cells["our"].Value);
            foreach (var o in cur.Others)
            {
                var c = row.Cells["p" + o.Id];
                if (!(c.Tag is bool b && b)) c.Value = db.PriceFor(o, our);
            }
        }
        else if (col.StartsWith("p")) row.Cells[col].Tag = true;
        MarkDirty(); Renumber(); UpdateTotals();
    }

    void UpdateTotals()
    {
        double our = 0; var sums = cur?.Others.ToDictionary(o => o.Id, _ => 0d) ?? new Dictionary<int, double>();
        int n = 0, zero = 0;
        foreach (DataGridViewRow r in grid.Rows)
        {
            if (r.IsNewRow || string.IsNullOrWhiteSpace(Convert.ToString(r.Cells["nm"].Value))) continue;
            n++;
            double q = D(r.Cells["qty"].Value), p = D(r.Cells["our"].Value);
            our += q * p; if (p <= 0) zero++;
            foreach (var id in sums.Keys.ToList()) sums[id] += q * D(r.Cells["p" + id].Value);
        }
        var parts = new List<string> { $"{(string.IsNullOrWhiteSpace(cur?.Our?.Name) ? "우리" : cur.Our.Name)}  {Ui.Won(our)}원" };
        if (cur != null) foreach (var o in cur.Others) parts.Add($"{o.Name}  {Ui.Won(sums[o.Id])}원");
        lblTotals.Text = $"품목 {n}개" + (zero > 0 ? $"  (우리 단가가 비어 있는 품목 {zero}개)" : "") + "   ·   합계(VAT 포함)\n" + string.Join("   |   ", parts);
        ourTotal = our; UpdateTaxLabel();
    }

    // ---------------- 타업체 선택 / 인상률 ----------------
    void ApplyOtherSelection()
    {
        if (cur == null) return;
        var items = ReadGrid();
        var chosen = new List<Company>();
        for (int i = 0; i < regOthers.Count; i++)
        {
            if (!clbOthers.GetItemChecked(i)) continue;
            var old = cur.Others.FirstOrDefault(o => o.Id == regOthers[i].Id);
            chosen.Add(old ?? regOthers[i].Clone());
        }
        var added = chosen.Where(c => !cur.Others.Any(o => o.Id == c.Id)).ToList();
        cur.Others = chosen;
        var ids = new HashSet<int>(chosen.Select(c => c.Id));
        foreach (var it in items)
        {
            foreach (var k in it.Prices.Keys.Where(k => !ids.Contains(k)).ToList()) { it.Prices.Remove(k); it.Manual.Remove(k); }
            foreach (var c in added) it.Prices[c.Id] = db.PriceFor(c, it.OurPrice);
        }
        BuildGrid(items);
        MarkDirty();
    }

    void ReapplyMarkup()
    {
        if (cur.Others.Count == 0) { MessageBox.Show("먼저 왼쪽에서 타업체를 체크하세요."); return; }
        if (MessageBox.Show("모든 타업체 단가를 '우리 단가 × 인상률'로 다시 계산합니다.\n직접 고친 단가도 초기화됩니다. 계속할까요?", "인상률 다시 적용", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
        var items = ReadGrid();
        foreach (var it in items) { it.Manual.Clear(); foreach (var o in cur.Others) it.Prices[o.Id] = db.PriceFor(o, it.OurPrice); }
        BuildGrid(items); MarkDirty();
    }

    void RefreshSnapshots()
    {
        if (MessageBox.Show("이 견적의 우리·타업체 정보(상호, 사업자번호, 도장, 인상률 등)를 설정의 최신 정보로 바꿉니다.\n이 견적에서 따로 고친 업체 정보는 사라집니다. 계속할까요?", "업체 정보 갱신", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
        var items = ReadGrid();
        cur.Our = db.GetOur();
        cur.Others = cur.Others.Select(o => db.GetCompany(o.Id) ?? o).ToList();
        BuildGrid(items); MarkDirty();
    }

    /// 이 견적에 들어가는 업체 정보만 수정 (설정에는 영향 없음)
    void EditCompanyMenu()
    {
        var m = new ContextMenuStrip();
        m.Items.Add($"우리: {cur.Our.Name}", null, (_, _) => EditCompany(cur.Our, false));
        foreach (var o in cur.Others) { var oo = o; m.Items.Add($"타업체: {o.Name}", null, (_, _) => EditCompany(oo, true)); }
        m.Show(Cursor.Position);
    }

    void EditCompany(Company c, bool other)
    {
        using var f = new Form { Text = "업체 정보 수정 - 이 견적에만 적용", Width = 700, Height = other ? 560 : 490, StartPosition = FormStartPosition.CenterParent, Font = Font, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false };
        var ed = new CompanyEditor(other) { Left = 10, Top = 10 }; ed.LoadFrom(c);
        int by = other ? 430 : 360;
        var ok = Ui.Btn("적용", null, 1); ok.Left = 10; ok.Top = by;
        var no = Ui.Btn("취소", null); no.Left = 100; no.Top = by; no.DialogResult = DialogResult.Cancel;
        ok.Click += (_, _) => { var msg = ed.Apply(c); if (msg != null) { MessageBox.Show(msg); return; } f.DialogResult = DialogResult.OK; };
        f.Controls.AddRange(new Control[] { ed, ok, no }); f.CancelButton = no;
        if (f.ShowDialog(this) != DialogResult.OK) return;
        var items = ReadGrid();
        if (other)   // 인상률이 바뀌었을 수 있으므로 직접 고치지 않은 단가만 다시 계산
            foreach (var it in items) if (!it.Manual.Contains(c.Id)) it.Prices[c.Id] = db.PriceFor(c, it.OurPrice);
        BuildGrid(items); MarkDirty();
    }

    // ---------------- 엑셀 불러오기 ----------------
    void PickAndImport()
    {
        using var d = new OpenFileDialog { Filter = "견적의뢰 엑셀|*.xls;*.xlsx;*.csv|모든 파일|*.*", Title = "견적의뢰 엑셀 선택" };
        if (d.ShowDialog() == DialogResult.OK) ImportFile(d.FileName);
    }

    void ImportFile(string path)
    {
        try
        {
            Cursor = Cursors.WaitCursor;
            var rows = RequestReader.ReadRows(path);
            var m = RequestReader.Detect(rows);
            if (!m.Ok)
            {
                Cursor = Cursors.Default;
                using var dlg = new MapDialog(rows, m);
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                m = dlg.Result();
                if (m.Name < 0) { MessageBox.Show("품목 열을 지정해 주세요."); return; }
                Cursor = Cursors.WaitCursor;
            }
            var items = RequestReader.Parse(rows, m);
            if (items.Count == 0) { MessageBox.Show("읽을 수 있는 품목이 없습니다. 머리글(품목/단위/수량) 위치를 확인하세요."); return; }
            int filled = 0;
            foreach (var it in items)   // 바코드로 우리 상품 목록에서 이름·단위·단가 보충
            {
                if (it.Barcode == "") continue;
                var p = db.GetProduct(it.Barcode); if (p == null) continue;
                if (it.OurPrice <= 0 && p.Price > 0) { it.OurPrice = p.Price; filled++; }
                if (it.Unit == "" && p.Unit != "") it.Unit = p.Unit;
            }
            var existing = ReadGrid();
            if (existing.Count > 0)
            {
                var a = MessageBox.Show($"이미 품목 {existing.Count}개가 있습니다.\n[예] 뒤에 이어 붙이기   [아니오] 모두 바꾸기", "엑셀 불러오기", MessageBoxButtons.YesNoCancel);
                if (a == DialogResult.Cancel) return;
                if (a == DialogResult.Yes) { existing.AddRange(items); items = existing; }
            }
            foreach (var it in items) foreach (var o in cur.Others) if (!it.Prices.ContainsKey(o.Id)) it.Prices[o.Id] = db.PriceFor(o, it.OurPrice);
            if (string.IsNullOrWhiteSpace(txtRecipient.Text)) { suppress = true; txtRecipient.Text = Path.GetFileNameWithoutExtension(path); suppress = false; }
            cur.SourceFile = Path.GetFileName(path);
            BuildGrid(items); MarkDirty();
            lblStatus.Text = $"{Path.GetFileName(path)} 에서 품목 {items.Count}개를 읽었습니다." + (filled > 0 ? $" (바코드로 단가 {filled}개 보충)" : "");
            if (items.Any(i => i.OurPrice <= 0)) lblStatus.Text += " · 단가가 비어 있는 품목은 '우리 단가'를 입력하세요.";
        }
        catch (Exception e) { MessageBox.Show("엑셀을 읽지 못했습니다.\n" + e.Message + "\n\n그래도 안 읽히면 엑셀에서 열어 .xlsx 로 다시 저장해 보세요.", "오류"); }
        finally { Cursor = Cursors.Default; }
    }

    // ---------------- 바코드 수집 ----------------
    void ScanBarcode(string code)
    {
        foreach (DataGridViewRow r in grid.Rows)
        {
            if (r.IsNewRow) continue;
            if (Convert.ToString(r.Cells["bc"].Value)?.Trim() == code)
            {
                r.Cells["qty"].Value = D(r.Cells["qty"].Value) + 1;
                grid.CurrentCell = r.Cells["qty"]; MarkDirty(); UpdateTotals();
                lblStatus.Text = $"{code}  수량 +1 → {D(r.Cells["qty"].Value):0.##}";
                txtScan.Focus(); return;
            }
        }
        var p = db.GetProduct(code);
        if (p == null || string.IsNullOrWhiteSpace(p.Name))
        {
            var np = AskNewProduct(code, p);
            if (np == null) { txtScan.Focus(); return; }
            db.SaveProduct(np); p = np;
        }
        var it = new QItem { Barcode = code, Name = p.Name, Unit = p.Unit, Qty = 1, OurPrice = p.Price, Note = p.Note };
        var vals = new List<object> { it.Barcode, it.Name, it.Unit, it.Qty, it.OurPrice };
        foreach (var o in cur.Others) vals.Add(db.PriceFor(o, it.OurPrice));
        vals.Add(it.Note);
        suppress = true; int idx = grid.Rows.Add(vals.ToArray()); suppress = false;
        grid.CurrentCell = grid.Rows[idx].Cells["qty"];
        Renumber(); MarkDirty(); UpdateTotals();
        lblStatus.Text = $"{code}  {p.Name} 추가";
        txtScan.Focus();
    }

    Product AskNewProduct(string code, Product old)
    {
        using var f = new Form { Text = "새 바코드 수집", Width = 440, Height = 280, StartPosition = FormStartPosition.CenterParent, Font = Font, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false };
        f.Controls.Add(new Label { Left = 14, Top = 12, Width = 400, Text = "처음 보는 바코드입니다: " + code, ForeColor = Ui.BlueDark });
        TextBox T(string label, int y, string v) { f.Controls.Add(new Label { Left = 14, Top = y + 4, Width = 70, Text = label }); var t = new TextBox { Left = 90, Top = y, Width = 320, Text = v }; f.Controls.Add(t); return t; }
        var nm = T("품목", 44, old?.Name ?? ""); var un = T("단위", 80, old?.Unit ?? ""); var pr = T("우리 단가", 116, old != null && old.Price > 0 ? old.Price.ToString("0.##") : "");
        var ok = Ui.Btn("수집", null, 1); ok.Left = 224; ok.Top = 170; ok.DialogResult = DialogResult.OK;
        var no = Ui.Btn("건너뛰기", null); no.Left = 314; no.Top = 170; no.DialogResult = DialogResult.Cancel;
        f.Controls.Add(ok); f.Controls.Add(no); f.AcceptButton = ok; f.CancelButton = no;
        if (f.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(nm.Text)) return null;
        return new Product { Barcode = code, Name = nm.Text.Trim(), Unit = un.Text.Trim(), Price = RequestReader.Num(pr.Text) };
    }

    // ---------------- 저장 / 완료 ----------------
    bool Save(bool complete)
    {
        var items = ReadGrid();
        if (string.IsNullOrWhiteSpace(txtRecipient.Text)) { MessageBox.Show("수신처(견적의뢰 업체명)를 입력하세요."); txtRecipient.Focus(); return false; }
        if (items.Count == 0) { MessageBox.Show("품목이 없습니다."); return false; }
        if (string.IsNullOrWhiteSpace(cur.Our.Name)) { MessageBox.Show("우리 업체 정보가 비어 있습니다. 설정 > 우리 업체에서 먼저 등록하세요."); return false; }
        if (complete)
        {
            int zero = items.Count(i => i.OurPrice <= 0);
            if (zero > 0 && MessageBox.Show($"우리 단가가 0인 품목이 {zero}개 있습니다. 그래도 완료할까요?", "견적 완료", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return false;
            if (MessageBox.Show("견적을 완료 처리합니다.\n완료된 견적을 삭제하려면 관리자 확인이 필요합니다. (수정은 계속 가능)", "견적 완료", MessageBoxButtons.YesNo) != DialogResult.Yes) return false;
        }
        else if (cur.Completed && !askedCompletedEdit)
        {
            if (MessageBox.Show("완료된 견적을 수정해서 저장합니다. 계속할까요?", "완료 견적 수정", MessageBoxButtons.YesNo) != DialogResult.Yes) return false;
            askedCompletedEdit = true;
        }
        try
        {
            cur.Recipient = txtRecipient.Text.Trim(); cur.RecipientTel = txtRecipientTel.Text.Trim(); cur.Date = dtp.Value.Date; cur.Items = items; SyncTax();
            if (complete) cur.Status = "완료";
            bool wasCompleted = cur.Completed && cur.Id != 0 && !complete;
            db.SaveQuote(cur);
            if (wasCompleted) db.Audit("완료견적수정", "id=" + cur.Id);
            int bc = db.CollectBarcodes(items);
            dirty = false; UpdateState();
            string msg = $"저장했습니다. (견적 #{cur.Id}" + (bc > 0 ? $", 바코드 {bc}개 수집" : "") + ")";
            if (complete && db.AutoBackup)
            {
                try { var f = db.Backup(); msg += "\n자동 백업: " + f; } catch (Exception e) { msg += "\n(자동 백업 실패: " + e.Message + ")"; }
            }
            lblStatus.Text = msg.Replace("\n", " · ");
            if (complete) MessageBox.Show(msg + "\n\n[엑셀 내려받기]로 견적서를 받으세요.", "견적 완료");
            return true;
        }
        catch (Exception e) { MessageBox.Show("저장하지 못했습니다.\n" + e.Message, "오류"); return false; }
    }

    void Uncomplete()
    {
        if (cur.Id == 0 || !cur.Completed) { MessageBox.Show("완료된 견적이 아닙니다."); return; }
        if (!Ui.AskAdmin(db, this, "완료된 견적을 '작성중'으로 되돌립니다.")) return;
        cur.Status = "작성중"; dirty = true; db.Audit("완료해제", "id=" + cur.Id); UpdateState();
        MessageBox.Show("작성중으로 바꿨습니다. [저장]을 눌러 반영하세요.");
    }

    // ---------------- 내려받기 ----------------
    void ExportCurrent()
    {
        var items = ReadGrid();
        if (items.Count == 0) { MessageBox.Show("내려받을 품목이 없습니다."); return; }
        if (string.IsNullOrWhiteSpace(txtRecipient.Text)) { MessageBox.Show("수신처(견적의뢰 업체명)를 입력하세요."); return; }
        if (dirty && MessageBox.Show("저장하지 않은 변경이 있습니다. 지금 화면 내용으로 내려받습니다.\n(먼저 저장하려면 [아니오] 후 저장하세요)", "내려받기", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
        cur.Recipient = txtRecipient.Text.Trim(); cur.RecipientTel = txtRecipientTel.Text.Trim(); cur.Date = dtp.Value.Date; cur.Items = items; SyncTax();
        ExportFlow(cur);
    }

    void PreviewCurrent()
    {
        var items = ReadGrid();
        if (items.Count == 0) { MessageBox.Show("미리볼 품목이 없습니다."); return; }
        cur.Recipient = txtRecipient.Text.Trim(); cur.RecipientTel = txtRecipientTel.Text.Trim(); cur.Date = dtp.Value.Date; cur.Items = items; SyncTax();
        using var f = new PreviewForm(cur, db.PerPage);
        f.ShowDialog(this);
    }

    void ExportFlow(Quote q)
    {
        using var f = new Form { Text = "엑셀 내려받기", Width = 560, Height = 520, StartPosition = FormStartPosition.CenterParent, Font = Font, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false };
        f.Controls.Add(new Label { Left = 14, Top = 12, Width = 520, Text = "내려받을 견적서를 고르세요" });
        var clb = new CheckedListBox { Left = 14, Top = 38, Width = 520, Height = 120, CheckOnClick = true };
        var all = new List<Company> { q.Our }; all.AddRange(q.Others);
        foreach (var c in all) clb.Items.Add((c.Kind == "our" ? "[우리 견적] " : "[타업체] ") + c.Name, true);
        var cbSep = new CheckBox { Left = 14, Top = 166, Width = 520, Height = 44, Text = "타업체를 업체별 파일로 따로 저장\n(끄면 '타업체' 파일 1개 안에 업체별 시트로 저장)" };
        f.Controls.Add(new Label { Left = 14, Top = 220, Width = 520, Text = "견적서 저장 기본 폴더" });
        var tRoot = new TextBox { Left = 14, Top = 246, Width = 410, Text = db.ExportDir };
        var bRoot = Ui.Btn("변경", (_, _) =>
        {
            using var d = new FolderBrowserDialog { SelectedPath = tRoot.Text, Description = "견적서를 저장할 기본 폴더 선택" };
            if (d.ShowDialog() == DialogResult.OK) tRoot.Text = d.SelectedPath;
        });
        bRoot.Left = 432; bRoot.Top = 243;
        var lblPath = new Label { Left = 14, Top = 282, Width = 520, Height = 90, ForeColor = Ui.BlueDark };
        void UpdPath()
        {
            string dir = Exporter.OutputDir(tRoot.Text, q), d = q.Date.ToString("yyyyMMdd"), rc = Exporter.SafeFile(q.Recipient);
            lblPath.Text = "저장 위치(없으면 자동으로 만듭니다)\n" + dir + "\n파일: " + rc + "_우리견적_" + d + ".xlsx,  " + rc + "_타업체_" + d + ".xlsx";
        }
        tRoot.TextChanged += (_, _) => UpdPath(); UpdPath();
        var ok = Ui.Btn("내려받기", null, 1); ok.Left = 330; ok.Top = 410; ok.DialogResult = DialogResult.OK;
        var no = Ui.Btn("취소", null); no.Left = 440; no.Top = 410; no.DialogResult = DialogResult.Cancel;
        f.Controls.AddRange(new Control[] { clb, cbSep, tRoot, bRoot, lblPath, ok, no }); f.AcceptButton = ok; f.CancelButton = no;
        if (f.ShowDialog(this) != DialogResult.OK) return;
        bool incOur = clb.GetItemChecked(0);
        var others = new List<Company>();
        for (int i = 1; i < all.Count; i++) if (clb.GetItemChecked(i)) others.Add(all[i]);
        if (!incOur && others.Count == 0) { MessageBox.Show("견적서를 하나 이상 고르세요."); return; }
        string root = tRoot.Text.Trim();
        if (root == "") { MessageBox.Show("저장 폴더를 지정하세요."); return; }
        try
        {
            Cursor = Cursors.WaitCursor;
            db.ExportDir = root;
            var files = Exporter.ExportAll(db, q, incOur, others, cbSep.Checked, root);
            string dir = Exporter.OutputDir(root, q);
            lblStatus.Text = $"{files.Count}개 파일을 저장했습니다: {dir}";
            Cursor = Cursors.Default;
            if (MessageBox.Show($"{files.Count}개 파일을 저장했습니다.\n\n{dir}\n\n" + string.Join("\n", files.Select(x => Path.GetFileName(x))) + "\n\n저장 폴더를 열까요?", "저장 완료", MessageBoxButtons.YesNo) == DialogResult.Yes)
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", "\"" + dir + "\"") { UseShellExecute = true });
        }
        catch (Exception e) { MessageBox.Show("엑셀을 저장하지 못했습니다. 같은 이름의 파일이 엑셀에서 열려 있지 않은지 확인하세요.\n" + e.Message, "오류"); }
        finally { Cursor = Cursors.Default; }
    }

    // ---------------- 견적 목록 ----------------
    void FillList()
    {
        listGrid.Rows.Clear();
        foreach (var r in db.ListQuotes(txtListFind.Text.Trim()))
            listGrid.Rows.Add(r.Id, r.Recipient, r.Date, r.Status, r.Items, Ui.Won(r.Total), r.Updated);
    }

    long SelectedId() => listGrid.CurrentRow == null ? 0 : Convert.ToInt64(listGrid.CurrentRow.Cells["id"].Value);

    void OpenSelected()
    {
        long id = SelectedId(); if (id == 0) return;
        if (dirty && MessageBox.Show("저장하지 않은 내용이 있습니다. 버리고 선택한 견적을 열까요?", "열기", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
        var q = db.LoadQuote(id); if (q == null) { MessageBox.Show("견적을 찾을 수 없습니다."); return; }
        LoadQuoteToUi(q); tabs.SelectedIndex = 0;
    }

    void ExportSelected()
    {
        long id = SelectedId(); if (id == 0) return;
        var q = db.LoadQuote(id); if (q != null) ExportFlow(q);
    }

    void DeleteSelected()
    {
        long id = SelectedId(); if (id == 0) return;
        var q = db.LoadQuote(id); if (q == null) return;
        if (q.Completed)
        {
            if (!Ui.AskAdmin(db, this, $"완료된 견적 #{id} ({q.Recipient}) 을(를) 삭제합니다. 되돌릴 수 없습니다.")) return;
        }
        else if (MessageBox.Show($"견적 #{id} ({q.Recipient}) 을(를) 삭제할까요? 되돌릴 수 없습니다.", "삭제", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
        db.DeleteQuote(id);
        if (cur != null && cur.Id == id) { dirty = false; NewQuote(false); }
        FillList();
    }
}

static class GridExt
{
    /// 행 번호를 눌러 줄 전체를 선택할 수 있게 한다 (셀 선택은 그대로 유지)
    public static void RowHeadersSelectionModeHack(this DataGridView g) => g.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
}
