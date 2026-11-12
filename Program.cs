namespace ZQuote;

static class Program
{
    public const string Version = "0.5.0";

    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.ThreadException += (_, e) => MessageBox.Show(e.Exception.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
        Db db;
        try { db = Db.OpenDefault(); }
        catch (Exception e) { MessageBox.Show("데이터 파일을 열 수 없습니다.\n" + e.Message, "오류"); return; }
        // 관리자 비밀번호는 선택 사항입니다. 설치 직후에는 묻지 않고, 설정 > 관리자 에서 만들 수 있습니다.
        Application.Run(new MainForm(db));
    }
}
