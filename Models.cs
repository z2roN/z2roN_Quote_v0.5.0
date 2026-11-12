namespace ZQuote;

public class Company
{
    public int Id;                 // 원본 업체 번호 (견적 안에서는 가격 열의 키)
    public string Kind = "other";  // our | other
    public string Name = "", BizNo = "", Ceo = "", BizType = "", BizItem = "", Address = "", Tel = "", Fax = "", Memo = "";
    public int Form;               // 견적서 양식: 0 = 우리 기본 양식, 1~3 = 타업체용 변형 양식
    public byte[] Stamp;           // 도장 이미지(PNG)
    public double? Markup;         // 우리 단가 대비 인상률 %, null 이면 기본값
    public int SortNo;
    public Company Clone() => (Company)MemberwiseClone();
    public override string ToString() => Name;
}

public class QItem
{
    public string Name = "", Unit = "", Note = "", Barcode = "";
    public double Qty, OurPrice;
    public Dictionary<int, double> Prices = new();   // 타업체 번호 → 단가
    public HashSet<int> Manual = new();             // 직접 고친 타업체 단가
}

public class Quote
{
    public long Id;
    public int SupplierId;                           // 이 견적을 진행한 견적업체(0 = 지정 안 함)
    public string Title = "";                        // 우리 견적서 제목(견적서/거래명세서/발주서...), 비면 견적서
    public string TaxMode = "auto";                  // 우리 견적서 하단 금액: auto(공급가·세액 자동) | manual(세액 직접, 공급가 자동) | hide(합계금액만)
    public double TaxVal;                            // manual 일 때의 세액
    public string Recipient = "", RecipientTel = "", Status = "작성중", Note = "", SourceFile = "";
    public DateTime Date = DateTime.Today, Created = DateTime.Now, Updated = DateTime.Now;
    public Company Our;                              // 견적 시점의 우리 업체 정보(복사본)
    public List<Company> Others = new();             // 견적 시점의 타업체 정보(복사본)
    public List<QItem> Items = new();
    public bool Completed => Status == "완료";
}

/// 우리에게 견적을 주는 업체(공급처) 등록 정보
public class Supplier
{
    public int Id;
    public string Name = "", BizNo = "", Ceo = "", Contact = "", Tel = "", Fax = "", Email = "", Address = "", Items = "", Memo = "", Created = "";
}

public class Product
{
    public string Barcode = "", Name = "", Unit = "", Note = "";
    public double Price;
    public string Updated = "";
}

public static class Pricing
{
    /// 우리 단가에 인상률을 적용하고 단위(1/10/100)로 올림/반올림/버림.
    public static double Apply(double our, double pct, int unit, string mode)
    {
        if (our <= 0) return 0;
        double v = our * (1 + pct / 100.0);
        if (unit < 1) unit = 1;
        double q = v / unit;
        q = mode == "올림" ? Math.Ceiling(q - 1e-9) : mode == "버림" ? Math.Floor(q + 1e-9) : Math.Round(q, MidpointRounding.AwayFromZero);
        return q * unit;
    }
}
