from yedaham.master.registry import Alias, Fund, Registry, normalize, tokens


def reg():
    return Registry([
        Fund("site:1", "한국민간운영권전문투자형사모투자신탁제3호", "대체투자", "운용중"),
        Fund("site:2", "한국민간운영권전문투자형사모투자신탁제4호", "대체투자", "운용중",
             [Alias("KPCF IV", "site_alias")]),
        Fund("site:3", "파인스트리트글로벌일반사모투자신탁 13호", "대체투자", "운용중",
             [Alias("파인스트리트글로벌일산사모투자신탁 13호", "site_alias")]),
        Fund("site:4", "하나증권 제14회 무보증 후순위채권", "채권", "운용중"),
    ])


def test_normalize_ignores_spacing_and_corp_marks():
    assert normalize("㈜하나 금융-지주 (주)") == normalize("하나금융지주")


def test_exact_by_alias_and_spacing():
    r = reg()
    assert r.resolve("한국민간운영권전문투자형사모투자신탁 제4호").fund.fid == "site:2"
    assert r.resolve("kpcf iv").method == "exact"
    assert r.resolve("파인스트리트글로벌일산사모투자신탁13호").fund.fid == "site:3"


def test_number_mismatch_never_matches():
    m = reg().resolve("한국민간운영권 5호")
    assert m.fund is None or m.fund.fid not in ("site:1", "site:2")


def test_suggest_requires_confirmation_then_becomes_exact():
    r = reg()
    m = r.resolve("하나증권 제14회 무보증 후순위사채", category="채권")
    assert m.method == "suggest" and m.fund.fid == "site:4"
    r.add_alias(m.fund, "하나증권 제14회 무보증 후순위사채", "manual")
    assert r.resolve("하나증권 제14회 무보증 후순위사채").method == "exact"


def test_tokens_keep_series_numbers():
    assert "10-4" in tokens("파인스트리트글로벌일반사모 10-4호")
    assert "2-3" in tokens("미래에셋PGIF일반사모특별자산투자신탁2-3호")
