from m365_cli.core.text import clean_body, cut_quoted, one_line, plain_to_html, truncate, with_signoff


def test_one_line_collapses_and_cuts() -> None:
    assert one_line("a\n  b\tc", 10) == "a b c"
    assert one_line("abcdefghij", 5) == "abcd…"
    assert one_line(None, 5) == ""


def test_clean_body_unwraps_safelinks_and_blank_runs() -> None:
    body = "Hi\r\n\r\n\r\n\r\nsee https://eur01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fexample.com%2Fx&data=abc"
    assert clean_body(body) == "Hi\n\nsee https://example.com/x"


def test_cut_quoted_english_and_czech_headers() -> None:
    en = "Thanks!\n\nFrom: Jan Novak\nSent: Monday\nSubject: x\nold text"
    cz = "Diky\n\nOd: Jan Novak\nOdesláno: pondělí\nPředmět: x\nstary text"
    assert cut_quoted(en) == ("Thanks!", True)
    assert cut_quoted(cz) == ("Diky", True)
    assert cut_quoted("no history") == ("no history", False)


def test_truncate() -> None:
    assert truncate("abc", 5) == ("abc", False)
    assert truncate("abcdef", 3) == ("abc", True)


def test_plain_to_html_escapes_and_keeps_lines() -> None:
    assert plain_to_html("a < b\nc") == "a &lt; b<br>c"


def test_with_signoff() -> None:
    assert with_signoff("Hi", "") == "Hi"
    assert with_signoff("Hi\n", "-- bot") == "Hi\n\n-- bot"
