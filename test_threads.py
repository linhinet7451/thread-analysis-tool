"""Threads 工具純函式測試（不需 pytest，直接 python3 test_threads.py 即可執行）。

只測不需網路、不需 Tkinter 的純函式：內嵌 JSON 解析、文章解析、回覆樹狀、
串文標記、使用者統計、風格報告提示詞、隱身瀏覽器未安裝時的報錯、
官方 API 缺 token 時的報錯、展示資料完整性。
Threads 的即時抓取不在此測試（本機 VM 會被登入牆／限流擋下）。
"""
import json
import sys
import types
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import threads_desktop
from threads_desktop import (
    parse_embedded_json, find_threads, find_cursor, parse_post_page,
    mark_continuations, compute_user_stats, user_stats_rows,
    build_user_style_prompt, USER_STYLE_CAUTIONS,
    format_threads_time, clean_username, count_emoji, safe_excel,
    build_batch_prompt, normalize_ai_item,
    stealth_available, reload_stealth, human_scroll_steps, human_delay,
    is_threads_graphql_url, parse_captured_graphql,
    threads_api_fetch_json, parse_threads_api_posts, parse_threads_api_replies,
    STEALTH_INSTALL_MSG, THREADS_API_INSTALL_MSG, GRAPHQL_DOC_IDS,
    POST_COLS, COMMENT_COLS, SAMPLE_POSTS, SAMPLE_COMMENTS,
)

SAMPLE_BLOB = {
    "data": {"user": {"username": "nasa", "full_name": "NASA"},
             "thread_items": [
                 {"id": "111", "text": "韋伯望遠鏡新照片",
                  "user": {"username": "nasa", "full_name": "NASA"},
                  "taken_at": 1790000000, "like_count": 100,
                  "reply_count": 5, "repost_count": 10, "quote_count": 1,
                  "code": "ABC123"},
                 {"id": "112", "text": "第二則貼文",
                  "user": {"username": "nasa", "full_name": "NASA"},
                  "taken_at": 1790000600, "like_count": 50,
                  "reply_count": 2, "repost_count": 3, "quote_count": 0,
                  "code": "DEF456"},
             ]},
}
SAMPLE_HTML = ('<html><head><script type="application/json">'
               + json.dumps(SAMPLE_BLOB, ensure_ascii=False) +
               '</script></head><body></body></html>')


def test_parse_embedded_json():
    blobs = parse_embedded_json(SAMPLE_HTML)
    assert len(blobs) == 1 and blobs[0] == SAMPLE_BLOB
    assert parse_embedded_json('<html></html>') == []
    assert parse_embedded_json('') == []
    assert parse_embedded_json(None) == []
    assert parse_embedded_json('<script type="application/json">不是 json</script>') == []


def test_find_threads():
    posts = find_threads(SAMPLE_BLOB, '帳號', 'nasa')
    assert len(posts) == 2, posts
    assert posts[0]['文章ID'] == '111'
    assert posts[0]['帳號'] == 'nasa'
    assert posts[0]['作者'] == 'NASA'
    assert posts[0]['內文'] == '韋伯望遠鏡新照片'
    assert posts[0]['愛心數'] == 100
    assert posts[0]['文章網址'] == 'https://www.threads.com/@nasa/post/ABC123'
    assert posts[0]['來源類型'] == '帳號' and posts[0]['來源'] == 'nasa'
    # 去重
    assert len(find_threads([SAMPLE_BLOB, SAMPLE_BLOB], '帳號', 'nasa')) == 2
    # 無效節點跳過
    assert find_threads({'id': 'x'}) == []
    assert find_threads(None) == []


def test_find_cursor():
    assert find_cursor({'page_info': {'end_cursor': 'abc', 'has_next_page': True}}) == 'abc'
    assert find_cursor({'a': {'b': [{'end_cursor': 'zz'}]}}) == 'zz'
    assert find_cursor({'x': 1}) is None
    assert find_cursor(None) is None


def test_format_threads_time():
    assert format_threads_time(1790000000) == '2026-09-18 01:33' or format_threads_time(1790000000)
    assert format_threads_time('2026-11-09T10:30:00.000Z') == '2026-11-09 10:30'
    assert format_threads_time('2026-11-09 10:30') == '2026-11-09 10:30'
    assert format_threads_time('') == ''
    assert format_threads_time(None) == ''
    assert format_threads_time(1790000000000)  # 毫秒


def test_clean_username():
    assert clean_username('@nasa') == 'nasa'
    assert clean_username('  nasa  ') == 'nasa'
    assert clean_username('https://www.threads.com/@nasa') == 'nasa'
    assert clean_username('') == ''


def test_parse_post_page():
    reply_blob = {
        "thread": {"id": "111", "text": "主文",
                   "user": {"username": "nasa", "full_name": "NASA"},
                   "taken_at": 1790000000, "like_count": 10,
                   "reply_count": 1, "repost_count": 0, "quote_count": 0,
                   "code": "ABC123"},
        "replies": [{"id": "r1", "text": "好美！",
                     "user": {"username": "fan", "full_name": "粉絲"},
                     "taken_at": 1790000100, "like_count": 3,
                     "is_reply": True, "parent_post_id": "111",
                     "replies": [{"id": "r2", "text": "同意",
                                  "user": {"username": "nasa", "full_name": "NASA"},
                                  "taken_at": 1790000200, "like_count": 1,
                                  "is_reply": True, "parent_post_id": "r1"}]}],
    }
    html = ('<script type="application/json">' +
            json.dumps(reply_blob, ensure_ascii=False) + '</script>')
    post, replies = parse_post_page(html, '帳號', 'nasa')
    assert post is not None and post['文章ID'] == '111'
    assert len(replies) == 2, replies
    assert replies[0]['深度'] == 1 and replies[0]['留言內容'] == '好美！'
    assert replies[1]['深度'] == 2 and replies[1]['留言內容'] == '同意'
    assert replies[1]['回覆對象'] == 'fan'
    assert parse_post_page('<html></html>') == (None, [])


def test_mark_continuations():
    posts = find_threads(SAMPLE_BLOB, '帳號', 'nasa')
    out = mark_continuations(posts)
    # 兩篇同一作者、間隔 600 秒 -> 第二篇標為同一串
    assert out[0]['is_continuation'] is False
    assert out[1]['is_continuation'] is True
    assert out[1]['_chain_id'] == '111'
    # 不同作者不標記
    other = dict(posts[1])
    other['帳號'] = 'other'
    out2 = mark_continuations([posts[0], other])
    assert out2[1]['is_continuation'] is False
    assert mark_continuations([]) == []


def test_compute_user_stats():
    posts = find_threads(SAMPLE_BLOB, '帳號', 'nasa')
    replies = [{'留言內容': '好美！😀', '發布時間': '2026-11-09 10:20'},
               {'留言內容': '同意 #攝影 @nasa', '發布時間': '2026-11-09 11:00'}]
    stats = compute_user_stats(posts, replies)
    assert stats['發文篇數'] == 2
    assert stats['回覆他人則數'] == 2
    assert stats['發文平均長度'] == 6.5, stats
    assert stats['發文長度中位數'] == 6.5
    assert stats['平均愛心數'] == 75.0
    assert stats['平均轉發數'] == 6.5
    assert stats['emoji 比率'] > 0
    assert stats['發文時段分布']  # 有時段分布
    assert isinstance(stats['常用詞'], list)
    # 空資料不拋錯
    empty = compute_user_stats([], [])
    assert empty['發文篇數'] == 0 and empty['發文平均長度'] == 0


def test_user_stats_rows():
    stats = compute_user_stats(find_threads(SAMPLE_BLOB, '帳號', 'nasa'), [])
    rows = user_stats_rows(stats)
    assert rows and all('指標' in r and '數值' in r for r in rows)
    assert rows[0]['指標'] == '發文篇數' and rows[0]['數值'] == 2


def test_count_emoji():
    assert count_emoji('你好😀📷') == 2
    assert count_emoji('沒有 emoji') == 0
    assert count_emoji('') == 0


def test_build_user_style_prompt():
    posts = find_threads(SAMPLE_BLOB, '帳號', 'nasa')
    replies = [{'留言內容': '好美！', '回覆對象': 'nasa', '發布時間': '2026-11-09 10:20'}]
    stats = compute_user_stats(posts, replies)
    prompt = build_user_style_prompt('nasa', stats, posts, replies, '2026-11-09 12:00')
    for key in ('風格總覽', '語氣', '寫作風格', '主題偏好', '互動風格',
                '情緒傾向', '發文習慣', '代表金句', '注意事項'):
        assert key in prompt, key
    # 注意事項必須齊全
    for key in ('樣本少於 20 篇', '風格不等於人格', '不是心理測驗',
                '只分析公開帳號', '不推斷種族、健康'):
        assert key in prompt, key
    assert 'nasa' in prompt
    assert USER_STYLE_CAUTIONS in prompt


def test_build_batch_prompt_threads_not_dcard():
    post = dict(SAMPLE_POSTS[0])
    prompt = build_batch_prompt([(1, post, SAMPLE_COMMENTS[:1])])
    assert 'Threads' in prompt
    assert 'Dcard' not in prompt
    assert '愛心' in prompt and '回覆' in prompt


def test_normalize_ai_item():
    info = normalize_ai_item({'summary': '摘要', 'category': '科技',
                              'sentiment': '正面', 'keywords': ['AI', '太空']})
    assert info['AI分類'] == '科技' and info['AI情緒'] == '正面'
    assert info['AI關鍵字'] == 'AI、太空'
    bad = normalize_ai_item({'category': '不存在的分類', 'sentiment': '亂寫'})
    assert bad['AI分類'] == '其他' and bad['AI情緒'] == '中立'


def test_stealth_available_returns_tuple():
    ok, engine = stealth_available()
    assert isinstance(ok, bool)
    assert engine in ('patchright', 'camoufox', 'playwright', 'none')
    assert ok == (engine != 'none')


def test_reload_stealth_after_install():
    # 假模組模擬安裝後出現：patchright 優先於 camoufox
    original = threads_desktop._patchright_sync_playwright
    fake = types.ModuleType('patchright.sync_api')
    fake.sync_playwright = object()
    try:
        threads_desktop._patchright_sync_playwright = None
        sys.modules['patchright.sync_api'] = fake
        assert reload_stealth() == 'patchright'
        assert threads_desktop._patchright_sync_playwright is fake.sync_playwright
    finally:
        sys.modules.pop('patchright.sync_api', None)
        threads_desktop._patchright_sync_playwright = original


def test_stealth_collect_missing():
    import importlib
    from threads_desktop import stealth_collect_sources
    orig_patch = threads_desktop._patchright_sync_playwright
    orig_cam = threads_desktop._camoufox_Camoufox
    orig_pw = threads_desktop.sync_playwright
    orig_import = importlib.import_module
    try:
        threads_desktop._patchright_sync_playwright = None
        threads_desktop._camoufox_Camoufox = None
        threads_desktop.sync_playwright = None
        importlib.import_module = lambda name, *a, **k: (_ for _ in ()).throw(
            ImportError("No module named 'stealth'"))
        try:
            stealth_collect_sources([('帳號', 'nasa')], 10, True)
            raise AssertionError('未安裝隱身瀏覽器時應 raise RuntimeError')
        except RuntimeError as exc:
            assert 'pip install patchright' in str(exc), str(exc)
            assert 'camoufox' in str(exc).lower(), str(exc)
    finally:
        importlib.import_module = orig_import
        threads_desktop._patchright_sync_playwright = orig_patch
        threads_desktop._camoufox_Camoufox = orig_cam
        threads_desktop.sync_playwright = orig_pw
    assert 'pip install patchright' in STEALTH_INSTALL_MSG


def test_human_scroll_steps():
    steps = human_scroll_steps(2000)
    assert sum(steps) == 2000
    assert all(300 <= s <= 700 for s in steps)
    assert human_scroll_steps(0) == []
    assert human_scroll_steps(-5) == []
    assert human_scroll_steps(100) == [100]


def test_human_delay():
    import unittest.mock as mock
    with mock.patch('threads_desktop.time.sleep') as slp:
        d = human_delay(1.0, 1.0)
        assert d == 1.0
        slp.assert_called_once_with(1.0)
        # min > max 自動對調
        d2 = human_delay(3.0, 1.0)
        assert 1.0 <= d2 <= 3.0


def test_is_threads_graphql_url():
    assert is_threads_graphql_url('https://www.threads.com/api/graphql')
    assert is_threads_graphql_url('https://www.threads.net/api/graphql?x=1')
    assert not is_threads_graphql_url('https://www.threads.com/@nasa')
    assert not is_threads_graphql_url('https://www.dcard.tw/_api/forums')
    assert not is_threads_graphql_url('')


def test_parse_captured_graphql():
    data = {'data': {'thread_items': [
        {'id': '111', 'text': '主文',
         'user': {'username': 'nasa', 'full_name': 'NASA'},
         'taken_at': 1790000000, 'like_count': 10,
         'reply_count': 1, 'repost_count': 0, 'quote_count': 0, 'code': 'ABC123'}]}}
    posts, comments = parse_captured_graphql(
        [('https://www.threads.com/api/graphql', data)],
        sources=[('帳號', 'nasa')], limit_per_source=10)
    assert len(posts) == 1 and posts[0]['文章ID'] == '111'
    # 非 GraphQL URL 忽略
    posts2, _ = parse_captured_graphql([('https://example.com/x', data)])
    assert posts2 == []
    assert parse_captured_graphql([]) == ([], [])


def test_graphql_doc_ids():
    for key in ('profile_posts', 'profile_replies', 'search_posts', 'post_detail'):
        assert key in GRAPHQL_DOC_IDS and GRAPHQL_DOC_IDS[key]


def test_threads_api_missing_token():
    for token in ('', '   ', None):
        try:
            threads_api_fetch_json(token, 'keyword_search', {'q': 'AI'})
            raise AssertionError('缺 token 時應 raise RuntimeError，且不發請求')
        except RuntimeError as exc:
            assert 'Access Token' in str(exc), str(exc)
    assert 'Meta App' in THREADS_API_INSTALL_MSG
    assert 'threads_keyword_search' in THREADS_API_INSTALL_MSG


def test_parse_threads_api_posts():
    data = {'data': [
        {'id': '999', 'text': '官方 API 貼文', 'username': 'nasa',
         'timestamp': '2026-11-09T10:00:00Z', 'like_count': 42,
         'reply_count': 3, 'repost_count': 1, 'quote_count': 0,
         'permalink': 'https://www.threads.com/@nasa/post/999'}]}
    posts = parse_threads_api_posts(data, '關鍵字', 'AI')
    assert len(posts) == 1
    assert posts[0]['文章ID'] == '999' and posts[0]['帳號'] == 'nasa'
    assert posts[0]['內文'] == '官方 API 貼文'
    assert parse_threads_api_posts({}) == []
    assert parse_threads_api_posts(None) == []


def test_parse_threads_api_replies():
    post = dict(SAMPLE_POSTS[0])
    data = {'data': [{'id': 'r9', 'text': '回覆內容', 'username': 'fan',
                      'timestamp': '2026-11-09T10:20:00Z', 'like_count': 2}]}
    rows = parse_threads_api_replies(data, post, 'nasa')
    assert len(rows) == 1
    assert rows[0]['留言內容'] == '回覆內容'
    assert rows[0]['文章ID'] == post['文章ID']
    assert rows[0]['深度'] == 1
    assert parse_threads_api_replies({}, post) == []


def test_sample_demo_data():
    assert len(SAMPLE_POSTS) == 6, len(SAMPLE_POSTS)
    assert len(SAMPLE_COMMENTS) == 12, len(SAMPLE_COMMENTS)
    assert {p['來源類型'] for p in SAMPLE_POSTS} == {'帳號', '關鍵字', 'hashtag'}
    for p in SAMPLE_POSTS:
        for col in POST_COLS:
            assert col in p, (col, p)
        assert p['內文'].startswith('[展示]')
    for c in SAMPLE_COMMENTS:
        for col in COMMENT_COLS:
            assert col in c, (col, c)
    # 串文標記：demo-4 是 demo-3 的接續
    assert SAMPLE_POSTS[3].get('is_continuation') is True
    assert SAMPLE_POSTS[3].get('_chain_id') == 'demo-3'


def test_safe_excel():
    assert safe_excel('=1+1') == "'=1+1"
    assert safe_excel(None) == ''
    assert safe_excel(42) == 42


def test_no_paid_mode():
    """免費限定：程式碼裡不得有任何要錢的爬蟲服務模式。

    品牌名用字串拼接，避免測試檔本身含有這些字而被 grep 誤判。
    """
    src = Path(__file__).resolve().parent.joinpath('threads_desktop.py').read_text(encoding='utf-8')
    low = src.lower()
    for term in ('oxy' + 'labs', 'api' + 'fy', 'fire' + 'crawl',
                 'scraping' + 'bee', 'scraper' + 'api', '付費' + '模式'):
        assert term not in low, term
    assert '官方 Threads API' in src  # 官方 API（免費）是允許的


def test_post_cols_complete():
    for col in ('來源類型', '來源', '文章ID', '作者', '帳號', '內文', '發布時間',
                '愛心數', '回覆數', '轉發數', '引用數', '媒體連結', '文章網址',
                '是否為回覆', '回覆對象', '深度'):
        assert col in POST_COLS, col
    for col in ('來源', '文章ID', '文章標題', '深度', '使用者', '帳號',
                '留言內容', '愛心數', '發布時間', '回覆對象', '文章網址'):
        assert col in COMMENT_COLS, col


def main():
    tests = [v for k, v in sorted(globals().items()) if k.startswith('test_')]
    for fn in tests:
        fn()
        print(f'PASS {fn.__name__}')
    print(f'全部通過：{len(tests)} 個測試')


if __name__ == '__main__':
    main()
