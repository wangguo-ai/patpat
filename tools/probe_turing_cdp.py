"""只读探查图灵页面：连本机已登录 Chrome 的 CDP，截图 + dump 导航菜单。
绝不点击任何业务按钮、不触发任何写入操作。
"""
import sys
from playwright.sync_api import sync_playwright

CDP = "http://127.0.0.1:9222"
OUT = "C:/Users/Administrator/AppData/Local/Temp"


def dump_links(page):
    """尽量全面地收集页面里的导航/菜单链接（文本 + href）。"""
    sels = [
        "a[href]", "nav a", "aside a", ".menu a", ".sidebar a",
        "[class*=menu] a", "[class*=nav] a", "[class*=side] a",
        "[role=menuitem]", "li[class*=menu] span", ".el-menu-item",
    ]
    seen = set()
    for s in sels:
        try:
            items = page.eval_on_selector_all(
                s,
                "els=>els.map(e=>({t:(e.textContent||'').replace(/\\s+/g,' ').trim(), "
                "h:(e.getAttribute('href')||'')}))",
            )
            for it in items:
                t = it.get("t", "")
                h = it.get("h", "")
                if not t:
                    continue
                key = (t, h)
                if key in seen:
                    continue
                seen.add(key)
                print(f"  LINK: {t}  =>  {h}")
        except Exception as e:
            pass


def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP)
        print("connected, contexts:", len(browser.contexts))
        target = None
        for ctx in browser.contexts:
            for pg in ctx.pages:
                if "turing.patpat.shop" in pg.url:
                    target = pg
                    break
            if target:
                break
        if not target:
            target = browser.contexts[0].pages[0] if (
                browser.contexts and browser.contexts[0].pages) else browser.new_page()
        print("TARGET url:", target.url)
        print("TARGET title:", target.title())
        target.wait_for_timeout(2500)
        cur = f"{OUT}/_probe_current.png"
        target.screenshot(path=cur)
        print("SCREENSHOT:", cur)
        print("=== 导航/菜单链接 ===")
        dump_links(target)
        browser.close()
    print("DONE")


if __name__ == "__main__":
    main()
