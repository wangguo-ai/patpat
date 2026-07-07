# -*- coding: utf-8 -*-
"""exec 联调探针 v2：用 page.content() 拿 HTML 在 Python 侧解析，避免大 evaluate 段错误。只读。"""
from playwright.sync_api import sync_playwright

PORT = 9222
HTML_OUT = r"D:\ai共享盘\MyBrain\tools\_turing_demand.html"


def main():
    with sync_playwright() as p:
        b = p.chromium.connect_over_cdp(f"http://127.0.0.1:{PORT}")
        ctx = b.contexts[0]
        page = None
        for pg in ctx.pages:
            if "ver-material-demand-list" in pg.url:
                page = pg
                break
        if page is None:
            page = ctx.pages[0]
        print("TARGET_URL:", page.url)
        html = page.content()          # 浏览器只返回字符串，避免大 evaluate
        print("HTML_LEN:", len(html))
        with open(HTML_OUT, "w", encoding="utf-8") as f:
            f.write(html)
        print("HTML_SAVED:", HTML_OUT)
        b.close()
    print("PROBE_OK")


if __name__ == "__main__":
    main()
