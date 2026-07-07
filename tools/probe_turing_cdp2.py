"""阶段2：只读探查图灵「版料需求」等目标页。
新开 tab 导航，不破坏大王当前视图。
"""
import json, sys
from playwright.sync_api import sync_playwright

CDP = "http://127.0.0.1:9222"
BASE = "https://turing.patpat.shop"
OUT = "C:/Users/Administrator/AppData/Local/Temp"


def screenshot(page, name):
    p = f"{OUT}/_probe_{name}.png"
    page.screenshot(path=p)
    print(f"SCREENSHOT: {p}")


def dump_all_tabs(page):
    """dump 所有 tab 文本与属性。"""
    sels = [
        "[role=tab]",
        ".el-tabs__item",
        "[class*=tabs] [class*=item]",
    ]
    for s in sels:
        try:
            items = page.eval_on_selector_all(
                s,
                "els=>els.map(e=>({t:(e.textContent||'').trim(), cls:e.className}))",
            )
            if items:
                print(f"  TABS({s}):")
                for it in items:
                    if it.get("t"):
                        print(f"    {it['t']}  class={it['cls']}")
                return items
        except Exception as e:
            pass


def dump_filters(page):
    """dump 状态筛选区文本（如「全部/待下单/待收货」）。"""
    try:
        # 尝试常见侧边栏/面板选择器
        sels = [
            ".el-menu .el-submenu__title",
            ".filter-panel [class*=status]",
            ".el-card .el-tag",
            "[class*=filter] span",
            "[class*=panel] li",
            ".el-radio-group label",
            ".el-checkbox-group label",
            ".status-group span",
            # 更宽泛：左侧面板里的可点击文字
            ".el-aside span",
            ".left-panel div[class*=item]",
        ]
        for s in sels:
            items = page.eval_on_selector_all(
                s,
                "els => els.map(e => (e.textContent || '').replace(/\\s+/g,' ').trim())",
            )
            if items and any(it for it in items if len(it) > 0):
                texts = sorted(set(it for it in items if it))
                if texts[:10]:
                    print(f"  FILTERS({s}): {texts[:20]}")
    except Exception as e:
        pass


def dump_table_headers(page):
    """dump 表格列头。"""
    sels = ["th", "[role=columnheader]", ".el-table th", ".ant-table-thead th"]
    for s in sels:
        try:
            items = page.eval_on_selector_all(
                s, "els=>els.map(e=>(e.textContent||'').trim())"
            )
            if items and any(items):
                print(f"  TH({s}): {[t for t in items if t][:30]}")
                return items
        except Exception:
            pass


def search_keywords(page):
    """搜索页面中含指定关键词的可点击元素。"""
    kws = ["齐套", "出库", "领料", "BOM", "打印", "版单", "确认收货", "提交"]
    for kw in kws:
        try:
            found = page.evaluate(f"""
                (() => {{
                    const results = [];
                    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT);
                    while(walker.nextNode()){{
                        const e = walker.currentNode;
                        if(e.textContent && e.textContent.includes('{kw}')){{
                            if(e.tagName === 'BUTTON' || e.tagName === 'A' ||
                               (e.getAttribute('role') && ['button','link','menuitem','tab'].includes(e.getAttribute('role')))){{
                                results.push((e.textContent||'').replace(/\\s+/g,' ').trim().substring(0,80) +
                                    ' | tag=' + e.tagName +
                                    ' | class=' + (e.className||'').substring(0,60) +
                                    ' | id=' + (e.id||''));
                            }}
                        }}
                        if(results.length >= 8) break;
                    }}
                    return results;
                }})()
            """)
            if found:
                print(f"  KW[{kw}]:")
                for f in found:
                    print(f"    {f}")
        except Exception as e:
            pass


def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP)
        ctx = browser.contexts[0]
        target = None
        for pg in ctx.pages:
            if "turing.patpat.shop" in pg.url:
                target = pg
                break
        np = ctx.new_page()

        # ---- 目标1：版料需求列表 ----
        demand_url = (
            BASE +
            "/mrp/supply-chain/surface-accessories-m-r-p/"
            "plate-material-management/ver-material-demand-list"
        )
        print("\n=== 版料需求列表 ===")
        np.goto(demand_url, wait_until="networkidle", timeout=30000)
        np.wait_for_timeout(3000)
        print("URL:", np.url)
        print("TITLE:", np.title())
        screenshot(np, "demand_list")

        print("--- tabs ---")
        dump_all_tabs(np)

        print("--- filters ---")
        dump_filters(np)

        print("--- table headers ---")
        dump_table_headers(np)

        print("--- keywords ---")
        search_keywords(np)

        # ---- 目标2：尝试点击「待齐套」筛选/tab ----
        print("\n=== 尝试点「待齐套」===")
        clicked = False
        try:
            # 找含"齐套"的可点击元素
            qk_el = np.query_selector('text=待齐套')
            if qk_el:
                tag = qk_el.evaluate("e => e.tagName")
                role = qk_el.evaluate("e => e.getAttribute('role') || ''")
                text = qk_el.inner_text()
                print(f"  FOUND: tag={tag} role={role} text={text}")
                qk_el.click()
                np.wait_for_timeout(2000)
                screenshot(np, "demand_ready")
                print("  CLICKED OK, after-click:")
                dump_table_headers(np)
                dump_filters(np)
                clicked = True
            else:
                # 宽搜
                all_qk = np.evaluate("""
                    [...document.querySelectorAll('*')].filter(
                      e => e.children.length === 0 &&
                           e.textContent.trim() === '待齐套'
                    ).map(e => ({tag: e.tagName, cls: e.className, id: e.id}))
                """)
                print("  NOT_FOUND clickable; raw elements:", json.dumps(all_qk, ensure_ascii=False)[:500])
        except Exception as e:
            print(f"  CLICK_ERR: {e}")

        # ---- 目标3：探索出库相关 ----
        print("\n=== 探索出库 ===")
        # 先在当前页搜出库关键词
        out_kws = ["出库", "outbound", "领料", "pick"]
        for k in out_kws:
            try:
                r = np.evaluate(f"""
                    [...document.querySelectorAll('a,[button],[role=button],span')].filter(
                      e => (e.textContent||'').includes('{k}')
                    ).map(e => (e.textContent||'').trim().substring(0,80)).slice(0,5)
                """)
                if r:
                    print(f"  KW[{k}] on current page:", r)
            except Exception:
                pass

        # 如果当前页没有出库，试几个可能的 URL
        if not clicked or True:
            out_candidates = [
                BASE + "/mrp/supply-chain/surface-accessories-m-r-p/" + x
                for x in [
                    "plate-material-management/outbound-manage",
                    "plate-material-management/pick-manage",
                    "warehouse-management/pick-outbound",
                    "inventory-management/outbound",
                    "plate-material-management/version-pick",
                ]
            ]
            for u in out_candidates:
                try:
                    np.goto(u, wait_until="domcontentloaded", timeout=10000)
                    np.wait_for_timeout(1500)
                    t = np.title()
                    url = np.url
                    if "404" not in t.lower() and "not found" not in t.lower():
                        print(f"  CANDIDATE {u.split('/')[-1]}: title={t} url={url}")
                        screenshot(np, f"out_{u.split('/')[-1]}")
                        search_keywords(np)
                except Exception as e:
                    pass

        browser.close()
    print("DONE")


if __name__ == "__main__":
    main()
