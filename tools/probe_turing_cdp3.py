"""阶段3：版料需求列表深度探查 + 需求详情页探索。
只读：截图、读 DOM、不触发写入。
"""
import json, sys
from playwright.sync_api import sync_playwright

CDP = "http://127.0.0.1:9222"
BASE = "https://turing.patpat.shop"
OUT = "C:/Users/Administrator/AppData/Local/Temp"
DEMAND_URL = (
    BASE +
    "/mrp/supply-chain/surface-accessories-m-r-p/"
    "plate-material-management/ver-material-demand-list"
)


def screenshot(page, name):
    p = f"{OUT}/_probe_{name}.png"
    page.screenshot(path=p)
    return p


def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP)
        ctx = browser.contexts[0]
        np = ctx.new_page()

        # ---- 1. 回到版料需求列表 ----
        print("=== 1. 版料需求列表 ===")
        np.goto(DEMAND_URL, wait_until="networkidle", timeout=30000)
        np.wait_for_timeout(3000)

        # 1a. 全文搜"齐套"（各种可能写法）
        for kw in ["齐套", "待齐套", "已齐套", "齐套检查", "ready", "complete"]:
            try:
                found = np.evaluate(f"""
                    (() => {{
                        const text = document.body.innerText;
                        const idx = text.indexOf('{kw}');
                        if (idx === -1) return null;
                        // 取前后各 80 字符作为上下文
                        const start = Math.max(0, idx - 60);
                        const end = Math.min(text.length, idx + kw.length + 60);
                        return {{pos: idx, context: text.substring(start, end).replace(/\\n/g,' ')}};
                    }})()
                """)
                if found:
                    print(f"  FOUND [{kw}] at pos={found['pos']}: ...{found['context']}...")
            except Exception:
                pass

        # 1b. 表格右侧列头（滚动表格容器后截图）
        try:
            # 找表格容器的 scrollLeft 最大值并滚动
            table_container = np.query_selector(".el-table__body-wrapper") or np.query_selector(
                ".ant-table-body"
            ) or np.query_selector("table").evaluate(
                "e => e.parentElement.parentElement"
            )
            if table_container:
                # 先截左侧（原始）
                s1 = screenshot(np, "demand_left")

                # 滚动到最右
                np.evaluate("""
                    (() => {
                        const el = document.querySelector('.el-table__body-wrapper')
                            || document.querySelector('table').closest('[style*="overflow"]');
                        if(el){el.scrollLeft=el.scrollWidth;return 'scrolled';}
                        return 'not_found';
                    })()
                """)
                np.wait_for_timeout(500)
                s2 = screenshot(np, "demand_right")
                print(f"  TABLE_SCROLL: left={s1}, right={s2}")
        except Exception as e:
            print(f"  SCROLL_ERR: {e}")

        # 1c. dump 所有表头文本（完整）
        headers = np.evaluate("""
            (() => {
                const ths = document.querySelectorAll('th');
                return [...ths].map(th => th.textContent.trim()).filter(t => t);
            })()
        """)
        print(f"  ALL_HEADERS ({len(headers)}):", [h for h in headers])

        # 1d. dump 第一行数据样本（前 5 行的关键列值）
        row_sample = np.evaluate("""
            (() => {
                const rows = document.querySelectorAll('.el-table__body tr, tbody tr');
                const samples = [];
                for(let i=0; i<Math.min(rows.length,3); i++){
                    const cells = rows[i].querySelectorAll('td');
                    samples.push([...cells].map(c => c.textContent.trim()).filter(t => t));
                }
                return samples;
            })()
        """)
        print(f"  ROW_SAMPLE:")
        for i, r in enumerate(row_sample):
            print(f"    row{i}: {r[:15]}...")

        # ---- 2. 点击第一行需求单号进入详情 ----
        print("\n=== 2. 需求详情页 ===")
        detail_clicked = False
        try:
            # 找第一个需求单号链接（MR2026... 格式）
            first_link = np.query_selector("td a[href*='MR'], td a:has-text('MR')")
            if not first_link:
                # 宽搜任何可点击的需求单号文本
                first_link = np.query_selector("text=/^MR\\d+/")
            if first_link:
                link_text = first_link.inner_text()
                link_href = first_link.get_attribute("href") or ""
                print(f"  FIRST_LINK: text={link_text} href={link_href}")
                # 新开 tab 去详情（保持列表页不动）
                detail_page = ctx.new_page()
                if link_href:
                    full_url = BASE + link_href if link_href.startswith("/") else link_href
                    detail_page.goto(full_url, wait_until="networkidle", timeout=20000)
                else:
                    first_link.click()
                    # 等新 tab 打开
                    np.wait_for_timeout(3000)
                    for pg in ctx.pages:
                        if pg != np and "turing.patpat.shop" in pg.url and pg.url != DEMAND_URL:
                            detail_page = pg
                            break
                detail_page.wait_for_timeout(3000)
                print(f"  DETAIL_URL: {detail_page.url}")
                print(f"  DETAIL_TITLE: {detail_page.title()}")
                sd = screenshot(detail_page, "demand_detail")
                print(f"  DETAIL_SCREENSHOT: {sd}")

                # 2a. 详情页搜关键词
                kws = ["齐套", "出库", "领料", "BOM", "打印", "版单", "确认收货",
                       "提交", "入库", "采购"]
                for kw in kws:
                    try:
                        r = detail_page.evaluate(f"""
                            (() => {{
                                const results = [];
                                const walker = document.createTreeWalker(document.body,
                                    NodeFilter.SHOW_ELEMENT);
                                while(walker.nextNode()){{
                                    const e = walker.currentNode;
                                    if(e.textContent && e.textContent.includes('{kw}')
                                        && (e.tagName==='BUTTON'||e.tagName==='A'
                                            ||['button','link','menuitem'].includes(e.getAttribute('role')||''))){{
                                        results.push(e.textContent.trim().substring(0,60)
                                            + ' | tag='+e.tagName+' | class='+(e.className||'').substring(0,40));
                                    }}
                                }}
                                return results.slice(0,5);
                            }})()
                        """)
                        if r:
                            print(f"  DETAIL_KW[{kw}]:", r)
                    except Exception:
                        pass

                # 2b. 截详情页 tab 区域
                detail_tabs = detail_page.eval_on_selector_all("[role=tab], .el-tabs__item", 
                    "els=>els.map(e=>({t:e.textContent.trim(),c:e.className}))")
                if detail_tabs:
                    print(f"  DETAIL_TABS:", detail_tabs)

                detail_clicked = True
        except Exception as e:
            print(f"  DETAIL_ERR: {e}")

        browser.close()
    print("DONE")


if __name__ == "__main__":
    main()
