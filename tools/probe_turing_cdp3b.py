"""阶段3精简：版料需求列表深度探查（只读，不点按钮、不开新tab）。
"""
from playwright.sync_api import sync_playwright

CDP = "http://127.0.0.1:9222"
BASE = "https://turing.patpat.shop"
OUT = "C:/Users/Administrator/AppData/Local/Temp"
DEMAND_URL = (
    BASE +
    "/mrp/supply-chain/surface-accessories-m-r-p/"
    "plate-material-management/ver-material-demand-list"
)


def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP)
        ctx = browser.contexts[0]
        # 复用已打开的 tab 或新建
        target = None
        for pg in ctx.pages:
            if DEMAND_URL.split("/")[-1] in pg.url or "ver-material-demand" in pg.url:
                target = pg
                break
        if not target:
            target = ctx.new_page()
            target.goto(DEMAND_URL, wait_until="networkidle", timeout=30000)
        target.wait_for_timeout(2000)

        # ---- 1. 全文搜齐套（各种写法）----
        print("=== 齐套关键词搜索 ===")
        kws = ["齐套", "待齐套", "已齐套"]
        for kw in kws:
            r = target.evaluate(f"document.body.innerText.indexOf('{kw}')")
            print(f"  [{kw}] index={r}")

        # ---- 2. 所有表格列头（完整列表）----
        print("\n=== 表头 ===")
        headers = target.evaluate(
            "[...document.querySelectorAll('th')].map(e=>e.textContent.trim()).filter(t=>t)"
        )
        for i, h in enumerate(headers):
            print(f"  [{i}] {h}")

        # ---- 3. 滚动到最右并截图 ----
        print("\n=== 表格右侧 ===")
        target.evaluate("""
            (()=>{
                const el=document.querySelector('.el-table__body-wrapper')
                    || document.querySelector('[style*="overflow"]');
                if(el){el.scrollLeft=el.scrollWidth;return 'ok'}
                return 'skip';
            })()
        """)
        target.wait_for_timeout(800)
        p_right = f"{OUT}/_probe_demand_right.png"
        target.screenshot(path=p_right)
        print(f"  SCREENSHOT: {p_right}")

        # ---- 4. 前3行数据样本（所有列文本）----
        print("\n=== 数据行样本 ===")
        rows = target.evaluate("""
            (()=>{
                const trs=document.querySelectorAll('.el-table__body tbody tr');
                const out=[];
                for(let i=0;i<Math.min(trs.length,3);i++){
                    out.push([...trs[i].querySelectorAll('td')].map(c=>c.textContent.trim()));
                }
                return out;
            })()
        """)
        for i, row in enumerate(rows):
            print(f"  row{i}: {row[:8]}...({len(row)} cols)")

        # ---- 5. 页面所有可点击文字中搜出库/BOM/打印 ----
        print("\n=== 关键词定位 ===")
        search_kws = ["出库", "领料状态", "BOM", "打印", "版单图", "确认收货"]
        for kw in search_kws:
            r = target.evaluate(f"document.body.innerText.indexOf('{kw}')")
            if r >= 0:
                ctx_text = target.evaluate(
                    f"document.body.innerText.substring(Math.max(0,{r}-40),{r}+kw.length+40).replace(/\\n/g,' ')"
                )
                print(f"  [{kw}] FOUND at {r}: ...{ctx_text}...")
            else:
                print(f"  [{kw}] NOT_FOUND")

        browser.close()
    print("DONE")


if __name__ == "__main__":
    main()
