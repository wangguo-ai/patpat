"""阶段4：纯 evaluate 探查版料需求列表（不用 screenshot，避免段错误）。
只读 DOM 数据，不触发任何写入操作。
"""
from playwright.sync_api import sync_playwright
import json

CDP = "http://127.0.0.1:9222"
BASE = "https://turing.patpat.shop"
DEMAND_URL = (
    BASE +
    "/mrp/supply-chain/surface-accessories-m-r-p/"
    "plate-material-management/ver-material-demand-list"
)


def safe_eval(page, js, label=""):
    """安全 evaluate，出错返回 None 而不是崩。"""
    try:
        return page.evaluate(js)
    except Exception as e:
        print(f"  EVAL_ERR[{label}]: {e}")
        return None


def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP)
        ctx = browser.contexts[0]
        # 找到版料需求 tab 或新建
        target = None
        for pg in ctx.pages:
            if "ver-material-demand" in pg.url:
                target = pg
                break
        if not target:
            target = ctx.new_page()
            target.goto(DEMAND_URL, wait_until="networkidle", timeout=30000)
            target.wait_for_timeout(2000)
        print("TARGET:", target.url)

        # ---- 1. 全文搜"齐套" ----
        print("\n=== 1. 齐套搜索 ===")
        for kw in ["齐套", "待齐套", "已齐套", "齐套检查"]:
            idx = safe_eval(target, f"document.body.innerText.indexOf('{kw}')", f"search_{kw}")
            print(f"  [{kw}] index={idx}")

        # ---- 2. 全部表格列头 ----
        print("\n=== 2. 表头 (th) ===")
        headers = safe_eval(target,
            "[...document.querySelectorAll('th')].map(e=>e.textContent.trim()).filter(t=>t)",
            "headers")
        if headers:
            for i, h in enumerate(headers):
                print(f"  [{i:2d}] {h}")

        # ---- 3. 前3行数据（每行所有列值）----
        print("\n=== 3. 数据行样本 ===")
        rows = safe_eval(target, """
            (()=>{
                const trs=document.querySelectorAll('.el-table__body tbody tr');
                const out=[];
                for(let i=0;i<Math.min(trs.length,3);i++){
                    const tds=[...trs[i].querySelectorAll('td')];
                    out.push(tds.map(c=>c.textContent.trim()));
                }
                return out;
            })()
        """, "rows")
        if rows:
            for i, row in enumerate(rows):
                print(f"  row{i}: cols={len(row)} | first8={row[:8]}")

        # ---- 4. 关键词在可点击元素中的位置 ----
        print("\n=== 4. 可点击元素关键词 ===")
        kws = ["出库", "领料", "BOM", "打印", "版单", "收货", "齐套"]
        for kw in kws:
            r = safe_eval(target, f"""
                (({{
                    const all=[...document.querySelectorAll('a,button,[role=button],[role=link],span.el-button')];
                    return all.filter(e=>(e.textContent||'').includes('{kw}'))
                               .map(e=>({{tag:e.tagName,text:(e.textContent||'').trim().substring(0,50),
                                          cls:(e.className||'').substring(0,40)}}))
                               .slice(0,5);
                }}))()
            """, f"clickable_{kw}")
            if r:
                print(f"  [{kw}]:", json.dumps(r, ensure_ascii=False))

        # ---- 5. 页面中"齐套"的上下文（如果在innerText里找到的话）----
        qk_idx = safe_eval(target, "document.body.innerText.indexOf('齐套')")
        if qk_idx is not None and isinstance(qk_idx, int) and qk_idx >= 0:
            start = max(0, qk_idx - 60)
            end = qk_idx + 80
            ctx = safe_eval(target,
                f"document.body.innerText.substring({start},{end}).replace(/\\n/g,' ')",
                "qk_context")
            print(f"\n=== 5. 齐套上下文 ===\n  ...{ctx}...")

        # ---- 6. 左侧筛选面板文字（如果有类似采购页的状态面板）----
        print("\n=== 6. 左侧/筛选区域文字 ===")
        panel_texts = safe_eval(target, """
            (()=>{
                // 尝试找左侧面板或状态筛选区
                const panels=document.querySelectorAll('.el-aside,.el-menu,[class*=filter],[class*=panel],[class*=status]');
                const texts=[];
                panels.forEach(p=>{
                    const spans=p.querySelectorAll('li,span,a,label,div[class*=item]');
                    spans.forEach(s=>{
                        const t=s.textContent.trim();
                        if(t && t.length<30) texts.push(t);
                    });
                });
                return [...new Set(texts)].slice(0,40);
            })()
        """, "panels")
        if panel_texts:
            for t in panel_texts:
                print(f"  {t}")

        browser.close()
    print("\nDONE")


if __name__ == "__main__":
    main()
