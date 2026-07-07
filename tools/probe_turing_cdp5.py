"""阶段5：精确提取版料需求列表的关键业务数据。
修正 Element UI 表格选择器，取行数据、齐套值域、可点击操作。
"""
from playwright.sync_api import sync_playwright
import json

CDP = "http://127.0.0.1:9222"


def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP)
        ctx = browser.contexts[0]
        target = None
        for pg in ctx.pages:
            if "ver-material-demand" in pg.url:
                target = pg
                break
        if not target:
            print("ERROR: demand page not found")
            return
        print("TARGET:", target.url)

        # ---- A. 行数据（修正选择器：Element UI 用 el-table__row）----
        print("\n=== A. 行数据 ===")
        row_data = target.evaluate("""
            (()=>{
                const rows=document.querySelectorAll('.el-table__body .el-table__row');
                const out=[];
                for(let i=0;i<Math.min(rows.length,5);i++){
                    const cells=[...rows[i].querySelectorAll('td')];
                    out.push(cells.map(c=>c.textContent.trim()));
                }
                return {count:rows.length, rows:out};
            })()
        """)
        if row_data:
            print(f"  total_rows={row_data['count']}")
            for i, r in enumerate(row_data["rows"]):
                # 只印每行的关键字段（序号、需求单号、设计款、状态类列）
                key_cols = [r[j] if j < len(r) else "?" for j in [0, 1, 2, 6, 9, 10]]
                print(f"  row{i}: {key_cols}")

        # ---- B. 齐套情况位列的所有唯一值 ----
        print("\n=== B. 齐套情况位值域 ===")
        qk_values = target.evaluate("""
            (()=>{
                // 找表头中含"齐套"的列索引
                const ths=[...document.querySelectorAll('th')];
                let colIdx=-1;
                ths.forEach((th,i)=>{
                    if(th.textContent.includes('齐套')) colIdx=i;
                });
                if(colIdx<0) return {error:'header_not_found'};
                // 取该列所有行的值
                const tds=document.querySelectorAll('.el-table__body .el-table__row');
                const values=new Set();
                tds.forEach(tr=>{
                    const cells=[...tr.querySelectorAll('td')];
                    if(colIdx < cells.length) values.add(cells[colIdx].textContent.trim());
                });
                return {colIndex:colIdx, headerName:ths[colIdx].textContent.trim(),
                        uniqueValues:[...values], totalRows:tds.length};
            })()
        """)
        if qk_values and "error" not in qk_values:
            print(f"  column_index={qk_values.get('colIndex')}")
            print(f"  header_name={qk_values.get('headerName')}")
            print(f"  unique_values={qk_values.get('uniqueValues')}")
            print(f"  sample_rows={qk_values.get('totalRows')}")

        # ---- C. 领料状态列值域 ----
        print("\n=== C. 领料状态值域 ===")
        pick_values = target.evaluate("""
            (()=>{
                const ths=[...document.querySelectorAll('th')];
                let colIdx=-1;
                ths.forEach((th,i)=>{if(th.textContent.includes('领料状态')) colIdx=i;});
                if(colIdx<0) return {error:'not_found'};
                const vals=new Set();
                document.querySelectorAll('.el-table__body .el-table__row').forEach(tr=>{
                    const cells=[...tr.querySelectorAll('td')];
                    if(colIdx<cells.length) vals.add(cells[colIdx].textContent.trim());
                });
                return {colIndex:colIdx, uniqueValues:[...vals]};
            })()
        """)
        if pick_values and "error" not in pick_values:
            print(f"  column_index={pick_values.get('colIndex')}")
            print(f"  unique_values={pick_values.get('uniqueValues')}")

        # ---- D. 操作列内容（最后一列或标记为"操作"的列）----
        print("\n=== D. 操作列 ===")
        op_col = target.evaluate("""
            (()=>{
                const ths=[...document.querySelectorAll('th')];
                let colIdx=-1;
                ths.forEach((th,i)=>{if(th.textContent.includes('操作')) colIdx=i;});
                if(colIdx<0) return {error:'not_found'};
                const ops=new Set();
                document.querySelectorAll('.el-table__body .el-table__row').forEach(tr=>{
                    const cells=[...tr.querySelectorAll('td')];
                    if(colIdx<cells.length){
                        const html=cells[colIdx].innerHTML.substring(0,200);
                        ops.add(html);
                    }
                });
                return {colIndex:colIdx, samples:[...ops].slice(0,5)};
            })()
        """)
        if op_col and "error" not in op_col:
            print(f"  col={op_col.get('colIndex')} samples:")
            for s in op_col.get("samples", []):
                print(f"    {s[:120]}")

        # ---- E. 全页搜出库/BOM/打印等关键词（简化JS避免语法问题）----
        print("\n=== E. 全页关键词 ===")
        kws = ["出库", "领料", "BOM", "打印", "版单", "收货", "提交", "入库"]
        body_text = target.evaluate("document.body.innerText")
        for kw in kws:
            positions = []
            start = 0
            while True:
                idx = body_text.find(kw, start)
                if idx == -1:
                    break
                positions.append(idx)
                start = idx + 1
                if len(positions) >= 5:
                    break
            if positions:
                ctx_samples = []
                for pos in positions[:3]:
                    ctx_start = max(0, pos - 30)
                    ctx_end = min(len(body_text), pos + len(kw) + 30)
                    ctx_samples.append(body_text[ctx_start:ctx_end].replace("\n", " "))
                print(f"  [{kw}] count={len(positions)} pos={positions[:3]}")
                for cs in ctx_samples:
                    print(f"    ...{cs}...")
            else:
                print(f"  [{kw}] NOT_FOUND")

        browser.close()
    print("DONE")


if __name__ == "__main__":
    main()
