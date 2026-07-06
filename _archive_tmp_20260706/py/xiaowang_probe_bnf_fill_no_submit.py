import asyncio
import importlib.util
from pathlib import Path


SCRIPT = Path(r"D:\报销工作台\03_脚本\xiaowang_agent.py")
ORDER = "MCG202606180021"

spec = importlib.util.spec_from_file_location("xiaowang_agent", SCRIPT)
xw = importlib.util.module_from_spec(spec)
spec.loader.exec_module(xw)


async def main():
    async with xw.async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False)
        context = await browser.new_context(viewport={"width": 1400, "height": 900}, locale="zh-CN")
        page = await context.new_page()
        await xw.ensure_logged_in(page, xw.PURCHASE_URL)
        await page.goto(xw.PURCHASE_URL, wait_until="domcontentloaded", timeout=90000)
        await page.wait_for_timeout(8000)
        await xw.find_order_for_completion(page, ORDER)
        await xw.select_order_checkbox(page, ORDER)
        await xw.click_purchase_complete(page)
        await xw.fill_completion_form(page, 1, 17.05)
        values = await page.evaluate(
            """() => {
                const modal = document.querySelector('.ant-modal, .ant-drawer');
                const headers = Array.from(modal.querySelectorAll('th')).map(th => {
                    const r = th.getBoundingClientRect();
                    return {text: (th.innerText || th.textContent || '').trim(), left: r.left, right: r.right};
                });
                const inputs = Array.from(modal.querySelectorAll('tbody input, .ant-table-body input'));
                const byHeader = {};
                for (const key of ['实际采购数量', '实际采购单价']) {
                    const h = headers.find(x => x.text.includes(key));
                    const input = inputs.find(input => {
                        const r = input.getBoundingClientRect();
                        const center = r.left + r.width / 2;
                        return h && center >= h.left - 8 && center <= h.right + 8;
                    });
                    byHeader[key] = input ? input.value : null;
                }
                return byHeader;
            }"""
        )
        print(values)
        shot = Path(r"D:\ai共享盘\MyBrain") / f"xiaowang_bnf_fill_probe_{ORDER}.png"
        await page.screenshot(path=str(shot), full_page=True)
        print(shot)
        await xw.try_click_first_visible(page, [
            'button:has-text("取 消")',
            'button:has-text("取消")',
            '.ant-modal-close',
            '.ant-drawer-close',
        ], timeout=5000)
        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
