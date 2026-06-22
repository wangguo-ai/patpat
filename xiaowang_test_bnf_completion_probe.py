import asyncio
import importlib.util
from pathlib import Path


SCRIPT = Path(r"D:\报销工作台\03_脚本\xiaowang_agent.py")
ORDERS = ["MCG202606180021", "MCG202606180018"]


spec = importlib.util.spec_from_file_location("xiaowang_agent", SCRIPT)
xw = importlib.util.module_from_spec(spec)
spec.loader.exec_module(xw)


async def main():
    async with xw.async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False)
        context = await browser.new_context(viewport={"width": 1400, "height": 900}, locale="zh-CN")
        page = await context.new_page()
        await xw.ensure_logged_in(page, xw.PURCHASE_URL)

        results = []
        for order_no in ORDERS:
            item = {"order_no": order_no, "found_tab": "", "can_select": False, "can_open_complete": False, "error": ""}
            try:
                tab = await xw.find_order_for_completion(page, order_no)
                item["found_tab"] = tab
                await xw.select_order_checkbox(page, order_no)
                item["can_select"] = True
                await xw.click_purchase_complete(page)
                item["can_open_complete"] = True
                shot = Path(r"D:\ai共享盘\MyBrain") / f"xiaowang_bnf_probe_{order_no}.png"
                await page.screenshot(path=str(shot), full_page=True)
                item["screenshot"] = str(shot)
                # Close the modal/drawer without submitting.
                await xw.try_click_first_visible(page, [
                    'button:has-text("取 消")',
                    'button:has-text("取消")',
                    '.ant-modal-close',
                    '.ant-drawer-close',
                ], timeout=5000)
                await xw.wait_purchase_page_ready(page)
            except Exception as exc:
                item["error"] = str(exc)
                shot = Path(r"D:\ai共享盘\MyBrain") / f"xiaowang_bnf_probe_failed_{order_no}.png"
                try:
                    await page.screenshot(path=str(shot), full_page=True)
                    item["screenshot"] = str(shot)
                except Exception:
                    pass
            results.append(item)

        await browser.close()
        for item in results:
            print(item)


if __name__ == "__main__":
    asyncio.run(main())
