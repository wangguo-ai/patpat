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
        await page.wait_for_timeout(1000)
        info = await page.locator(".ant-modal").evaluate(
            """modal => {
                const headers = Array.from(modal.querySelectorAll('th')).map((th, i) => ({i, text: th.innerText.trim()}));
                const inputs = Array.from(modal.querySelectorAll('input')).map((input, i) => {
                    const r = input.getBoundingClientRect();
                    return {
                        i,
                        value: input.value,
                        placeholder: input.getAttribute('placeholder') || '',
                        visible: r.width > 0 && r.height > 0,
                        x: Math.round(r.x),
                        y: Math.round(r.y),
                        width: Math.round(r.width)
                    };
                });
                const text = modal.innerText;
                return {headers, inputs, text};
            }"""
        )
        print(info)
        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
