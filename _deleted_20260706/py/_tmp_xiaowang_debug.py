import asyncio
import importlib.util
from pathlib import Path

from playwright.async_api import async_playwright


SCRIPT = Path(r"D:\报销工作台\03_脚本\xiaowang_agent.py")
spec = importlib.util.spec_from_file_location("xiaowang_agent", SCRIPT)
xw = importlib.util.module_from_spec(spec)
spec.loader.exec_module(xw)


async def main():
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False)
        context = await browser.new_context(
            accept_downloads=True,
            viewport={"width": 1400, "height": 900},
            locale="zh-CN",
        )
        page = await context.new_page()
        await xw.ensure_logged_in(page, xw.PURCHASE_URL)
        await xw.open_purchase_all(page)
        await xw.switch_purchase_status(page, "待下单")
        await page.screenshot(
            path=r"D:\ai共享盘\MyBrain\xiaowang_purchase_debug.png",
            full_page=True,
        )
        items = await page.locator(
            "button, a, [role=button], [title], [aria-label]"
        ).evaluate_all(
            """els => els.slice(0, 250).map((e, i) => ({
                i,
                tag: e.tagName,
                text: (e.innerText || e.textContent || '').trim(),
                title: e.getAttribute('title') || '',
                aria: e.getAttribute('aria-label') || '',
                cls: String(e.className || '')
            }))"""
        )
        for item in items:
            text = (item["text"] or item["title"] or item["aria"] or "").strip()
            if text:
                print(
                    item["i"],
                    item["tag"],
                    "text=",
                    text[:120].replace("\n", " | "),
                    "title=",
                    item["title"],
                    "aria=",
                    item["aria"],
                )
        await browser.close()


asyncio.run(main())
