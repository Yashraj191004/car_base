import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        url = 'https://www.startmycar.com/ford/pickmodel'
        print('goto', url)
        try:
            await page.goto(url, timeout=60000, wait_until='domcontentloaded')
            print('loaded')
            links = await page.evaluate("() => Array.from(document.querySelectorAll('a[href]')).slice(0,100).map(a => ({href: a.href, text: a.textContent.trim(), classes: a.className}))")
            print('links sample', links)
        except Exception as e:
            print('error', e)
        await browser.close()

asyncio.run(main())
