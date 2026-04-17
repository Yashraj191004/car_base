import asyncio
import aiohttp

async def main():
    url = 'https://manuals.startmycar.com/published/Honda-Accord_2023_EN-US_US_08869b4b25.pdf'
    async with aiohttp.ClientSession() as session:
        async with session.get(url, timeout=aiohttp.ClientTimeout(total=30)) as resp:
            print('status', resp.status)
            print('content-type', resp.headers.get('content-type'))
            data = await resp.read()
            print('len', len(data))

asyncio.run(main())
