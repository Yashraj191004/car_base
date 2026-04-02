import asyncio
from pathlib import Path
from playwright.async_api import async_playwright
import aiohttp

DOWNLOAD_DIR = Path(__file__).parent.parent / "manuals"
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

async def main():
    """
    Asynchronously scrapes car manual PDFs from startmycar.com and downloads them.
    """
    async with aiohttp.ClientSession() as session:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)
            page = await browser.new_page()
            
            # Step 1: Extract all makes with slugs
            await page.goto("https://www.startmycar.com/pickmake")
            makes = await page.evaluate('''
                () => {
                    const links = document.querySelectorAll('a[href*="/pickmodel"]');
                    return Array.from(links).map(a => {
                        const href = a.getAttribute('href');
                        const make_slug = href.split('/')[1];
                        return {
                            name: a.textContent.trim(),
                            make_slug: make_slug
                        };
                    }).filter((v, i, a) => a.findIndex(x => x.make_slug === v.make_slug) === i);
                }
            ''')
            
            print(f"Found {len(makes)} makes")
            
            # Step 2: Extract models for each make and download PDFs
            for make in makes:
                print(f"\nProcessing make: {make['name']}")
                
                await asyncio.sleep(2)  # Rate limiting
                
                # Get models
                await page.goto(f"https://www.startmycar.com/{make['make_slug']}/pickmodel")
                models = await page.evaluate('''
                    () => {
                        const links = document.querySelectorAll('.margin-medium.ColumnsMax5 a');
                        
                        return Array.from(links).map(a => {
                            const href = a.getAttribute('href');
                            const parts = href.split('/');
                            
                            return {
                                name: a.textContent.trim(),
                                model_slug: parts[2]
                            };
                        });
                    }
                ''')
                
                print(f"Found {len(models)} models")
                
                # Step 3: Download PDFs for each model
                for model in models:
                    manual_page_url = f"https://www.startmycar.com/{make['make_slug']}/{model['model_slug']}/info/manuals"
                    await page.goto(manual_page_url)
                    
                    # 2. Extract links for all years listed on this page
                    # Look for the container that lists the years (usually a set of links)
                    year_links = await page.evaluate('''
                        () => {
                            // Adjust this selector based on the actual HTML of the year list
                            const links = document.querySelectorAll('a[href*="/manuals/"]');
                            return Array.from(links).map(a => a.href);
                        }
                    ''')
                    
                    # If no year links found, try to get PDFs directly from current page
                    if not year_links:
                        year_links = [manual_page_url]
                    
                    # 3. For each year link, go to it and download the PDF
                    for year_url in year_links:
                        await asyncio.sleep(1) # Be nice to the server
                        await page.goto(year_url)
                        year = year_url.split('/')[-1] or year_url.split('/')[-2]
                        
                        # Now find the PDF on THIS specific year page
                        pdf_links = await page.evaluate('''
                            () => Array.from(document.querySelectorAll('a[href$=".pdf"]')).map(a => a.href)
                        ''')
                        
                        for pdf_url in pdf_links:
                            # Create a more specific folder structure
                            original_filename = pdf_url.split("/")[-1].split("?")[0]
                            clean_make = make['name'].replace(" ", "_")
                            clean_model = model['name'].replace(" ", "_")
                            
                            unique_filename = f"{clean_make}_{clean_model}_{year}_{original_filename}"
                            
                            await download_pdf(pdf_url, unique_filename, session)
            
            await browser.close()

async def download_pdf(url, filename, session):
    """Download PDF to organized folder structure"""
    try:
        filepath = DOWNLOAD_DIR / filename

        if filepath.exists():
            print(f"Skipping (already exists): {filename}")
            return
        
        async with session.get(url, timeout=aiohttp.ClientTimeout(total=30)) as resp:
            if resp.status == 200:
                with open(filepath, 'wb') as f:
                    f.write(await resp.read())
                print(f"Downloaded: {filepath}")
    except Exception as e:
        print(f"Error downloading {url}: {e}")

if __name__ == "__main__":
    asyncio.run(main())
