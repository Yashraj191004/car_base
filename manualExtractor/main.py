import asyncio
from pathlib import Path
from playwright.async_api import async_playwright
import aiohttp

DOWNLOAD_DIR = Path(__file__).parent.parent / "manuals"
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

async def process_make(make, browser, session, pdf_tasks, sem):
    """
    Process a single make: get models, years, and collect PDF URLs.
    """
    async with sem:  # Limit concurrent processing
        page = await browser.new_page()
        try:
            make_name = make.get('name') or make['make_slug']
            print(f"\nProcessing make: {make_name}")
            
            await asyncio.sleep(2)  # Rate limiting
            
            # Get models
            await page.goto(f"https://www.startmycar.com/{make['make_slug']}/pickmodel", timeout=60000, wait_until='domcontentloaded')
            
            models = await page.evaluate(f'''
                () => {{
                    const makeSlug = "{make['make_slug']}";
                    const anchors = Array.from(document.querySelectorAll(`a[href^="/${{makeSlug}}/"], a[href^="https://www.startmycar.com/${{makeSlug}}/"]`));
                    
                    return anchors
                        .map(a => {{
                            const href = a.getAttribute('href');
                            const url = new URL(href, location.origin);
                            const segments = url.pathname.split('/').filter(Boolean);
                            if (segments.length !== 2 || segments[1] === 'pickmodel') return null;
                            return {{
                                name: a.textContent.trim() || segments[1],
                                model_slug: segments[1]
                            }};
                        }})
                        .filter(Boolean)
                        .filter((v, i, arr) => arr.findIndex(x => x.model_slug === v.model_slug) === i);
                }}
            ''')
            
            print(f"Found {len(models)} models for {make_name}")
            
            # Collect PDFs for each model
            for model in models:
                try:
                    manual_page_url = f"https://www.startmycar.com/{make['make_slug']}/{model['model_slug']}/info/manuals"
                    await page.goto(manual_page_url, timeout=60000, wait_until='domcontentloaded')
                    await asyncio.sleep(1)
                    
                    # Extract links for all years listed on this page
                    year_links = await page.evaluate(r'''
                        () => {
                            const links = Array.from(document.querySelectorAll('a[href*="/info/manuals/"]'));
                            return links
                                .map(a => a.href)
                                .filter(href => /(\/info\/manuals\/\d{4}$|\/\d{4}\/info\/manuals$)/.test(href));
                        }
                    ''')
                    
                    # If no year links found, try to get PDFs directly from current page
                    if not year_links:
                        year_links = [manual_page_url]
                    
                    # For each year link, go to it and collect PDF URLs
                    for year_url in year_links:
                        try:
                            await asyncio.sleep(1)  # Be nice to the server
                            await page.goto(year_url, timeout=60000, wait_until='domcontentloaded')
                            year = year_url.split('/')[-1] or year_url.split('/')[-2]
                            
                            # Find the PDF on THIS specific year page
                            pdf_links = await page.evaluate('''
                                () => Array.from(document.querySelectorAll('a[href$=".pdf"]')).map(a => a.href)
                            ''')
                            
                            for pdf_url in pdf_links:
                                # Create filename
                                original_filename = pdf_url.split("/")[-1].split("?")[0]
                                clean_make = (make.get('name') or make['make_slug']).replace(" ", "_")
                                clean_model = (model.get('name') or model['model_slug']).replace(" ", "_")
                                
                                unique_filename = f"{clean_make}_{clean_model}_{year}_{original_filename}"
                                pdf_tasks.append((pdf_url, unique_filename))
                                
                        except Exception as e:
                            print(f"Error processing year {year_url}: {e}")
                            continue
                            
                except Exception as e:
                    print(f"Error processing model {model.get('name') or model['model_slug']}: {e}")
                    continue
                    
        except Exception as e:
            print(f"Error processing make {make_name}: {e}")
        finally:
            await page.close()

async def main():
    """
    Asynchronously scrapes car manual PDFs from startmycar.com and downloads them.
    """
    async with aiohttp.ClientSession() as session:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)
            
            # Step 1: Extract all makes with slugs
            page = await browser.new_page()
            await page.goto("https://www.startmycar.com/pickmake", timeout=60000, wait_until='domcontentloaded')
            await page.wait_for_selector('a[href*="/pickmodel"]', timeout=15000, state='attached')
            makes = await page.evaluate('''
                () => {
                    const links = document.querySelectorAll('a[href*="/pickmodel"]');
                    return Array.from(links).map(a => {
                        const href = a.getAttribute('href');
                        const make_slug = href.split('/')[1];
                        const text = a.textContent.trim();
                        if (text) {
                            return {
                                name: text,
                                make_slug: make_slug
                            };
                        }
                        const image = a.querySelector('img');
                        const alt = image ? image.getAttribute('alt') : '';
                        return {
                            name: alt.trim() || make_slug,
                            make_slug: make_slug
                        };
                    }).filter((v, i, a) => a.findIndex(x => x.make_slug === v.make_slug) === i);
                }
            ''')
            await page.close()
            
            print(f"Found {len(makes)} makes")
            
            # Step 2: Process makes concurrently
            pdf_tasks = []  # Shared list to collect all PDF download tasks
            sem = asyncio.Semaphore(5)  # Limit to 5 concurrent makes
            
            tasks = [process_make(make, browser, session, pdf_tasks, sem) for make in makes[0:5]]
            await asyncio.gather(*tasks)
            
            # Step 3: Download all collected PDFs
            print(f"\nStarting download of {len(pdf_tasks)} PDFs...")
            for pdf_url, filename in pdf_tasks:
                await download_pdf(pdf_url, filename, session)
            
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
