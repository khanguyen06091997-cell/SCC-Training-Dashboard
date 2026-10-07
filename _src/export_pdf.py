import asyncio,sys
from playwright.async_api import async_playwright
html,xlsx,out=sys.argv[1:4]
async def m():
    async with async_playwright() as p:
        b=await p.chromium.launch(); e=[]
        pg=await b.new_page(viewport={'width':1360,'height':900},accept_downloads=True); pg.on('pageerror',lambda x:e.append(str(x)))
        await pg.goto('file://'+html); await pg.set_input_files('#fileIn',[xlsx]); await pg.wait_for_selector('#app:not([hidden])')
        async with pg.expect_download(timeout=180000) as d: await pg.click('#pdfBtn')
        dl=await d.value; await dl.save_as(out); print('errors',e)
        await b.close()
asyncio.run(m())
