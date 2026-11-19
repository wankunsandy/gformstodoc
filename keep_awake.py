"""
Membuka aplikasi Streamlit Community Cloud dengan browser sungguhan (Playwright)
agar aplikasi tidak tertidur. Jika aplikasi sudah tidur, tombol
"Yes, get this app back up!" diklik otomatis.

URL diambil dari variabel lingkungan STREAMLIT_URLS (pisahkan dengan koma).
"""
import asyncio
import os
import sys

from playwright.async_api import async_playwright

URLS = [u.strip() for u in os.environ.get("STREAMLIT_URLS", "").split(",") if u.strip()]
WAKE_TEXT = "get this app back up"


async def find_wake_button(page):
    """Cari tombol bangun di halaman utama maupun di dalam iframe."""
    for frame in page.frames:
        btn = frame.locator("button", has_text=WAKE_TEXT)
        if await btn.count() > 0:
            return btn.first
    return None


async def visit(browser, url: str) -> bool:
    page = await browser.new_page()
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=120_000)
        await page.wait_for_timeout(10_000)

        btn = await find_wake_button(page)
        if btn:
            print(f"BANGUNKAN  {url}  (aplikasi sedang tidur, menekan tombol)")
            await btn.click()
            # Tunggu aplikasi selesai "di dalam oven"
            await page.wait_for_timeout(120_000)
        else:
            print(f"OK         {url}  (aplikasi sudah aktif)")
            # Biarkan koneksi WebSocket tersambung agar terhitung sebagai kunjungan
            await page.wait_for_timeout(30_000)
        return True
    except Exception as e:
        print(f"GAGAL      {url}  ({type(e).__name__}: {e})")
        return False
    finally:
        await page.close()


async def main() -> int:
    if not URLS:
        print("STREAMLIT_URLS belum diisi.")
        return 1
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        results = [await visit(browser, u) for u in URLS]
        await browser.close()
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
