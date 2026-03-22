"""
深交所ABS项目爬虫
基于真实页面结构: 数据在第二个table, 每页20行, 翻页用 a:has-text('下一页')
优先Playwright页面爬取(API已被拦截), 逐页本地过滤持有型项目
停止条件: 与基准文件重复项目累计达到阈值即停
"""

import re
import random
from typing import List, Dict

from models import ABSProject

KEYWORDS = ["持有型", "机构间"]


class SZSECrawler:

    URL = "https://bond.szse.cn/disclosure/progressinfo/index.html"
    STOP_THRESHOLD = 3

    def __init__(self, baseline_index: Dict[str, dict] = None, headless: bool = True):
        self.baseline_index = baseline_index or {}
        self.headless = headless
        self.duplicate_count = 0

    def crawl(self) -> List[ABSProject]:
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            print("错误: pip install playwright && playwright install chromium")
            return []

        all_projects = []

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=self.headless)
            ctx = browser.new_context(
                user_agent=self._random_ua(),
                viewport={"width": 1920, "height": 1080},
            )
            page = ctx.new_page()

            try:
                print("正在访问深交所页面...")
                page.goto(self.URL, wait_until="networkidle", timeout=45000)
                page.wait_for_timeout(3000)

                total_pages = self._get_total_pages(page)
                print(f"共 {total_pages} 页, 开始逐页筛选持有型项目...")

                for page_num in range(1, total_pages + 1):
                    print(f"  第 {page_num}/{total_pages} 页...")

                    projects = self._extract_page(page)
                    all_projects.extend(projects)

                    if self.duplicate_count >= self.STOP_THRESHOLD:
                        print(f"  累计 {self.duplicate_count} 个与基准重复项目, 停止爬取")
                        break

                    if page_num < total_pages:
                        if not self._go_next_page(page):
                            print("  无法翻页, 结束")
                            break

                    self._random_delay(page)

            except Exception as e:
                print(f"爬取出错: {e}")
            finally:
                browser.close()

        print(f"深交所共获取 {len(all_projects)} 个持有型ABS项目")
        return all_projects

    def _get_total_pages(self, page) -> int:
        try:
            text = page.evaluate("""() => {
                const body = document.body.innerText;
                const match = body.match(/共\\s*(\\d+)\\s*页/);
                return match ? match[1] : null;
            }""")
            if text:
                return int(text)
        except Exception:
            pass
        return 1

    def _extract_page(self, page) -> List[ABSProject]:
        """
        从当前页面提取持有型项目
        深交所结构: table[0]=表头, table[1]=数据(20行)
        每行: 序号, 债券名称, 发行人/管理人, 债券品种, 申报规模(亿元), 项目状态, 更新日期, 受理日期
        """
        projects = []

        try:
            page.wait_for_selector("table", timeout=10000)
            tables = page.locator("table").all()

            data_table = tables[1] if len(tables) > 1 else tables[0]
            rows = data_table.locator("tr").all()

            for row in rows:
                cells = row.locator("td").all()
                if len(cells) < 8:
                    continue

                try:
                    texts = [self._clean_html(c.inner_text()) for c in cells]
                    bond_name = texts[1]

                    if not any(kw in bond_name for kw in KEYWORDS):
                        continue

                    project = ABSProject(
                        bond_name=bond_name,
                        manager=texts[2],
                        bond_type=texts[3],
                        amount=texts[4],
                        status=texts[5],
                        update_date=texts[6],
                        accept_date=texts[7],
                        exchange="深交所",
                    )

                    if self.baseline_index:
                        from baseline_reader import check_duplicate
                        if check_duplicate(self.baseline_index, "深交所", bond_name):
                            self.duplicate_count += 1
                        else:
                            self.duplicate_count = 0

                    projects.append(project)
                    print(f"    [{project.status}] {bond_name[:50]}")

                except Exception:
                    continue

        except Exception as e:
            print(f"  提取出错: {e}")

        return projects

    def _go_next_page(self, page) -> bool:
        try:
            btn = page.locator("a:has-text('下一页')").first
            if btn.is_visible(timeout=2000):
                cls = btn.get_attribute("class") or ""
                if "disabled" in cls:
                    return False
                btn.click()
                page.wait_for_timeout(2000)
                return True
        except Exception:
            pass
        return False

    def _clean_html(self, text: str) -> str:
        clean = re.sub(r"<[^>]+>", "", str(text))
        return clean.strip()

    def _random_delay(self, page):
        page.wait_for_timeout(random.randint(1500, 2500))

    def _random_ua(self) -> str:
        uas = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        ]
        return random.choice(uas)


def crawl_szse(baseline_index: Dict[str, dict] = None, headless: bool = True) -> List[ABSProject]:
    crawler = SZSECrawler(baseline_index=baseline_index, headless=headless)
    return crawler.crawl()


if __name__ == "__main__":
    projects = crawl_szse(headless=False)
    for p in projects:
        print(f"[{p.exchange}] {p.bond_name} | {p.status} | {p.accept_date}")
