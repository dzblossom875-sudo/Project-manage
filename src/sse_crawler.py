"""
上交所ABS项目爬虫 - Playwright版本
基于真实页面结构: 品种筛选是链接点击，翻页用 a.paging_next / button.next-page
停止条件: 与基准文件重复项目累计达到阈值即停
"""

import random
from typing import List, Dict

from models import ABSProject

KEYWORDS = ["持有型", "机构间"]


class SSECrawler:

    URL = "https://bond.sse.com.cn/bridge/information/#"
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
            context = browser.new_context(
                user_agent=self._random_ua(),
                viewport={"width": 1920, "height": 1080},
            )
            page = context.new_page()

            try:
                print("正在访问上交所页面...")
                page.goto(self.URL, wait_until="domcontentloaded", timeout=30000)
                page.wait_for_timeout(5000)

                self._select_abs_type(page)
                page.wait_for_timeout(3000)

                total_pages = self._get_total_pages(page)
                print(f"ABS品种共 {total_pages} 页, 开始逐页筛选持有型项目...")

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

        print(f"上交所共获取 {len(all_projects)} 个持有型ABS项目")
        return all_projects

    def _select_abs_type(self, page):
        """点击'资产支持证'链接筛选ABS品种"""
        try:
            abs_link = page.locator("text=资产支持证").first
            if abs_link.is_visible(timeout=3000):
                abs_link.click()
                print("  已点击'资产支持证券'筛选")
                page.wait_for_timeout(3000)
            else:
                print("  未找到ABS筛选链接, 将爬取全部数据")
        except Exception as e:
            print(f"  ABS筛选失败: {e}, 将爬取全部数据并本地过滤")

    def _get_total_pages(self, page) -> int:
        """从页面获取总页数 (上交所格式: '共 N 页')"""
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

        try:
            links = page.locator("a.js_pjinfos_container_paginationNumLink").all()
            max_p = 1
            for link in links:
                t = link.inner_text().strip()
                if t.isdigit():
                    max_p = max(max_p, int(t))
            if max_p > 1:
                return max_p
        except Exception:
            pass

        return 1

    def _extract_page(self, page) -> List[ABSProject]:
        """提取当前页的持有型项目, 本地过滤关键词"""
        projects = []

        try:
            page.wait_for_selector("table", timeout=10000)
            rows = page.locator("table tr").all()

            for row in rows[1:]:
                cells = row.locator("td").all()
                if len(cells) < 8:
                    continue

                try:
                    texts = [c.inner_text().strip() for c in cells]
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
                        exchange="上交所",
                    )

                    if self.baseline_index:
                        from baseline_reader import check_duplicate
                        if check_duplicate(self.baseline_index, "上交所", bond_name):
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
        """点击下一页 (真实选择器: a.paging_next 或 button.next-page)"""
        selectors = [
            "a.paging_next",
            "button.next-page",
            "a:has-text('下一页')",
        ]
        for sel in selectors:
            try:
                btn = page.locator(sel).first
                if btn.is_visible(timeout=2000):
                    cls = btn.get_attribute("class") or ""
                    if "disabled" in cls:
                        return False
                    btn.click()
                    page.wait_for_timeout(2000)
                    return True
            except Exception:
                continue
        return False

    def _random_delay(self, page):
        page.wait_for_timeout(random.randint(1500, 2500))

    def _random_ua(self) -> str:
        uas = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        ]
        return random.choice(uas)


def crawl_sse(baseline_index: Dict[str, dict] = None, headless: bool = True) -> List[ABSProject]:
    crawler = SSECrawler(baseline_index=baseline_index, headless=headless)
    return crawler.crawl()


if __name__ == "__main__":
    projects = crawl_sse(headless=False)
    for p in projects:
        print(f"[{p.exchange}] {p.bond_name} | {p.status} | {p.accept_date}")
