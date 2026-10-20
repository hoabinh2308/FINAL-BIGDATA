#!/usr/bin/env python3
"""
crawler.py - Công cụ crawl dữ liệu sản phẩm dùng chung (điều khiển bằng file config).

Hai chế độ (trường "mode" trong config):
  1. sitemap_jsonld : đọc sitemap -> mở từng trang sản phẩm -> lấy JSON-LD (schema.org Product)
  2. api            : gọi API JSON có phân trang -> ánh xạ trường theo đường dẫn dạng "a.b.0.c"

Cách chạy:
  python crawler.py config.cellphones.json --dry-run        # chỉ đếm URL, chưa tải trang sản phẩm
  python crawler.py config.cellphones.json --limit 50       # cào thử 50 sản phẩm
  python crawler.py config.cellphones.json                  # cào toàn bộ

Dữ liệu thô được cache trong thư mục raw/<name>/, chạy lại sau khi bị ngắt sẽ không tải lại
những gì đã có. Kết quả ghi ra data/<name>_<thời gian>.csv (UTF-8 có BOM, mở bằng Excel không lỗi tiếng Việt).
"""
import argparse
import csv
import gzip
import hashlib
import json
import random
import re
import sys
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup

FIELDS = [
    "source", "product_id", "name", "brand", "category", "price", "original_price",
    "currency", "availability", "rating", "review_count", "image", "url", "crawled_at",
]

DEFAULT_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
)


def log(msg):
    print(msg, file=sys.stderr, flush=True)


# --------------------------------------------------------------------------- #
# Fetcher: HTTP + retry + delay + cache + robots.txt
# --------------------------------------------------------------------------- #
class Fetcher:
    def __init__(self, cfg):
        self.base = cfg["base_url"]
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": DEFAULT_UA, "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8"})
        self.s.headers.update(cfg.get("headers", {}))
        self.dmin, self.dmax = cfg.get("delay", [1.0, 3.0])
        self.retries = cfg.get("retries", 3)
        self.timeout = cfg.get("timeout", 20)
        self.respect_robots = cfg.get("respect_robots", True)
        self.raw_dir = Path(cfg.get("raw_dir", "raw")) / cfg["name"]
        self.raw_dir.mkdir(parents=True, exist_ok=True)
        self._robots = None
        self._robots_text = ""

    def _load_robots(self):
        if self._robots is not None:
            return
        try:
            r = self.s.get(urljoin(self.base, "/robots.txt"), timeout=self.timeout)
            if r.status_code == 200:
                self._robots_text = r.text
        except requests.RequestException:
            pass
        self._robots = RobotFileParser()
        self._robots.parse(self._robots_text.splitlines())
        self._robots.modified()  # đánh dấu đã đọc, nếu không can_fetch luôn trả False

    def allowed(self, url):
        self._load_robots()
        return self._robots.can_fetch(self.s.headers["User-Agent"], url)

    def robots_sitemaps(self):
        self._load_robots()
        return re.findall(r"(?im)^\s*sitemap:\s*(\S+)", self._robots_text)

    def get(self, url, params=None, method="GET", body=None):
        """Trả về bytes, hoặc None nếu không lấy được."""
        key = hashlib.sha1(
            json.dumps([method, url, params, body], sort_keys=True, default=str).encode()
        ).hexdigest()
        cache = self.raw_dir / key
        if cache.exists():
            return cache.read_bytes()
        if self.respect_robots and not self.allowed(url):
            log(f"  [bỏ qua] robots.txt không cho phép: {url}")
            return None
        for attempt in range(self.retries):
            try:
                r = self.s.request(method, url, params=params, json=body, timeout=self.timeout)
                if r.status_code == 404:
                    return None
                if r.status_code in (403, 429) or r.status_code >= 500:
                    wait = 5 * (attempt + 1)
                    log(f"  [{r.status_code}] {url} - chờ {wait}s rồi thử lại")
                    time.sleep(wait)
                    continue
                r.raise_for_status()
                cache.write_bytes(r.content)
                time.sleep(random.uniform(self.dmin, self.dmax))
                return r.content
            except requests.RequestException as e:
                log(f"  [lỗi] {url}: {e}")
                time.sleep(2 ** attempt)
        log(f"  [thất bại] {url}")
        return None


# --------------------------------------------------------------------------- #
# Tiện ích chuẩn hóa
# --------------------------------------------------------------------------- #
def to_num(v):
    """'27.280.000' / '27280000.00' / 27280000 -> số."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() else v
    s = re.sub(r"[^\d.,]", "", str(v))
    if not s:
        return None
    if re.fullmatch(r"\d{1,3}([.,]\d{3})+", s):      # dấu chấm/phẩy là phân cách hàng nghìn
        s = re.sub(r"[.,]", "", s)
    else:
        s = s.replace(",", "")
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return None


def get_path(obj, path):
    """Lấy giá trị theo đường dẫn 'a.b.0.c'. Không có thì trả None."""
    cur = obj
    for part in str(path).split("."):
        if isinstance(cur, dict):
            cur = cur.get(part)
        elif isinstance(cur, list) and part.isdigit() and int(part) < len(cur):
            cur = cur[int(part)]
        else:
            return None
        if cur is None:
            return None
    return cur


# --------------------------------------------------------------------------- #
# Chế độ 1: sitemap + JSON-LD
# --------------------------------------------------------------------------- #
def iter_sitemap(fetcher, url, sitemap_include=None, depth=0):
    if depth > 4:
        return
    content = fetcher.get(url)
    if not content:
        return
    if content[:2] == b"\x1f\x8b":                    # file .xml.gz
        content = gzip.decompress(content)
    try:
        root = ET.fromstring(content)
    except ET.ParseError:
        log(f"  [lỗi] sitemap không đọc được: {url}")
        return
    kind = root.tag.split("}")[-1]
    locs = [e.text.strip() for e in root.iter() if e.tag.split("}")[-1] == "loc" and e.text]
    if kind == "sitemapindex":
        for loc in locs:
            if sitemap_include and not re.search(sitemap_include, loc):
                continue
            yield from iter_sitemap(fetcher, loc, sitemap_include, depth + 1)
    else:
        yield from locs


def _flatten_ld(data):
    if isinstance(data, list):
        for x in data:
            yield from _flatten_ld(x)
    elif isinstance(data, dict):
        if "@graph" in data:
            yield from _flatten_ld(data["@graph"])
        yield data


def _has_type(node, t):
    nt = node.get("@type")
    return nt == t or (isinstance(nt, list) and t in nt)


def parse_product_page(html, url, source):
    soup = BeautifulSoup(html, "html.parser")
    product, breadcrumb = None, None
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or tag.get_text())
        except (json.JSONDecodeError, TypeError):
            continue
        for node in _flatten_ld(data):
            if _has_type(node, "Product") and product is None:
                product = node
            elif _has_type(node, "BreadcrumbList") and breadcrumb is None:
                breadcrumb = node
    if not product:
        return None

    offers = product.get("offers") or {}
    if isinstance(offers, list):
        offers = offers[0] if offers else {}
    price = to_num(offers.get("price") or offers.get("lowPrice"))
    agg = product.get("aggregateRating") or {}
    brand = product.get("brand")
    if isinstance(brand, dict):
        brand = brand.get("name")
    image = product.get("image")
    if isinstance(image, list):
        image = image[0] if image else None
    if isinstance(image, dict):
        image = image.get("url")

    category = None
    if breadcrumb:
        names = []
        for el in breadcrumb.get("itemListElement", []):
            item = el.get("item")
            names.append(el.get("name") or (item.get("name") if isinstance(item, dict) else None))
        category = " > ".join(n for n in names[:-1] if n) or None   # bỏ phần tử cuối (chính sản phẩm)

    return {
        "source": source,
        "product_id": product.get("sku") or product.get("productID") or product.get("mpn"),
        "name": product.get("name"),
        "brand": brand,
        "category": category,
        "price": price,
        "original_price": to_num(offers.get("highPrice")),
        "currency": offers.get("priceCurrency"),
        "availability": str(offers.get("availability", "")).rsplit("/", 1)[-1] or None,
        "rating": to_num(agg.get("ratingValue")),
        "review_count": to_num(agg.get("reviewCount") or agg.get("ratingCount")),
        "image": image,
        "url": url,
    }


def collect_product_urls(fetcher, cfg):
    sitemaps = cfg.get("sitemap_urls") or fetcher.robots_sitemaps()
    if not sitemaps:
        log("Không tìm thấy sitemap. Hãy điền 'sitemap_urls' trong config.")
        return []
    log(f"Sitemap: {sitemaps}")
    url_re = cfg.get("url_include")
    keywords = [k.lower() for k in cfg.get("url_keywords") or []]
    seen, urls = set(), []
    for sm in sitemaps:
        for u in iter_sitemap(fetcher, sm, cfg.get("sitemap_include")):
            if u in seen:
                continue
            seen.add(u)
            if url_re and not re.search(url_re, u):
                continue
            if keywords and not any(k in u.lower() for k in keywords):
                continue
            urls.append(u)
    return urls


def crawl_sitemap(fetcher, cfg, limit, dry_run):
    urls = collect_product_urls(fetcher, cfg)
    log(f"Tìm thấy {len(urls)} URL sản phẩm phù hợp bộ lọc.")
    if dry_run:
        for u in urls[:10]:
            log(f"  ví dụ: {u}")
        return
    if limit:
        urls = urls[:limit]
    for i, u in enumerate(urls, 1):
        content = fetcher.get(u)
        if not content:
            continue
        rec = parse_product_page(content.decode("utf-8", errors="replace"), u, cfg["name"])
        if rec is None:
            log(f"  [{i}/{len(urls)}] không có JSON-LD Product: {u}")
            continue
        log(f"  [{i}/{len(urls)}] {rec['name']} - {rec['price']}")
        yield rec


# --------------------------------------------------------------------------- #
# Chế độ 2: API JSON có phân trang
# --------------------------------------------------------------------------- #
def _subst(obj, page):
    """Thay {page} trong params/body (đệ quy)."""
    if isinstance(obj, str):
        return obj.replace("{page}", str(page))
    if isinstance(obj, dict):
        return {k: _subst(v, page) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_subst(v, page) for v in obj]
    return obj


def crawl_api(fetcher, cfg, limit, dry_run):
    api = cfg["api"]
    count = 0
    for page in range(api.get("start_page", 1), api.get("start_page", 1) + api.get("max_pages", 50)):
        content = fetcher.get(
            api["url"],
            params=_subst(api.get("params"), page),
            method=api.get("method", "GET"),
            body=_subst(api.get("body"), page),
        )
        if not content:
            break
        try:
            data = json.loads(content)
        except json.JSONDecodeError:
            log("  [lỗi] phản hồi không phải JSON (có thể bị chặn hoặc sai URL).")
            break
        items = get_path(data, api["items_path"]) if api.get("items_path") else data
        if not items:
            log(f"Trang {page} rỗng, dừng.")
            break
        log(f"Trang {page}: {len(items)} sản phẩm")
        if dry_run:
            log(f"  ví dụ 1 phần tử: {json.dumps(items[0], ensure_ascii=False)[:400]}")
            return
        for it in items:
            rec = {"source": cfg["name"]}
            for field, path in api["fields"].items():
                val = get_path(it, path)
                rec[field] = to_num(val) if field in ("price", "original_price", "rating", "review_count") else val
            if rec.get("url") and api.get("url_prefix") and not str(rec["url"]).startswith("http"):
                rec["url"] = api["url_prefix"] + str(rec["url"])
            yield rec
            count += 1
            if limit and count >= limit:
                return


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="Công cụ crawl sản phẩm dùng chung")
    ap.add_argument("config")
    ap.add_argument("--limit", type=int, default=None, help="số sản phẩm tối đa")
    ap.add_argument("--dry-run", action="store_true", help="chỉ thăm dò, không lưu dữ liệu")
    args = ap.parse_args()

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    fetcher = Fetcher(cfg)
    runner = {"sitemap_jsonld": crawl_sitemap, "api": crawl_api}[cfg["mode"]]

    out_dir = Path(cfg.get("output_dir", "data"))
    out_dir.mkdir(parents=True, exist_ok=True)
    now = datetime.now(timezone.utc)
    out_file = out_dir / f"{cfg['name']}_{now.strftime('%Y%m%d_%H%M')}.csv"

    n = 0
    f = None
    try:
        for rec in runner(fetcher, cfg, args.limit, args.dry_run) or []:
            if f is None:
                f = open(out_file, "w", newline="", encoding="utf-8-sig")
                writer = csv.DictWriter(f, fieldnames=FIELDS, extrasaction="ignore")
                writer.writeheader()
            rec["crawled_at"] = now.isoformat(timespec="seconds")
            writer.writerow(rec)
            f.flush()                                   # ngắt giữa chừng vẫn giữ được dữ liệu
            n += 1
    except KeyboardInterrupt:
        log("Đã dừng theo yêu cầu.")
    finally:
        if f:
            f.close()
    if not args.dry_run:
        log(f"Xong: {n} dòng -> {out_file if n else '(không có dữ liệu)'}")


if __name__ == "__main__":
    main()
