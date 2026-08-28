#!/usr/bin/env python3
"""
generate_blog.py — Warung Piksel blog generator

Saran ukuran gambar biar nggak berat: cover ~1200×675px, sisipan ~1000×700px, format JPG kualitas 80-85%.

Cara pakai:
    1. Edit articles.xlsx (sheet "Artikel") — tambah/ubah baris artikel.
    2. Jalankan: python generate_blog.py
    3. Hasilnya muncul di folder output/ — upload semua isinya ke hosting.

Butuh: pip install openpyxl
"""

import os
import re
import random
import html
import shutil
from datetime import datetime
import openpyxl

# ---------- Konfigurasi ----------
EXCEL_FILE = "articles.xlsx"
SHEET_NAME = "Artikel"
TEMPLATE_INDEX = "template_index.html"
TEMPLATE_ARTIKEL = "template_artikel.html"
OUTPUT_DIR = "../blog"

DATE_FORMATS = ["%d %b %Y", "%d/%m/%Y", "%Y-%m-%d"]
BULAN_ID = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "Mei", 6: "Jun",
            7: "Jul", 8: "Agu", 9: "Sep", 10: "Okt", 11: "Nov", 12: "Des"}


def parse_tanggal(raw):
    """Coba parse tanggal ke object datetime buat pengurutan. Kalau gagal, kembalikan None."""
    raw = str(raw).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def esc(text):
    """Escape karakter HTML biar aman, tapi biarkan text kosong tetap kosong."""
    return html.escape(str(text or ""), quote=False)


def build_inline_image(directive):
    """directive = teks setelah 'gambar:' , format: nama-file.jpg atau nama-file.jpg | Keterangan"""
    if "|" in directive:
        filename, caption = directive.split("|", 1)
        filename, caption = filename.strip(), caption.strip()
    else:
        filename, caption = directive.strip(), ""
    fig = (
        '      <figure class="inline-img">\n'
        f'        <img src="images/{esc(filename)}" alt="{esc(caption or filename)}" loading="lazy">\n'
    )
    if caption:
        fig += f'        <figcaption>{esc(caption)}</figcaption>\n'
    fig += "      </figure>"
    return fig, filename


def isi_to_html(isi_text, image_refs):
    """
    Ubah isi artikel (plain text) jadi HTML.
    - Baris kosong = pemisah paragraf baru
    - Baris diawali '## ' = jadi <h2>
    - Baris diawali 'gambar:' = jadi gambar sisipan, format:
      gambar: nama-file.jpg   ATAU   gambar: nama-file.jpg | Keterangan gambar
    image_refs: list yang akan diisi nama file gambar yang dipakai (buat validasi ada/tidaknya file)
    """
    if not isi_text:
        return ""
    blocks = re.split(r"\n\s*\n", str(isi_text).strip())
    out = []
    for block in blocks:
        lines = [l for l in block.split("\n") if l.strip()]
        if not lines:
            continue
        first = lines[0].strip()
        if first.lower().startswith("gambar:"):
            fig_html, filename = build_inline_image(first[len("gambar:"):].strip())
            out.append(fig_html)
            image_refs.append(filename)
            sisa = lines[1:]
            if sisa:
                paragraf = " ".join(l.strip() for l in sisa)
                out.append(f"      <p>{esc(paragraf)}</p>")
        elif first.startswith("## "):
            out.append(f"      <h2>{esc(first[3:].strip())}</h2>")
            sisa = lines[1:]
            if sisa:
                paragraf = " ".join(l.strip() for l in sisa)
                out.append(f"      <p>{esc(paragraf)}</p>")
        else:
            paragraf = " ".join(l.strip() for l in lines)
            out.append(f"      <p>{esc(paragraf)}</p>")
    return "\n".join(out)


def build_cover_html(article):
    fn = article.get("gambar_cover")
    if not fn or not str(fn).strip():
        return ""
    fn = str(fn).strip()
    return f'    <img class="cover-img" src="images/{esc(fn)}" alt="{esc(article["judul"])}" loading="lazy">'


def build_thumb_html(article):
    fn = article.get("gambar_cover")
    if not fn or not str(fn).strip():
        return ""
    fn = str(fn).strip()
    return f'      <img class="thumb" src="images/{esc(fn)}" alt="{esc(article["judul"])}" loading="lazy">\n'


def hitung_read_time(isi_text):
    kata = len(str(isi_text or "").split())
    menit = max(1, round(kata / 200))
    return f"{menit} menit baca"


def generate_hash():
    return "".join(random.choice("0123456789abcdef") for _ in range(6))


def build_pullquote_block(quote):
    if not quote or not str(quote).strip():
        return ""
    lines = [l for l in str(quote).split("\n") if l.strip()]
    n = max(1, len(lines))
    gutter = "<br>".join(f"{i+1:02d}" for i in range(n))
    quote_html = " ".join(l.strip() for l in lines) if lines else esc(quote)
    return (
        '      <div class="pull-quote">\n'
        f'        <div class="gutter">{gutter}</div>\n'
        f'        <blockquote>{esc(quote_html)}</blockquote>\n'
        '      </div>'
    )


def build_diffnote_block(minus, plus):
    if not (minus and str(minus).strip()) and not (plus and str(plus).strip()):
        return ""
    parts = []
    if minus and str(minus).strip():
        parts.append(f'        <span class="minus">- sebelum:</span> {esc(minus)}<br>')
    if plus and str(plus).strip():
        parts.append(f'        <span class="plus">+ sesudah:</span> {esc(plus)}')
    return '      <div class="diff-note">\n' + "\n".join(parts) + "\n      </div>"


def build_related_html(article, by_slug, all_articles):
    slugs = []
    if article.get("related_slugs"):
        slugs = [s.strip() for s in str(article["related_slugs"]).split(",") if s.strip()]

    related = []
    for s in slugs:
        if s in by_slug and s != article["slug"]:
            related.append(by_slug[s])

    if not related:
        # fallback: artikel lain dengan kategori sama, terbaru dulu, maksimal 3
        same_cat = [a for a in all_articles if a["kategori"] == article["kategori"] and a["slug"] != article["slug"]]
        related = same_cat[:3]

    related = related[:3]
    if not related:
        return '    <p style="font-family:var(--mono);font-size:0.85rem;color:var(--ink-soft);">Belum ada artikel terkait.</p>'

    items = []
    for r in related:
        items.append(
            '    <a class="related-item" href="artikel-{slug}.html">\n'
            '      <span class="hash-small">#{hash}</span>\n'
            '      <h4>{judul}</h4>\n'
            '    </a>'.format(slug=r["slug"], hash=r["hash"], judul=esc(r["judul"]))
        )
    return "\n".join(items)


def build_filter_chips(categories):
    chips = ['    <button class="chip all active" data-filter="all">semua</button>']
    for cat in categories:
        chips.append(f'    <button class="chip" data-filter="{esc(cat)}">{esc(cat)}</button>')
    return "\n".join(chips)


def build_post_card(article):
    return (
        f'    <a class="post-card" href="artikel-{article["slug"]}.html" data-category="{esc(article["kategori"])}">\n'
        f'      <div class="commit">\n'
        f'{build_thumb_html(article)}'
        f'        <span class="hash">#{article["hash"]}</span>\n'
        f'        <span class="date">{esc(article["tanggal"])}</span>\n'
        f'      </div>\n'
        f'      <div class="post-body">\n'
        f'        <h2>{esc(article["judul"])}</h2>\n'
        f'        <p class="post-excerpt">{esc(article["excerpt"])}</p>\n'
        f'        <div class="meta-row">\n'
        f'          <span class="tag-pill">{esc(article["kategori"])}</span>\n'
        f'          <span class="read-time">{esc(article["read_time"])}</span>\n'
        f'        </div>\n'
        f'      </div>\n'
        f'    </a>'
    )


def load_articles():
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[SHEET_NAME]
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() for h in rows[0]]
    articles = []
    for row in rows[1:]:
        if not any(row):
            continue
        data = dict(zip(headers, row))
        if not data.get("slug") or not str(data.get("slug")).strip():
            continue
        data["slug"] = str(data["slug"]).strip()
        data["kategori"] = str(data.get("kategori") or "umum").strip()
        data["judul"] = str(data.get("judul") or "").strip()
        data["tanggal"] = str(data.get("tanggal") or "").strip()
        if not data.get("hash") or not str(data.get("hash")).strip():
            data["hash"] = generate_hash()
        else:
            data["hash"] = str(data["hash"]).strip()
        if not data.get("read_time") or not str(data.get("read_time")).strip():
            data["read_time"] = hitung_read_time(data.get("isi"))
        articles.append(data)

    # urutkan dari yang tanggalnya paling baru
    def sort_key(a):
        d = parse_tanggal(a["tanggal"])
        return d or datetime.min
    articles.sort(key=sort_key, reverse=True)
    return articles


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    articles = load_articles()
    if not articles:
        print("Tidak ada artikel valid ditemukan di articles.xlsx. Cek lagi kolom 'slug'.")
        return

    by_slug = {a["slug"]: a for a in articles}
    categories = sorted(set(a["kategori"] for a in articles))

    with open(TEMPLATE_ARTIKEL, "r", encoding="utf-8") as f:
        template_artikel = f.read()
    with open(TEMPLATE_INDEX, "r", encoding="utf-8") as f:
        template_index = f.read()

    semua_gambar_dipakai = []  # buat validasi file ada/tidak di akhir

    # ---- generate tiap halaman artikel ----
    for article in articles:
        image_refs = []
        isi_html = isi_to_html(article.get("isi"), image_refs)
        semua_gambar_dipakai.extend(image_refs)
        if article.get("gambar_cover") and str(article["gambar_cover"]).strip():
            semua_gambar_dipakai.append(str(article["gambar_cover"]).strip())

        page = template_artikel
        page = page.replace("{{JUDUL}}", esc(article["judul"]))
        page = page.replace("{{META_DESC}}", esc(article.get("meta_deskripsi", "")))
        page = page.replace("{{KATEGORI}}", esc(article["kategori"]))
        page = page.replace("{{HASH}}", article["hash"])
        page = page.replace("{{TANGGAL}}", esc(article["tanggal"]))
        page = page.replace("{{READ_TIME}}", esc(article["read_time"]))
        page = page.replace("{{COVER_HTML}}", build_cover_html(article))
        page = page.replace("{{ISI_HTML}}", isi_html)
        page = page.replace("{{PULLQUOTE_BLOCK}}", build_pullquote_block(article.get("pull_quote")))
        page = page.replace("{{DIFFNOTE_BLOCK}}", build_diffnote_block(article.get("diff_minus"), article.get("diff_plus")))
        page = page.replace("{{RELATED_HTML}}", build_related_html(article, by_slug, articles))

        out_path = os.path.join(OUTPUT_DIR, f"artikel-{article['slug']}.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(page)
        print(f"  -> {out_path}")

    # ---- generate index ----
    filters_html = build_filter_chips(categories)
    cards_html = "\n\n".join(build_post_card(a) for a in articles)
    index_page = template_index
    index_page = index_page.replace("{{FILTERS_HTML}}", filters_html)
    index_page = index_page.replace("{{POST_CARDS_HTML}}", cards_html)

    index_path = os.path.join(OUTPUT_DIR, "index.html")
    with open(index_path, "w", encoding="utf-8") as f:
        f.write(index_page)
    print(f"  -> {index_path}")

    # ---- salin folder images/ ke output/images/ ----
    if os.path.isdir("images"):
        shutil.copytree("images", os.path.join(OUTPUT_DIR, "images"), dirs_exist_ok=True)
        print("  -> disalin folder images/ ke output/images/")
    else:
        if semua_gambar_dipakai:
            print("\nPERINGATAN: ada artikel yang pakai gambar, tapi folder 'images/' belum ada.")
            print("  Buat folder 'images' di tempat yang sama dengan generate_blog.py, isi dengan file gambarnya.")

    # ---- cek file gambar yang direferensi tapi belum ada ----
    if semua_gambar_dipakai:
        hilang = []
        for fn in set(semua_gambar_dipakai):
            if not os.path.isfile(os.path.join("images", fn)):
                hilang.append(fn)
        if hilang:
            print("\nPERINGATAN: file gambar berikut disebut di Excel tapi belum ditemukan di folder images/:")
            for fn in sorted(hilang):
                print(f"  - {fn}")

    print(f"\nSelesai! {len(articles)} artikel ter-generate ke folder '{OUTPUT_DIR}/'.")
    print("Upload semua isi folder itu ke hosting kamu.")


if __name__ == "__main__":
    main()
