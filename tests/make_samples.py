#!/usr/bin/env python3
"""Generate sample PDFs used to test the converter and merger.

Produces three files in tests/samples:
  * report_tables.pdf – headers/footers, paragraphs, and a table split
    across two pages (tests table extraction + page-span continuation).
  * notes_text.pdf    – pure text with several paragraphs and page breaks.
  * scan_image.pdf    – an embedded image, a paragraph, and a small table.
"""

import os

import pymupdf as fitz
from PIL import Image, ImageDraw

SAMPLES_DIR = os.path.join(os.path.dirname(__file__), "samples")


def draw_table(page, x0, y0, x1, y1, data, header_rows=1):
    rows = len(data)
    cols = max(len(r) for r in data)
    step_x = (x1 - x0) / cols
    step_y = (y1 - y0) / rows
    for i in range(rows + 1):
        y = y0 + i * step_y
        page.draw_line((x0, y), (x1, y), color=(0.1, 0.1, 0.1), width=0.7)
    for j in range(cols + 1):
        x = x0 + j * step_x
        page.draw_line((x, y0), (x, y1), color=(0.1, 0.1, 0.1), width=0.7)
    for r, row in enumerate(data):
        for c, value in enumerate(row):
            x = x0 + c * step_x + 4
            y = y0 + r * step_y + step_y * 0.35
            page.insert_text((x, y), str(value), fontsize=8)
    return (rows, cols)


def build_report_tables(path):
    doc = fitz.open()

    # -------- page 1 ------------------------------------------------------
    page = doc.new_page()
    page.insert_text((72, 54), "Quarterly Sales Report", fontsize=18)
    page.insert_text((72, 74), "Prepared on August 14, 2026 by Finance.",
                     fontsize=10)
    page.insert_text((72, 104),
                     "This report summarises product sales for Q2. "
                     "All figures are in USD and unaudited.",
                     fontsize=10)

    header = ["Item", "Amount", "Notes"]
    rows1 = [
        ["Widget A", "1200", "Nationwide"],
        ["Widget B", "0850", "Led by North region"],
        ["Gadget C", "3400.50", "Includes taxes"],
        ["Gadget D", "75", "Bundle only"],
        ["Tool E", "999", "Back-ordered"],
        ["Tool F", "2600", "Discount applied"],
        ["Part G", "12", "Per-unit price"],
        ["Part H", "180", "Discontinued after Q2"],
        ["Widget A", "900", "Repeat order"],
        ["Gadget I", "4500", "New customer"],
        ["Tool J", "340", "Pending approval"],
        ["Part K", "20", "Sample units"],
        ["Gadget L", "3100", "Regional promo"],
    ]
    draw_table(page, 72, 130, 520, 790, [header] + rows1)

    page.insert_text((300, 812), "Page 1", fontsize=9)  # footer

    # -------- page 2 (continuation of the table) ----------------------------
    page = doc.new_page()
    page.insert_text((72, 54), "Quarterly Sales Report (continued)",
                     fontsize=16)
    rows2 = [
        ["Gadget M", "2100", "Q3 pre-order"],
        ["Tool N", "45", "Bulk discount"],
        ["Widget O", "1300", "Restock shipped"],
        ["Part P", "8", "Per-crate price"],
        ["Gadget Q", "5600", "Top seller"],
    ]
    draw_table(page, 72, 70, 520, 200, [header] + rows2)

    page.insert_text((72, 250),
                     "Total units shipped across all items grew 14% "
                     "compared with the previous quarter.",
                     fontsize=10)
    page.insert_text((300, 782), "Page 2", fontsize=9)

    doc.save(path)
    doc.close()


def build_notes_text(path):
    doc = fitz.open()
    page = doc.new_page()
    paragraphs = [
        ("Meeting Notes — Product Sync", 16),
        ("August 14, 2026", 10),
        ("Attendees: Alice (Product), Bob (Eng), Carol (Design).", 10),
        ("We reviewed the roadmap for the next release. Three themes were "
         "prioritised: reliability of the sync engine, faster cold start, "
         "and offline-first editing. Bob proposed splitting the sync work "
         "into two milestones and running a canary rollout.", 10),
        ("Action items:", 10),
        ("1. Alice to share the analytics dashboard by Friday.",
         10),
        ("2. Bob to spike the cold-start optimisation and report back.",
         10),
        ("3. Carol to prepare mock-ups for the offline indicator.",
         10),
    ]
    y = 80
    for text, size in paragraphs:
        page.insert_text((72, y), text, fontsize=size)
        y += size + 14

    # second page
    page = doc.new_page()
    page.insert_text((72, 80), "Follow-up", fontsize=16)
    page.insert_text((72, 110),
                     "Carol shared two concepts for the offline banner. "
                     "Team preferred the minimal variant. Next sync review "
                     "is scheduled for next Tuesday.",
                     fontsize=10)
    page.insert_text((300, 782), "Page 2", fontsize=9)

    doc.save(path)
    doc.close()


def build_scan_image(path):
    # Build a source PNG with Pillow
    png = os.path.join(SAMPLES_DIR, "scan_source.png")
    img = Image.new("RGB", (640, 360), (245, 243, 238))
    d = ImageDraw.Draw(img)
    d.rectangle([40, 40, 600, 320], outline=(60, 60, 60), width=3)
    d.text((80, 100), "OFFICIAL STAMP - APPROVED", fill=(180, 40, 40))
    d.ellipse([420, 60, 580, 220], outline=(180, 40, 40), width=4)
    d.text((430, 130), "SEAL", fill=(180, 40, 40))
    img.save(png)

    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 54), "Scanned Document", fontsize=18)
    page.insert_text((72, 78),
                     "The image below is embedded as a picture so it must "
                     "be preserved in the Excel output.",
                     fontsize=10)
    page.insert_image(fitz.Rect(72, 100, 72 + 320, 100 + 180), filename=png)

    page.insert_text((72, 320),
                     "The table below records the scanned figures:",
                     fontsize=10)
    data = [["ID", "Value"], ["A-01", "1250"], ["A-02", "770"], ["A-03", "2200"]]
    draw_table(page, 72, 340, 300, 430, data, header_rows=1)
    page.insert_text((300, 782), "Page 1", fontsize=9)

    doc.save(path)
    doc.close()


def build_encrypted(path):
    """A password-protected PDF (user password: secret)."""
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 80), "Confidential Monthly Summary", fontsize=16)
    page.insert_text((72, 110),
                     "This document requires a password to open.",
                     fontsize=10)
    page.insert_text((72, 140), "Figures: 812, 3310, 44.", fontsize=10)
    doc.save(path, encryption=fitz.PDF_ENCRYPT_AES_256,
             user_pw="secret", owner_pw="secret")
    doc.close()


def build_catalog(path):
    """A product table with the exact extraction columns."""
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 54), "Product Catalogue", fontsize=18)
    page.insert_text((72, 74),
                     "Extract Material Code, Item Description, EAN No., "
                     "Quantity and Unit Base Cost from this table.",
                     fontsize=10)
    header = ["Material Code", "Item Description", "EAN No.",
              "Quantity", "Unit Base Cost"]
    rows = [
        ["MC-1001", "Aluminium Screw 3mm", "4002355021235", "1500", "0.02"],
        ["MC-1002", "Steel Washer 8mm", "4002355021242", "3200", "0.01"],
        ["MC-1003", "Nylon Bushing", "4002355021259", "85", "0.45"],
        ["MC-1004", "Rubber Seal", "4002355021266", "640", "0.30"],
    ]
    draw_table(page, 72, 100, 520, 240, [header] + rows)
    page.insert_text((300, 782), "Page 1", fontsize=9)
    doc.save(path)
    doc.close()


def main():
    os.makedirs(SAMPLES_DIR, exist_ok=True)
    build_report_tables(os.path.join(SAMPLES_DIR, "report_tables.pdf"))
    build_notes_text(os.path.join(SAMPLES_DIR, "notes_text.pdf"))
    build_scan_image(os.path.join(SAMPLES_DIR, "scan_image.pdf"))
    build_encrypted(os.path.join(SAMPLES_DIR, "confidential.pdf"))
    build_catalog(os.path.join(SAMPLES_DIR, "catalog.pdf"))
    print("Sample PDFs written to", SAMPLES_DIR)


if __name__ == "__main__":
    main()
