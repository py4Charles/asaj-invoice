from __future__ import annotations

import argparse
import json
import os
import random
import re
import subprocess
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Sequence

from openpyxl import load_workbook

PROJECT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = PROJECT_DIR / "assets" / "invoice-ORD-1772740748892.html"
DEFAULT_TABLE = PROJECT_DIR / "output" / "table_updated.xlsx"
DEFAULT_OUT = PROJECT_DIR / "invoices"

NOTES = (
    "All goods inspected and certified for international export. "
    "Certificate of origin available on request."
)

PRODUCTS: Dict[str, dict] = {
    "Cashew nuts": {
        "description": "Premium Raw Cashew Nuts – Grade W240",
        "spec": "Moisture ≤5% · Fresh & Natural · Export Quality",
        "unitPrice": 159.00,
    },
    "Ginger": {
        "description": "Dried Ginger Root (High Oil Content)",
        "spec": "Strong Aroma · Export Quality · Well Dried",
        "unitPrice": 326.00,
    },
    "Soybean": {
        "description": "Non-GMO Soybeans – Export Grade",
        "spec": "High Protein · Well Cleaned · Non-GMO Certified",
        "unitPrice": 340.00,
    },
    "Moringa": {
        "description": "Moringa Leaf Powder – Certified Organic",
        "spec": "Certified Organic · Rich in Nutrients · Export Quality",
        "unitPrice": 23.80,
    },
    "Mango": {
        "description": "Fresh Mango – Export Grade",
        "spec": "Fresh & Natural · Export Quality",
        "unitPrice": 160.00,
    },
}

FIRST_NAMES = [
    "Kofi", "Kwame", "Yaw", "Kojo", "Kwabena", "Kwesi", "Kwaku", "Nana",
    "Ama", "Akosua", "Abena", "Adwoa", "Efua", "Esi", "Afia", "Adjoa",
    "Yaa", "Akua", "Araba", "Efua", "Naa", "Ewurama", "Maame", "Serwaa",
]

LAST_NAMES = [
    "Oteng", "Mensah", "Boateng", "Owusu", "Asante", "Appiah", "Osei",
    "Addo", "Amoah", "Danso", "Frimpong", "Nyarko", "Agyeman", "Acheampong",
    "Oppong", "Sarpong", "Bedford", "Quaye", "Tetteh", "Lartey", "Annan",
    "Antwi", "Boadi", "Ofori",
]

PHONE_PREFIXES = ["020", "024", "059", "053", "054", "025"]
STATUSES = ["Delivered", "Processing", "Pending", "Cancelled"]
STATUS_WEIGHTS = [60, 20, 15, 5]
DATE_FORMAT = "%d-%b-%Y"
INVOICE_PATTERN = re.compile(r"const INVOICE = \{.*?\n    \};", re.DOTALL)


def load_weights(table_path: Path) -> List[float]:
    workbook = load_workbook(table_path, read_only=True, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    headers = list(rows[0])
    index = headers.index("weight_g")

    weights: List[float] = []
    for row in rows[1:]:
        value = row[index]
        if isinstance(value, (int, float)) and value > 0:
            weights.append(float(value))
    if not weights:
        raise SystemExit("No usable weight_g values found in the table.")
    return weights


def make_customer(rng: random.Random, used_names: set[str]) -> dict:
    for _ in range(10000):
        first = rng.choice(FIRST_NAMES)
        last = rng.choice(LAST_NAMES)
        full_name = f"{first} {last}"
        if full_name not in used_names:
            used_names.add(full_name)
            break
    else:
        raise SystemExit("Could not generate enough unique customer names.")

    email = f"{first}{last}{rng.randint(1, 99)}@gmail.com".lower()
    phone = rng.choice(PHONE_PREFIXES) + "".join(str(rng.randint(0, 9)) for _ in range(7))

    return {"name": full_name, "email": email, "phone": phone}


def make_reference(rng: random.Random, used_refs: set[str]) -> str:
    for _ in range(10000):
        digits = str(rng.randint(10**12, 10**13 - 1))
        if digits not in used_refs:
            used_refs.add(digits)
            return digits
    raise SystemExit("Could not generate enough unique references.")


def make_items(
    rng: random.Random,
    products: Sequence[str],
    total_weight_kg: float,
) -> List[dict]:
    per_product_weight = round(total_weight_kg / len(products), 3)
    items: List[dict] = []
    for name in products:
        product = PRODUCTS[name]
        qty = round(rng.uniform(0.2, 0.9), 2)
        unit_price = product["unitPrice"]
        items.append(
            {
                "description": product["description"],
                "qty": qty,
                "unit": "kg",
                "unitPrice": unit_price,
                "lineTotal": round(qty * unit_price, 2),
                "allocatedWeightKg": per_product_weight,
            }
        )
    return items


def build_invoice(
    rng: random.Random,
    weights: Sequence[float],
    used_names: set[str],
    used_refs: set[str],
) -> dict:
    products = rng.sample(list(PRODUCTS), rng.randint(2, 4))
    weight_g = rng.choice(list(weights))
    total_weight_kg = weight_g / 1000

    order_date = date(2026, 1, 1) + timedelta(days=rng.randint(0, (date(2026, 3, 3) - date(2026, 1, 1)).days))
    dispatch_date = order_date + timedelta(days=5)
    delivery_date = order_date + timedelta(days=12)

    reference = make_reference(rng, used_refs)
    customer = make_customer(rng, used_names)
    items = make_items(rng, products, total_weight_kg)

    return {
        "company": {"name": "ASAJ Company Limited", "subtitle": "ASAJ Commercial Invoice"},
        "invoiceNumber": f"INV-ORD-{reference}",
        "transaction": {
            "reference": f"ORD-{reference}",
            "date": order_date.strftime(DATE_FORMAT),
            "dispatchDate": dispatch_date.strftime(DATE_FORMAT),
            "deliveryDate": delivery_date.strftime(DATE_FORMAT),
        },
        "customer": customer,
        "delivery": {
            "status": rng.choices(STATUSES, weights=STATUS_WEIGHTS, k=1)[0],
            "type": "Export shipment",
        },
        "currency": "GHS",
        "items": items,
        "notes": NOTES,
        "_source": {"weight_g": weight_g, "totalWeightKg": round(total_weight_kg, 3)},
        "_products": products,
    }


def render_html(template: str, invoice: dict) -> str:
    payload = json.dumps(invoice, ensure_ascii=False, indent=2)
    if not INVOICE_PATTERN.search(template):
        raise SystemExit("Could not locate the INVOICE object in the template.")
    return INVOICE_PATTERN.sub(lambda _: f"const INVOICE = {payload};", template, count=1)


def find_browser() -> Path:
    candidates = [
        Path(os.environ.get("ProgramFiles", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("ProgramFiles(x86)", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("ProgramFiles", "")) / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("ProgramFiles(x86)", "")) / "Microsoft/Edge/Application/msedge.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise SystemExit("Google Chrome or Microsoft Edge is required to create PDFs.")


def html_to_pdf(browser: Path, html_path: Path, pdf_path: Path, profile: Path) -> None:
    command = [
        str(browser),
        "--headless=new",
        "--disable-gpu",
        "--no-pdf-header-footer",
        "--run-all-compositor-stages-before-draw",
        "--virtual-time-budget=2000",
        f"--user-data-dir={profile}",
        f"--print-to-pdf={pdf_path.resolve()}",
        html_path.resolve().as_uri(),
    ]
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if not pdf_path.exists() or pdf_path.stat().st_size == 0:
        raise RuntimeError(
            "PDF generation failed.\n"
            f"stdout: {completed.stdout.strip()}\n"
            f"stderr: {completed.stderr.strip()}"
        )


def parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate ASAJ invoice PDFs.")
    parser.add_argument("--table", type=Path, default=DEFAULT_TABLE)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT)
    parser.add_argument("--count", type=int, default=57)
    parser.add_argument("--seed", type=int, default=20261013)
    parser.add_argument("--keep-html", action="store_true", help="Keep generated HTML files.")
    parser.add_argument("--no-pdf", action="store_true", help="Only write HTML and manifest.")
    return parser.parse_args(argv)


def main(argv: Iterable[str] | None = None) -> int:
    args = parse_args(argv)

    if not args.template.exists():
        raise SystemExit(f"Template not found: {args.template}")
    if not args.table.exists():
        raise SystemExit(f"Table not found: {args.table}")

    template = args.template.read_text(encoding="utf-8")
    weights = load_weights(args.table)

    rng = random.Random(args.seed)
    used_names: set[str] = set()
    used_refs: set[str] = set()

    invoices = [
        build_invoice(rng, weights, used_names, used_refs) for _ in range(args.count)
    ]

    args.out_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = args.out_dir / "manifest.json"
    manifest_path.write_text(
        json.dumps(invoices, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    html_dir = args.out_dir / "html"
    html_dir.mkdir(exist_ok=True)

    browser = None if args.no_pdf else find_browser()
    profile = Path(tempfile.mkdtemp(prefix="asaj-chrome-"))

    for number, invoice in enumerate(invoices, start=1):
        reference = invoice["transaction"]["reference"].replace("ORD-", "")
        html = render_html(template, invoice)
        html_path = html_dir / f"invoice-ORD-{reference}.html"
        html_path.write_text(html, encoding="utf-8")

        if not args.no_pdf:
            pdf_path = args.out_dir / f"invoice-ORD-{reference}.pdf"
            html_to_pdf(browser, html_path, pdf_path, profile)
            if number % 10 == 0 or number == len(invoices):
                print(f"Generated {number}/{len(invoices)} PDFs")

    if not args.keep_html:
        for html_file in html_dir.glob("*.html"):
            html_file.unlink()
        html_dir.rmdir()

    print(f"Created {len(invoices)} invoices in {args.out_dir}")
    print(f"Manifest: {manifest_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
