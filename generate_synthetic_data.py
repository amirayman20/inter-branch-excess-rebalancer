import pandas as pd
import numpy as np
import os

OUTPUT_DIR = "synthetic_data"
SEED = 42
N_ITEMS_PER_BRANCH = 8800
N_WAREHOUSE_ITEMS = 5100

BRANCH_NAMES = [
    "101 فرع أ",
    "102 فرع ب",
    "103 فرع ج",
    "104 فرع د",
    "105 فرع هـ",
]

WAREHOUSE_NAME = "900 المستودع المركزي"

CATEGORIES = [
    "منتجات العناية بالبشرة",
    "منتجات العناية بالشعر",
    "منتجات العناية بالفم",
    "مستلزمات العناية الشخصية",
    "عطورات ومعطرات",
    "مستحضرات تجميل متنوعة",
    "مستلزمات طبية متنوعة",
    "منتجات العناية بالطفل",
]

PRODUCT_ADJECTIVES = [
    "كلاسيك", "سينسيتيف", "ناتشورال", "اكسترا", "بيور", "فريش",
    "اكتيف", "ادفانسد", "سوفت", "انتنس", "مويستشرايزينج", "ريفريشينج"
]

PRODUCT_TYPES = [
    "غسول", "كريم", "لوشن", "جل", "بخاخ", "شامبو", "بلسم", "صابون",
    "زيت", "مرطب", "سيروم", "معجون", "مزيل عرق", "مناديل", "بودرة"
]

PRODUCT_SIZES = ["50مل", "100مل", "150مل", "200مل", "250مل", "400مل", "500مل"]


def make_item_pool(n_items, rng):
    item_codes = rng.choice(np.arange(10_000_000, 99_999_999), size=n_items, replace=False)
    items = {}
    for code in item_codes:
        adj = rng.choice(PRODUCT_ADJECTIVES)
        ptype = rng.choice(PRODUCT_TYPES)
        size = rng.choice(PRODUCT_SIZES)
        name = f"{ptype} {adj} {size}"
        category = rng.choice(CATEGORIES)
        items[code] = {"ITEM_NAME": name, "GROUP_NAME": category}
    return items


def random_expire_dates(n, rng):
    months = rng.integers(1, 13, size=n)
    years = rng.integers(26, 30, size=n)
    return [f"01/{m:02d}/{y:02d}" for m, y in zip(months, years)]


def make_branch_file(branch_name, item_pool, rng):
    codes = list(item_pool.keys())
    chosen = rng.choice(codes, size=N_ITEMS_PER_BRANCH, replace=True)
    sales_qty = rng.gamma(shape=1.2, scale=12, size=N_ITEMS_PER_BRANCH).round(0)
    balance = rng.gamma(shape=1.5, scale=8, size=N_ITEMS_PER_BRANCH).round(0)

    df = pd.DataFrame({
        "ITEM_CODE": chosen,
        "ITEM_NAME": [item_pool[c]["ITEM_NAME"] for c in chosen],
        "EXPIRE_DATE": random_expire_dates(N_ITEMS_PER_BRANCH, rng),
        "SALES_QTY": sales_qty,
        "STORE_NAME": branch_name,
        "BALANCE": balance,
    })
    return df


def make_warehouse_file(item_pool, rng):
    codes = list(item_pool.keys())
    chosen = rng.choice(codes, size=N_WAREHOUSE_ITEMS, replace=True)
    st_bal = rng.lognormal(mean=4.0, sigma=1.5, size=N_WAREHOUSE_ITEMS).round(0)
    st_bal = np.clip(st_bal, 1, None)

    df = pd.DataFrame({
        "ST_BAL": st_bal.astype(int),
        "ITEM_CODE": chosen,
        "ITEM_NAME": [item_pool[c]["ITEM_NAME"] for c in chosen],
        "EXPIRE_DATE": np.nan,
        "GROUP_NAME": [item_pool[c]["GROUP_NAME"] for c in chosen],
        "STORE_NAME": WAREHOUSE_NAME,
    })
    return df


def main():
    rng = np.random.default_rng(SEED)

    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    item_pool = make_item_pool(2500, rng)

    for i, branch_name in enumerate(BRANCH_NAMES, start=1):
        df_branch = make_branch_file(branch_name, item_pool, rng)
        path = os.path.join(OUTPUT_DIR, f"branch_{i}.xlsx")
        df_branch.to_excel(path, index=False)

    df_warehouse = make_warehouse_file(item_pool, rng)
    wh_path = os.path.join(OUTPUT_DIR, "warehouse.xlsx")
    df_warehouse.to_excel(wh_path, index=False)


if __name__ == "__main__":
    main()
