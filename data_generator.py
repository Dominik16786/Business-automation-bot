import random
import pandas as pd
from datetime import datetime, timedelta



def _random_nip() -> str:
    return "".join([str(random.randint(0, 9)) for _ in range(10)])


def _random_date(start: datetime, end: datetime) -> datetime:
    delta = end - start
    return start + timedelta(days=random.randint(0, delta.days))


def _fake_company_name() -> str:
    """Produce a realistic-looking Polish company name."""
    prefixes = ["ABC", "Euro", "Pol", "Trans", "Mega", "Pro", "Global", "Inter",
                "Fast", "Smart", "Tech", "Nova", "Alpha", "Prime", "Alfa"]
    suffixes = ["Systems", "Solutions", "Trade", "Group", "Logistics", "Services",
                "Partners", "Commerce", "Industries", "Consulting", "Labs", "Works"]
    legal   = ["Sp. z o.o.", "S.A.", "Sp.j.", "S.K.A.", "Sp. k."]
    return f"{random.choice(prefixes)}{random.choice(suffixes)} {random.choice(legal)}"



def generate_transactions(n_nips: int = 40, n_rows: int = 150) -> pd.DataFrame:
    nip_pool = [_random_nip() for _ in range(n_nips)]

    start_date = datetime(2024, 1, 1)
    end_date   = datetime(2024, 12, 31)

    rows = []
    for _ in range(n_rows):
        rows.append({
            "nip":               random.choice(nip_pool),
            "price":             round(random.uniform(100.0, 50_000.0), 2),
            "date":              _random_date(start_date, end_date),
            "random_noise_column": random.choice(["A", "B", "C", "X", "Y"]),
        })

    for _ in range(10):
        bad_nip = "".join([str(random.randint(0, 9)) for _ in range(random.choice([7, 11, 5]))])
        rows.append({
            "nip":               bad_nip,
            "price":             round(random.uniform(100.0, 5_000.0), 2),
            "date":              _random_date(start_date, end_date),
            "random_noise_column": "INVALID",
        })

    random.shuffle(rows)
    return pd.DataFrame(rows)


def generate_companies(transaction_nips: list[str], extra: int = 15) -> pd.DataFrame:

    matched_nips = random.sample(transaction_nips,
                                 k=int(len(transaction_nips) * 0.7))

    extra_nips = [_random_nip() for _ in range(extra)]

    all_nips = matched_nips + extra_nips
    rows = [{"nip": nip, "company_name": _fake_company_name()} for nip in all_nips]

    return pd.DataFrame(rows)



def main():
    print("[generator] Creating transactions.xlsx …")
    transactions_df = generate_transactions(n_nips=40, n_rows=150)
    transactions_df.to_excel("transactions.xlsx", index=False)
    print(f"[generator] transactions.xlsx saved — {len(transactions_df)} rows")

    valid_nips = (
        transactions_df[transactions_df["nip"].str.len() == 10]["nip"]
        .unique()
        .tolist()
    )

    print("[generator] Creating companies.xlsx …")
    companies_df = generate_companies(transaction_nips=valid_nips, extra=15)
    companies_df.to_excel("companies.xlsx", index=False)
    print(f"[generator] companies.xlsx saved — {len(companies_df)} rows")


if __name__ == "__main__":
    main()
