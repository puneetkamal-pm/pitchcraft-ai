"""
SEC EDGAR Data Fetcher
Pulls company financials from public SEC filings
"""

import requests
import json
from typing import Dict, Optional, List
from dataclasses import dataclass, asdict


@dataclass
class CompanyFinancials:
    """Core financial data for DCF modeling"""
    ticker: str
    name: str
    cik: str

    # Income Statement
    revenue: List[float]  # Last 3-5 years
    revenue_years: List[int]
    ebitda: List[float]
    net_income: List[float]

    # Balance Sheet
    total_assets: float
    total_debt: float
    cash: float
    shares_outstanding: float

    # Derived
    ebitda_margin: float
    revenue_growth: List[float]

    # Market Data (if available)
    market_cap: Optional[float] = None
    beta: Optional[float] = None


class SECFetcher:
    """Fetch company data from SEC EDGAR API"""

    BASE_URL = "https://data.sec.gov"
    COMPANY_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"

    HEADERS = {
        "User-Agent": "PitchCraftAI contact@example.com",  # SEC requires identification
        "Accept": "application/json"
    }

    # SEC XBRL taxonomy tags for key financials (ordered by preference - newer standards first)
    REVENUE_TAGS = [
        "RevenueFromContractWithCustomerExcludingAssessedTax",  # ASC 606 (post-2018)
        "Revenues",
        "SalesRevenueNet",
        "TotalRevenuesAndOtherIncome",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
    ]

    EBITDA_TAGS = [
        "OperatingIncomeLoss",  # Proxy - will need D&A added back
    ]

    NET_INCOME_TAGS = [
        "NetIncomeLoss",
        "ProfitLoss"
    ]

    ASSETS_TAGS = ["Assets"]
    DEBT_TAGS = ["LongTermDebt", "LongTermDebtNoncurrent"]
    CASH_TAGS = ["CashAndCashEquivalentsAtCarryingValue", "Cash"]
    SHARES_TAGS = ["CommonStockSharesOutstanding", "WeightedAverageNumberOfSharesOutstandingBasic"]
    DA_TAGS = ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization"]

    def __init__(self):
        self._ticker_to_cik: Dict[str, str] = {}
        self._load_ticker_map()

    def _load_ticker_map(self):
        """Load ticker to CIK mapping from SEC"""
        try:
            resp = requests.get(self.COMPANY_TICKERS_URL, headers=self.HEADERS, timeout=10)
            resp.raise_for_status()
            data = resp.json()

            for entry in data.values():
                ticker = entry.get("ticker", "").upper()
                cik = str(entry.get("cik_str", "")).zfill(10)
                if ticker and cik:
                    self._ticker_to_cik[ticker] = cik
        except Exception as e:
            print(f"Warning: Could not load ticker map: {e}")

    def get_cik(self, ticker: str) -> Optional[str]:
        """Get CIK for a ticker symbol"""
        return self._ticker_to_cik.get(ticker.upper())

    def _get_company_facts(self, cik: str) -> Optional[Dict]:
        """Fetch all XBRL facts for a company"""
        url = f"{self.BASE_URL}/api/xbrl/companyfacts/CIK{cik}.json"
        try:
            resp = requests.get(url, headers=self.HEADERS, timeout=15)
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
            print(f"Error fetching company facts: {e}")
            return None

    def _extract_values(self, facts: Dict, tags: List[str], units: str = "USD") -> List[tuple]:
        """Extract annual values for given XBRL tags"""
        results = []

        us_gaap = facts.get("facts", {}).get("us-gaap", {})

        for tag in tags:
            if tag not in us_gaap:
                continue

            tag_data = us_gaap[tag]
            unit_data = tag_data.get("units", {}).get(units, [])

            for entry in unit_data:
                # Only get 10-K (annual) filings
                form = entry.get("form", "")
                if form not in ["10-K", "10-K/A"]:
                    continue

                # Only get full year (FY) data, not quarterly
                fp = entry.get("fp", "")
                if fp != "FY":
                    continue

                fy = entry.get("fy")
                val = entry.get("val")

                if fy and val is not None:
                    results.append((fy, val))

            if results:
                break  # Use first matching tag

        # Dedupe by year - keep the LARGEST value (consolidated total, not segments)
        by_year = {}
        for fy, val in results:
            if fy not in by_year or val > by_year[fy]:
                by_year[fy] = val

        return sorted(by_year.items(), key=lambda x: x[0])

    def _extract_latest(self, facts: Dict, tags: List[str], units: str = "USD") -> Optional[float]:
        """Extract most recent value for given tags"""
        values = self._extract_values(facts, tags, units)
        if values:
            return values[-1][1]
        return None

    def _extract_shares(self, facts: Dict) -> Optional[float]:
        """Extract shares outstanding (in shares, not USD)"""
        us_gaap = facts.get("facts", {}).get("us-gaap", {})

        for tag in self.SHARES_TAGS:
            if tag not in us_gaap:
                continue

            tag_data = us_gaap[tag]
            unit_data = tag_data.get("units", {}).get("shares", [])

            # Get most recent 10-K value
            latest = None
            latest_year = 0

            for entry in unit_data:
                form = entry.get("form", "")
                if form not in ["10-K", "10-K/A"]:
                    continue
                fy = entry.get("fy", 0)
                if fy > latest_year:
                    latest_year = fy
                    latest = entry.get("val")

            if latest:
                return latest

        return None

    def fetch(self, ticker: str) -> Optional[CompanyFinancials]:
        """Fetch all financials for a company"""
        cik = self.get_cik(ticker)
        if not cik:
            print(f"Could not find CIK for ticker: {ticker}")
            return None

        facts = self._get_company_facts(cik)
        if not facts:
            return None

        company_name = facts.get("entityName", ticker)

        # Extract historical revenue
        revenue_data = self._extract_values(facts, self.REVENUE_TAGS)
        if not revenue_data:
            print(f"No revenue data found for {ticker}")
            return None

        revenue_years = [r[0] for r in revenue_data[-5:]]
        revenues = [r[1] / 1_000_000 for r in revenue_data[-5:]]  # Convert to millions

        # Calculate revenue growth
        growth_rates = []
        for i in range(1, len(revenues)):
            if revenues[i-1] > 0:
                growth = (revenues[i] - revenues[i-1]) / revenues[i-1]
                growth_rates.append(growth)

        # Extract other data
        net_income_data = self._extract_values(facts, self.NET_INCOME_TAGS)
        net_incomes = [n[1] / 1_000_000 for n in net_income_data[-5:]] if net_income_data else []

        # Operating income as EBITDA proxy
        op_income_data = self._extract_values(facts, self.EBITDA_TAGS)
        da_data = self._extract_values(facts, self.DA_TAGS)

        ebitdas = []
        if op_income_data:
            for i, (year, op_inc) in enumerate(op_income_data[-5:]):
                da = 0
                # Find matching D&A
                for da_year, da_val in da_data:
                    if da_year == year:
                        da = da_val
                        break
                ebitdas.append((op_inc + da) / 1_000_000)

        # Balance sheet items
        total_assets = self._extract_latest(facts, self.ASSETS_TAGS)
        total_debt = self._extract_latest(facts, self.DEBT_TAGS) or 0
        cash = self._extract_latest(facts, self.CASH_TAGS) or 0
        shares = self._extract_shares(facts)

        # Calculate EBITDA margin
        ebitda_margin = 0
        if ebitdas and revenues:
            ebitda_margin = ebitdas[-1] / revenues[-1] if revenues[-1] > 0 else 0

        return CompanyFinancials(
            ticker=ticker.upper(),
            name=company_name,
            cik=cik,
            revenue=revenues,
            revenue_years=revenue_years,
            ebitda=ebitdas,
            net_income=net_incomes,
            total_assets=(total_assets / 1_000_000) if total_assets else 0,
            total_debt=(total_debt / 1_000_000) if total_debt else 0,
            cash=(cash / 1_000_000) if cash else 0,
            shares_outstanding=(shares / 1_000_000) if shares else 0,
            ebitda_margin=ebitda_margin,
            revenue_growth=growth_rates
        )


def fetch_company(ticker: str) -> Optional[CompanyFinancials]:
    """Convenience function to fetch company data"""
    fetcher = SECFetcher()
    return fetcher.fetch(ticker)


if __name__ == "__main__":
    # Test with Apple
    import sys
    ticker = sys.argv[1] if len(sys.argv) > 1 else "AAPL"

    print(f"Fetching data for {ticker}...")
    data = fetch_company(ticker)

    if data:
        print(f"\n{data.name} ({data.ticker})")
        print(f"CIK: {data.cik}")
        print(f"\nRevenue (last {len(data.revenue)} years): {data.revenue_years}")
        print(f"  Values ($M): {[f'{r:,.0f}' for r in data.revenue]}")
        print(f"  Growth rates: {[f'{g:.1%}' for g in data.revenue_growth]}")
        print(f"\nEBITDA ($M): {[f'{e:,.0f}' for e in data.ebitda]}")
        print(f"EBITDA Margin: {data.ebitda_margin:.1%}")
        print(f"\nBalance Sheet:")
        print(f"  Total Assets: ${data.total_assets:,.0f}M")
        print(f"  Total Debt: ${data.total_debt:,.0f}M")
        print(f"  Cash: ${data.cash:,.0f}M")
        print(f"  Net Debt: ${data.total_debt - data.cash:,.0f}M")
        print(f"  Shares Outstanding: {data.shares_outstanding:.1f}M")
    else:
        print("Failed to fetch data")
