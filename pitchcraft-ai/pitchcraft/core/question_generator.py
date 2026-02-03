"""
DCF Assumption Question Generator
Generates comprehensive IB-quality questions based on pulled company data
"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from data.sec_fetcher import CompanyFinancials


@dataclass
class DCFQuestion:
    """A question for the analyst to answer"""
    id: str
    category: str
    subcategory: str
    question: str
    default_value: Any
    value_type: str  # 'percent', 'number', 'currency', 'years', 'multiple'
    hint: str
    min_value: Optional[float] = None
    max_value: Optional[float] = None


@dataclass
class DCFAssumptions:
    """Collected assumptions for DCF model - comprehensive IB quality"""
    # Company Info
    company_name: str
    ticker: str

    # Revenue Projections
    base_revenue: float
    revenue_growth_y1: float
    revenue_growth_y2: float
    revenue_growth_y3: float
    revenue_growth_y4: float
    revenue_growth_y5: float

    # Margins & Operations
    ebitda_margin: float
    ebitda_margin_y5: float  # Terminal margin (can expand/contract)
    da_pct_revenue: float
    capex_pct_revenue: float
    maintenance_capex_pct: float  # vs growth capex
    sbc_pct_revenue: float  # Stock-based comp

    # Working Capital Detail
    days_sales_outstanding: float  # DSO
    days_inventory_outstanding: float  # DIO
    days_payables_outstanding: float  # DPO
    other_working_capital_pct: float

    # Tax
    tax_rate: float
    nol_balance: float  # Net operating losses

    # WACC Components - Full buildup
    risk_free_rate: float
    equity_risk_premium: float
    unlevered_beta: float
    size_premium: float  # Small cap premium
    company_specific_risk: float
    pre_tax_cost_of_debt: float
    target_debt_to_equity: float

    # Terminal Value
    terminal_growth: float
    exit_ebitda_multiple: float
    exit_revenue_multiple: float  # Cross-check

    # Capital Structure
    total_debt: float
    cash: float
    minority_interest: float
    preferred_stock: float
    shares_outstanding: float
    options_dilution: float  # Diluted shares

    # Model Settings
    use_mid_year_convention: bool = True
    projection_years: int = 5


class QuestionGenerator:
    """Generate comprehensive DCF assumption questions based on company data"""

    def __init__(self, financials: CompanyFinancials):
        self.fin = financials

    def _calc_avg_growth(self) -> float:
        """Calculate average historical revenue growth"""
        if not self.fin.revenue_growth:
            return 0.05
        return sum(self.fin.revenue_growth) / len(self.fin.revenue_growth)

    def _calc_implied_beta(self) -> float:
        """Estimate unlevered beta based on company characteristics"""
        margin = self.fin.ebitda_margin
        if margin > 0.30:
            return 0.85
        elif margin > 0.20:
            return 1.0
        elif margin > 0.10:
            return 1.15
        else:
            return 1.3

    def _estimate_dso(self) -> float:
        """Estimate DSO based on business type"""
        return 45.0  # Default - would calculate from AR/Revenue

    def _estimate_dio(self) -> float:
        """Estimate DIO - 0 for software, higher for hardware"""
        return 15.0  # Low for software companies

    def _estimate_dpo(self) -> float:
        """Estimate DPO"""
        return 35.0

    def generate_questions(self) -> List[DCFQuestion]:
        """Generate all DCF assumption questions - comprehensive set"""
        avg_growth = self._calc_avg_growth()
        implied_beta = self._calc_implied_beta()
        net_debt = self.fin.total_debt - self.fin.cash

        questions = [
            # ============ REVENUE PROJECTIONS ============
            DCFQuestion(
                id="revenue_growth_y1",
                category="Revenue Build",
                subcategory="Growth Rates",
                question=f"Year 1 Revenue Growth",
                default_value=min(avg_growth * 0.9, 0.20),
                value_type="percent",
                hint=f"Historical: {[f'{g:.1%}' for g in self.fin.revenue_growth[-3:]]}",
                min_value=-0.20,
                max_value=0.50
            ),
            DCFQuestion(
                id="revenue_growth_y2",
                category="Revenue Build",
                subcategory="Growth Rates",
                question="Year 2 Revenue Growth",
                default_value=min(avg_growth * 0.80, 0.15),
                value_type="percent",
                hint="Deceleration as base grows",
                min_value=-0.20,
                max_value=0.40
            ),
            DCFQuestion(
                id="revenue_growth_y3",
                category="Revenue Build",
                subcategory="Growth Rates",
                question="Year 3 Revenue Growth",
                default_value=min(avg_growth * 0.65, 0.12),
                value_type="percent",
                hint="Approaching maturity",
                min_value=-0.15,
                max_value=0.30
            ),
            DCFQuestion(
                id="revenue_growth_y4",
                category="Revenue Build",
                subcategory="Growth Rates",
                question="Year 4 Revenue Growth",
                default_value=min(avg_growth * 0.50, 0.08),
                value_type="percent",
                hint="Near terminal growth",
                min_value=-0.10,
                max_value=0.25
            ),
            DCFQuestion(
                id="revenue_growth_y5",
                category="Revenue Build",
                subcategory="Growth Rates",
                question="Year 5 Revenue Growth",
                default_value=min(avg_growth * 0.40, 0.05),
                value_type="percent",
                hint="Final projection year",
                min_value=-0.10,
                max_value=0.20
            ),

            # ============ MARGINS ============
            DCFQuestion(
                id="ebitda_margin",
                category="Operating Model",
                subcategory="Margins",
                question=f"Base EBITDA Margin",
                default_value=max(self.fin.ebitda_margin, 0.10) if self.fin.ebitda_margin < 0.80 else 0.25,
                value_type="percent",
                hint=f"Current: {self.fin.ebitda_margin:.1%}",
                min_value=0.05,
                max_value=0.60
            ),
            DCFQuestion(
                id="ebitda_margin_y5",
                category="Operating Model",
                subcategory="Margins",
                question="Terminal EBITDA Margin (Y5)",
                default_value=min(max(self.fin.ebitda_margin, 0.10) + 0.03, 0.40) if self.fin.ebitda_margin < 0.80 else 0.28,
                value_type="percent",
                hint="Margin expansion/contraction over forecast",
                min_value=0.05,
                max_value=0.60
            ),
            DCFQuestion(
                id="sbc_pct_revenue",
                category="Operating Model",
                subcategory="Margins",
                question="Stock-Based Comp (% Rev)",
                default_value=0.08,
                value_type="percent",
                hint="Tech companies: 5-15%. Add back for EBITDA, real cash cost.",
                min_value=0.0,
                max_value=0.25
            ),

            # ============ D&A / CAPEX ============
            DCFQuestion(
                id="da_pct_revenue",
                category="Operating Model",
                subcategory="D&A / CapEx",
                question="D&A (% of Revenue)",
                default_value=0.04,
                value_type="percent",
                hint="Software: 3-5%, Hardware: 5-10%",
                min_value=0.01,
                max_value=0.15
            ),
            DCFQuestion(
                id="capex_pct_revenue",
                category="Operating Model",
                subcategory="D&A / CapEx",
                question="Total CapEx (% of Revenue)",
                default_value=0.05,
                value_type="percent",
                hint="Maintenance + Growth CapEx",
                min_value=0.01,
                max_value=0.20
            ),
            DCFQuestion(
                id="maintenance_capex_pct",
                category="Operating Model",
                subcategory="D&A / CapEx",
                question="Maintenance CapEx (% of Total)",
                default_value=0.60,
                value_type="percent",
                hint="Rest is growth CapEx. Maintenance ≈ D&A at steady state.",
                min_value=0.30,
                max_value=1.0
            ),

            # ============ WORKING CAPITAL ============
            DCFQuestion(
                id="days_sales_outstanding",
                category="Working Capital",
                subcategory="Receivables",
                question="Days Sales Outstanding (DSO)",
                default_value=self._estimate_dso(),
                value_type="number",
                hint="AR collection period. SaaS: 30-45, Enterprise: 45-90",
                min_value=15,
                max_value=120
            ),
            DCFQuestion(
                id="days_inventory_outstanding",
                category="Working Capital",
                subcategory="Inventory",
                question="Days Inventory Outstanding (DIO)",
                default_value=self._estimate_dio(),
                value_type="number",
                hint="Software: ~0, Hardware: 30-90",
                min_value=0,
                max_value=180
            ),
            DCFQuestion(
                id="days_payables_outstanding",
                category="Working Capital",
                subcategory="Payables",
                question="Days Payables Outstanding (DPO)",
                default_value=self._estimate_dpo(),
                value_type="number",
                hint="AP payment period. Higher = better cash conversion.",
                min_value=15,
                max_value=120
            ),

            # ============ TAX ============
            DCFQuestion(
                id="tax_rate",
                category="Tax",
                subcategory="Effective Rate",
                question="Effective Tax Rate",
                default_value=0.24,
                value_type="percent",
                hint="Federal 21% + State ~3%. Check for NOLs.",
                min_value=0.0,
                max_value=0.40
            ),
            DCFQuestion(
                id="nol_balance",
                category="Tax",
                subcategory="NOLs",
                question="NOL Balance ($M)",
                default_value=0,
                value_type="currency",
                hint="Net Operating Loss carryforwards. Check 10-K.",
                min_value=0,
                max_value=10000
            ),

            # ============ WACC BUILDUP ============
            DCFQuestion(
                id="risk_free_rate",
                category="WACC",
                subcategory="Cost of Equity",
                question="Risk-Free Rate (10Y UST)",
                default_value=0.043,
                value_type="percent",
                hint="Current 10Y Treasury yield",
                min_value=0.01,
                max_value=0.10
            ),
            DCFQuestion(
                id="equity_risk_premium",
                category="WACC",
                subcategory="Cost of Equity",
                question="Equity Risk Premium",
                default_value=0.055,
                value_type="percent",
                hint="Duff & Phelps: 5.5%. Historical: 5-7%",
                min_value=0.03,
                max_value=0.10
            ),
            DCFQuestion(
                id="unlevered_beta",
                category="WACC",
                subcategory="Cost of Equity",
                question=f"Unlevered Beta",
                default_value=implied_beta,
                value_type="number",
                hint="From comps. Unlever at comp D/E, relever at target.",
                min_value=0.5,
                max_value=2.0
            ),
            DCFQuestion(
                id="size_premium",
                category="WACC",
                subcategory="Cost of Equity",
                question="Size Premium",
                default_value=0.01,
                value_type="percent",
                hint="Small cap: 2-4%, Mid cap: 0-2%, Large: 0%",
                min_value=0.0,
                max_value=0.06
            ),
            DCFQuestion(
                id="company_specific_risk",
                category="WACC",
                subcategory="Cost of Equity",
                question="Company-Specific Risk Premium",
                default_value=0.01,
                value_type="percent",
                hint="Execution risk, key man, customer concentration",
                min_value=0.0,
                max_value=0.05
            ),
            DCFQuestion(
                id="pre_tax_cost_of_debt",
                category="WACC",
                subcategory="Cost of Debt",
                question="Pre-Tax Cost of Debt",
                default_value=0.065,
                value_type="percent",
                hint="Based on credit rating. IG: 5-7%, HY: 8-12%",
                min_value=0.03,
                max_value=0.15
            ),
            DCFQuestion(
                id="target_debt_to_equity",
                category="WACC",
                subcategory="Capital Structure",
                question="Target Debt / Equity",
                default_value=0.25,
                value_type="number",
                hint="For relevering beta and WACC weights",
                min_value=0.0,
                max_value=2.0
            ),

            # ============ TERMINAL VALUE ============
            DCFQuestion(
                id="terminal_growth",
                category="Terminal Value",
                subcategory="Perpetuity",
                question="Perpetuity Growth Rate",
                default_value=0.025,
                value_type="percent",
                hint="Cannot exceed long-term GDP (2-3%)",
                min_value=0.0,
                max_value=0.04
            ),
            DCFQuestion(
                id="exit_ebitda_multiple",
                category="Terminal Value",
                subcategory="Exit Multiples",
                question="Exit EV/EBITDA Multiple",
                default_value=12.0,
                value_type="multiple",
                hint="From trading comps. Software: 12-20x, Media: 8-12x",
                min_value=4.0,
                max_value=30.0
            ),
            DCFQuestion(
                id="exit_revenue_multiple",
                category="Terminal Value",
                subcategory="Exit Multiples",
                question="Exit EV/Revenue Multiple",
                default_value=3.0,
                value_type="multiple",
                hint="Cross-check. High-growth SaaS: 5-10x, Mature: 1-3x",
                min_value=0.5,
                max_value=15.0
            ),

            # ============ CAPITAL STRUCTURE ============
            DCFQuestion(
                id="minority_interest",
                category="EV Bridge",
                subcategory="Adjustments",
                question="Minority Interest ($M)",
                default_value=0,
                value_type="currency",
                hint="Non-controlling interests. Check balance sheet.",
                min_value=0,
                max_value=50000
            ),
            DCFQuestion(
                id="options_dilution",
                category="EV Bridge",
                subcategory="Share Count",
                question="Options/RSU Dilution (%)",
                default_value=0.03,
                value_type="percent",
                hint="Treasury stock method dilution. Tech: 3-8%",
                min_value=0.0,
                max_value=0.15
            ),
        ]

        return questions

    def create_assumptions_from_answers(self, answers: Dict[str, Any]) -> DCFAssumptions:
        """Create DCFAssumptions from answered questions"""

        # Calculate NWC from DSO/DIO/DPO
        dso = answers.get("days_sales_outstanding", 45)
        dio = answers.get("days_inventory_outstanding", 15)
        dpo = answers.get("days_payables_outstanding", 35)
        # Cash conversion cycle as proxy for NWC
        ccc = dso + dio - dpo

        return DCFAssumptions(
            company_name=self.fin.name,
            ticker=self.fin.ticker,
            base_revenue=self.fin.revenue[-1] if self.fin.revenue else 0,

            revenue_growth_y1=answers.get("revenue_growth_y1", 0.10),
            revenue_growth_y2=answers.get("revenue_growth_y2", 0.08),
            revenue_growth_y3=answers.get("revenue_growth_y3", 0.06),
            revenue_growth_y4=answers.get("revenue_growth_y4", 0.05),
            revenue_growth_y5=answers.get("revenue_growth_y5", 0.04),

            ebitda_margin=answers.get("ebitda_margin", self.fin.ebitda_margin if self.fin.ebitda_margin < 0.80 else 0.25),
            ebitda_margin_y5=answers.get("ebitda_margin_y5", 0.28),
            da_pct_revenue=answers.get("da_pct_revenue", 0.04),
            capex_pct_revenue=answers.get("capex_pct_revenue", 0.05),
            maintenance_capex_pct=answers.get("maintenance_capex_pct", 0.60),
            sbc_pct_revenue=answers.get("sbc_pct_revenue", 0.08),

            days_sales_outstanding=dso,
            days_inventory_outstanding=dio,
            days_payables_outstanding=dpo,
            other_working_capital_pct=0.02,

            tax_rate=answers.get("tax_rate", 0.24),
            nol_balance=answers.get("nol_balance", 0),

            risk_free_rate=answers.get("risk_free_rate", 0.043),
            equity_risk_premium=answers.get("equity_risk_premium", 0.055),
            unlevered_beta=answers.get("unlevered_beta", 1.0),
            size_premium=answers.get("size_premium", 0.01),
            company_specific_risk=answers.get("company_specific_risk", 0.01),
            pre_tax_cost_of_debt=answers.get("pre_tax_cost_of_debt", 0.065),
            target_debt_to_equity=answers.get("target_debt_to_equity", 0.25),

            terminal_growth=answers.get("terminal_growth", 0.025),
            exit_ebitda_multiple=answers.get("exit_ebitda_multiple", 12.0),
            exit_revenue_multiple=answers.get("exit_revenue_multiple", 3.0),

            total_debt=self.fin.total_debt,
            cash=self.fin.cash,
            minority_interest=answers.get("minority_interest", 0),
            preferred_stock=0,
            shares_outstanding=self.fin.shares_outstanding,
            options_dilution=answers.get("options_dilution", 0.03),

            use_mid_year_convention=True,
            projection_years=5
        )

    def get_defaults(self) -> Dict[str, Any]:
        """Get default values for all questions"""
        questions = self.generate_questions()
        return {q.id: q.default_value for q in questions}


def generate_questions_for_company(ticker: str) -> tuple:
    """Convenience function to generate questions for a company"""
    from data.sec_fetcher import fetch_company

    financials = fetch_company(ticker)
    if not financials:
        return None, None

    generator = QuestionGenerator(financials)
    questions = generator.generate_questions()

    return financials, questions


if __name__ == "__main__":
    ticker = sys.argv[1] if len(sys.argv) > 1 else "HUBS"

    financials, questions = generate_questions_for_company(ticker)

    if questions:
        print(f"\n{'='*60}")
        print(f"DCF Assumptions for {financials.name}")
        print(f"{'='*60}")

        current_category = None
        for q in questions:
            if q.category != current_category:
                current_category = q.category
                print(f"\n## {current_category}")
                print("-" * 40)

            print(f"\n{q.question}")
            if q.value_type == "percent":
                print(f"  Default: {q.default_value:.1%}")
            elif q.value_type == "multiple":
                print(f"  Default: {q.default_value:.1f}x")
            else:
                print(f"  Default: {q.default_value}")
            print(f"  Hint: {q.hint}")
