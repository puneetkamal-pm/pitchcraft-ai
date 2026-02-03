#!/usr/bin/env python3
"""
PitchCraftAI MVP
AI-powered DCF model generator for investment banking analysts

Usage:
    python main.py TICKER [--output OUTPUT_PATH] [--interactive]

Example:
    python main.py MSFT --output msft_dcf.xlsx
    python main.py AAPL --interactive
"""

import argparse
import json
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from data.sec_fetcher import fetch_company, CompanyFinancials
from core.question_generator import QuestionGenerator, DCFAssumptions
from models.dcf_professional import generate_dcf_model


def print_banner():
    """Print application banner"""
    print("""
╔═══════════════════════════════════════════════════════════════╗
║                      PitchCraftAI MVP                         ║
║           AI-Powered DCF Model Generator                      ║
╚═══════════════════════════════════════════════════════════════╝
    """)


def print_company_summary(fin: CompanyFinancials):
    """Print company financial summary"""
    print(f"\n{'─'*60}")
    print(f"  {fin.name} ({fin.ticker})")
    print(f"{'─'*60}")

    if fin.revenue:
        print(f"\n  Revenue (last {len(fin.revenue)} years):")
        for year, rev in zip(fin.revenue_years, fin.revenue):
            print(f"    {year}: ${rev:,.0f}M")

        if fin.revenue_growth:
            avg_growth = sum(fin.revenue_growth) / len(fin.revenue_growth)
            print(f"\n  Avg Revenue Growth: {avg_growth:.1%}")

    print(f"  EBITDA Margin: {fin.ebitda_margin:.1%}")
    print(f"\n  Balance Sheet:")
    print(f"    Total Debt: ${fin.total_debt:,.0f}M")
    print(f"    Cash: ${fin.cash:,.0f}M")
    print(f"    Net Debt: ${fin.total_debt - fin.cash:,.0f}M")
    print(f"    Shares Outstanding: {fin.shares_outstanding:.1f}M")
    print(f"{'─'*60}\n")


def interactive_mode(generator: QuestionGenerator) -> dict:
    """Run interactive Q&A to collect assumptions"""
    questions = generator.generate_questions()
    answers = {}

    print("\n" + "="*60)
    print("  DCF ASSUMPTION QUESTIONNAIRE")
    print("  Press Enter to accept default, or type new value")
    print("="*60)

    current_category = None

    for q in questions:
        if q.category != current_category:
            current_category = q.category
            print(f"\n{'─'*40}")
            print(f"  {current_category}")
            print(f"{'─'*40}")

        # Format default display
        if q.value_type == "percent":
            default_display = f"{q.default_value:.1%}"
        elif q.value_type == "number":
            default_display = f"{q.default_value:.2f}"
        else:
            default_display = str(q.default_value)

        print(f"\n{q.question}")
        print(f"  Hint: {q.hint}")
        user_input = input(f"  [{default_display}]: ").strip()

        if user_input:
            try:
                if q.value_type == "percent":
                    # Handle both "10" and "0.10" and "10%"
                    user_input = user_input.replace("%", "")
                    val = float(user_input)
                    if val > 1:  # User entered "10" meaning 10%
                        val = val / 100
                    answers[q.id] = val
                else:
                    answers[q.id] = float(user_input)
            except ValueError:
                print(f"  Invalid input, using default: {default_display}")
                answers[q.id] = q.default_value
        else:
            answers[q.id] = q.default_value

    return answers


def quick_mode(generator: QuestionGenerator) -> dict:
    """Use all defaults - no questions asked"""
    return generator.get_defaults()


def run(ticker: str, output_path: str = None, interactive: bool = False):
    """Main execution flow"""
    print_banner()

    # Step 1: Fetch company data
    print(f"Fetching SEC data for {ticker}...")
    financials = fetch_company(ticker)

    if not financials:
        print(f"\nError: Could not fetch data for {ticker}")
        print("Make sure the ticker is valid and the company files with SEC.")
        return None

    print_company_summary(financials)

    # Step 2: Generate questions
    generator = QuestionGenerator(financials)

    # Step 3: Collect assumptions
    if interactive:
        answers = interactive_mode(generator)
    else:
        print("Using intelligent defaults based on company data...")
        answers = quick_mode(generator)

    # Step 4: Create assumptions object
    assumptions = generator.create_assumptions_from_answers(answers)

    # Step 5: Generate DCF model
    if not output_path:
        output_path = f"{ticker.lower()}_dcf_model.xlsx"

    print(f"\nGenerating DCF model...")
    output_file = generate_dcf_model(assumptions, output_path)

    print(f"\n{'='*60}")
    print(f"  SUCCESS!")
    print(f"  DCF model saved to: {output_file}")
    print(f"{'='*60}")

    # Print key assumptions used
    print(f"\n  Key Assumptions Used:")
    print(f"    Base Revenue: ${assumptions.base_revenue:,.0f}M")
    print(f"    EBITDA Margin: {assumptions.ebitda_margin:.1%}")
    print(f"    WACC: {assumptions.risk_free_rate + assumptions.beta * assumptions.equity_risk_premium:.1%} (calculated)")
    print(f"    Terminal Growth: {assumptions.terminal_growth:.1%}")
    print(f"    Exit Multiple: {assumptions.exit_multiple:.1f}x EV/EBITDA")

    return output_file


def main():
    parser = argparse.ArgumentParser(
        description="PitchCraftAI - Generate professional DCF models from SEC data",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py AAPL                    # Quick mode with intelligent defaults
  python main.py MSFT -i                 # Interactive mode with Q&A
  python main.py GOOGL -o google_dcf.xlsx  # Specify output file
        """
    )

    parser.add_argument("ticker", help="Stock ticker symbol (e.g., AAPL, MSFT)")
    parser.add_argument("-o", "--output", help="Output Excel file path")
    parser.add_argument("-i", "--interactive", action="store_true",
                        help="Run in interactive mode with Q&A")

    args = parser.parse_args()

    result = run(
        ticker=args.ticker.upper(),
        output_path=args.output,
        interactive=args.interactive
    )

    if result:
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
