"""Generate MIS balance sheet and P&L templates for l10n_fi_mis_templates.

Writes data/mis_report_fi_bs.xml and data/mis_report_fi_pl.xml. The report
structure and expressions are defined below; KPI descriptions are read from
the Excel exports of the reports (columns kpi_ids/name and
kpi_ids/description).

Usage:
    python tools/generate_mis_templates.py \
        --bs-xlsx "Tase päivitetty v2.xlsx" --pl-xlsx "Tulos päivitetty v2.xlsx"

Requires openpyxl. Run pre-commit afterwards: prettier reformats the
generated XML files.
"""

import argparse
import re
import sys
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl

parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
parser.add_argument("--bs-xlsx", required=True, help="Balance sheet Excel export")
parser.add_argument("--pl-xlsx", required=True, help="Profit & loss Excel export")
parser.add_argument(
    "--out",
    default=str(Path(__file__).resolve().parent.parent / "data"),
    help="Output directory (default: the module's data directory)",
)
ARGS = parser.parse_args()
OUT = ARGS.out
MOD = "l10n_fi_oca"

PL_TYPES = (
    "('income', 'income_other', 'expense', 'expense_depreciation', "
    "'expense_direct_cost')"
)


def tags(*names):
    refs = ", ".join(f"ref('{MOD}.account_tag_{n}').id" for n in names)
    # Tuple, not list: mis_builder's account selector regex (\[.*?\]) stops at
    # the first "]", so no square brackets are allowed inside the domain.
    return f"('tag_ids', 'in', ({refs},))"


def bs_asset(*names):
    return f"bale[{tags(*names)}][]"


def bs_liab(*names):
    return f"-bale[{tags(*names)}][]"


def pl(*names):
    return f"-balp[{tags(*names)}][]"


# Row styles (style xml ids are in STYLES below)
H = "header"  # bold, indent 0
HI = "header_indent"  # bold, indent 1
HDI = "header_double_indent"  # bold, indent 2
D = "details"  # indent 1
DD = "details_double_indent"  # indent 2
DTI = "details_triple_indent"  # indent 3


# Row: (name, description, expression, style).
# acc = account row, grp = sum of other rows; both render the same way
# ("Display details by account" is off on every row).
def acc(name, desc, expr, style):
    return (name, desc, expr, style)


def grp(name, desc, expr, style):
    return (name, desc, expr, style)


BS = [
    grp("assets", "VASTAAVAA", "noncurrent_assets+current_assets", H),
    grp(
        "noncurrent_assets",
        "PYSYVÄT VASTAAVAT",
        "account_tag_intangible_assets+tangible_assets+investments",
        HI,
    ),
    grp(
        "account_tag_intangible_assets",
        "Aineettomat hyödykkeet",
        "development_costs+immaterial_rights+goodwill+other_intangible_assets+advance_payments_intangible",
        HDI,
    ),
    acc("development_costs", "Kehittämismenot", bs_asset("development_costs"), DTI),
    acc(
        "immaterial_rights", "Aineettomat oikeudet", bs_asset("immaterial_rights"), DTI
    ),
    acc("goodwill", "Liikearvo", bs_asset("goodwill"), DTI),
    acc(
        "other_intangible_assets",
        "Muut aineettomat hyödykkeet",
        bs_asset("other_long_term_expenses"),
        DTI,
    ),
    acc(
        "advance_payments_intangible",
        "Ennakkomaksut",
        bs_asset("advance_payments_intangible"),
        DTI,
    ),
    grp(
        "tangible_assets",
        "Aineelliset hyödykkeet",
        "land_and_water_areas+buildings+machines_and_hardware+other_tangible_assets+payments_on_account_tangible",
        HDI,
    ),
    acc(
        "land_and_water_areas",
        "Maa- ja vesialueet",
        bs_asset("land_and_water_areas_owned", "land_and_water_rights_lease"),
        DTI,
    ),
    acc(
        "buildings",
        "Rakennukset ja rakennelmat",
        bs_asset("buildings_owned", "building_leases"),
        DTI,
    ),
    acc(
        "machines_and_hardware",
        "Koneet ja kalusto",
        bs_asset("machines_and_hardware"),
        DTI,
    ),
    acc(
        "other_tangible_assets",
        "Muut aineelliset hyödykkeet",
        bs_asset("other_tangible_assets"),
        DTI,
    ),
    acc(
        "payments_on_account_tangible",
        "Ennakkomaksut ja keskeneräiset hankinnat",
        bs_asset("payments_on_account_tangible"),
        DTI,
    ),
    grp(
        "investments",
        "Sijoitukset",
        "shares_in_group_companies+investments_from_group+shares_in_associated_companies"
        "+investments_from_associated_companies+other_shares_and_participations+other_investments",
        HDI,
    ),
    acc(
        "shares_in_group_companies",
        "Osuudet saman konsernin yrityksissä",
        bs_asset("shares_in_group_companies"),
        DTI,
    ),
    acc(
        "investments_from_group",
        "Saamiset saman konsernin yrityksiltä",
        bs_asset("receivables_from_group_company_investments"),
        DTI,
    ),
    acc(
        "shares_in_associated_companies",
        "Osuudet omistusyhteysyrityksissä",
        bs_asset("shares_in_associated_companies"),
        DTI,
    ),
    acc(
        "investments_from_associated_companies",
        "Saamiset omistusyhteysyrityksiltä",
        bs_asset("receivables_from_associated_companies_investments"),
        DTI,
    ),
    acc(
        "other_shares_and_participations",
        "Muut osakkeet ja osuudet",
        bs_asset("other_shares_and_participations"),
        DTI,
    ),
    acc(
        "other_investments",
        "Muut saamiset",
        bs_asset("other_receivables_investments"),
        DTI,
    ),
    grp(
        "current_assets",
        "VAIHTUVAT VASTAAVAT",
        "inventory+longterm_receivables+shortterm_receivables+financial_securities+liquidity",
        HI,
    ),
    grp(
        "inventory",
        "Vaihto-omaisuus",
        "materials_and_supplies+work_in_progress+finished_goods+other_current_assets+advance_payments",
        HDI,
    ),
    acc(
        "materials_and_supplies",
        "Aineet ja tarvikkeet",
        bs_asset("materials_and_supplies_inventories"),
        DTI,
    ),
    acc(
        "work_in_progress", "Keskeneräiset tuotteet", bs_asset("work_in_progress"), DTI
    ),
    acc(
        "finished_goods",
        "Valmiit tuotteet/tavarat",
        bs_asset("finished_goods_inventories", "goods_in_transit"),
        DTI,
    ),
    acc(
        "other_current_assets",
        "Muu vaihto-omaisuus",
        bs_asset("other_current_assets"),
        DTI,
    ),
    acc(
        "advance_payments",
        "Ennakkomaksut",
        bs_asset("advance_payments_for_property"),
        DTI,
    ),
    grp(
        "longterm_receivables",
        "Pitkäaikaiset saamiset",
        "longterm_trade_receivables+longterm_receivables_from_group+longterm_receivables_from_associated_companies"
        "+longterm_loan_receivables+longterm_deferred_taxes+longterm_receivables_other+longterm_unpaid_shares"
        "+longterm_prepaid_expenses",
        HDI,
    ),
    acc(
        "longterm_trade_receivables",
        "Myyntisaamiset",
        bs_asset("trade_receivables_long-term"),
        DTI,
    ),
    acc(
        "longterm_receivables_from_group",
        "Saamiset saman konsernin yrityksiltä",
        bs_asset("receivables_from_group_long-term"),
        DTI,
    ),
    acc(
        "longterm_receivables_from_associated_companies",
        "Saamiset omistusyhteysyrityksiltä",
        bs_asset("receivables_from_associated_companies_non-current"),
        DTI,
    ),
    acc(
        "longterm_loan_receivables",
        "Lainasaamiset",
        bs_asset("loan_receivables_long-term"),
        DTI,
    ),
    acc(
        "longterm_deferred_taxes",
        "Laskennalliset verosaamiset",
        bs_asset("deferred_tax_assets_are_long-term"),
        DTI,
    ),
    acc(
        "longterm_receivables_other",
        "Muut saamiset",
        bs_asset("other_long-term_receivables"),
        DTI,
    ),
    acc(
        "longterm_unpaid_shares",
        "Maksamattomat osakkeet/osuudet",
        bs_asset("unpaid_shares_of_long-term"),
        DTI,
    ),
    acc(
        "longterm_prepaid_expenses",
        "Siirtosaamiset",
        bs_asset("prepaid_expenses_and_long-term"),
        DTI,
    ),
    grp(
        "shortterm_receivables",
        "Lyhytaikaiset saamiset",
        "shortterm_trade_receivables+shortterm_receivables_from_group+shortterm_receivables_from_associated_companies"
        "+shortterm_loans+shortterm_deferred_tax_assets+shortterm_receivables_other+shortterm_unpaid_contributions"
        "+shortterm_prepayments_and_accrued_income",
        HDI,
    ),
    acc(
        "shortterm_trade_receivables",
        "Myyntisaamiset",
        bs_asset("accounts_receivable"),
        DTI,
    ),
    acc(
        "shortterm_receivables_from_group",
        "Saamiset saman konsernin yrityksiltä",
        bs_asset("receivables_from_group", "trade_group"),
        DTI,
    ),
    acc(
        "shortterm_receivables_from_associated_companies",
        "Saamiset omistusyhteysyrityksiltä",
        bs_asset(
            "receivables_from_associated_companies",
            "accounts_receivable_from_associated_companies",
        ),
        DTI,
    ),
    acc("shortterm_loans", "Lainasaamiset", bs_asset("loans"), DTI),
    acc(
        "shortterm_deferred_tax_assets",
        "Laskennalliset verosaamiset",
        bs_asset("deferred_tax_assets"),
        DTI,
    ),
    acc(
        "shortterm_receivables_other",
        "Muut saamiset",
        bs_asset("other_receivables"),
        DTI,
    ),
    acc(
        "shortterm_unpaid_contributions",
        "Maksamattomat osakkeet/osuudet",
        bs_asset("unpaid_contributions"),
        DTI,
    ),
    acc(
        "shortterm_prepayments_and_accrued_income",
        "Siirtosaamiset",
        bs_asset("prepayments_and_accrued_income"),
        DTI,
    ),
    grp(
        "financial_securities",
        "Rahoitusarvopaperit",
        "shares_in_consolidated_securities+other_shares_and_securities+other_financial_assets",
        HDI,
    ),
    acc(
        "shares_in_consolidated_securities",
        "Osuudet saman konsernin yrityksissä",
        bs_asset("shares_in_the_consolidated_financial_securities"),
        DTI,
    ),
    acc(
        "other_shares_and_securities",
        "Muut osakkeet ja osuudet",
        bs_asset("other_shares_and_marketable_securities"),
        DTI,
    ),
    acc(
        "other_financial_assets",
        "Muut arvopaperit",
        bs_asset("other_financial_assets"),
        DTI,
    ),
    # Liquidity also covers bank/cash accounts and untagged utility accounts
    # (bank suspense, outstanding receipts/payments, liquidity transfer)
    # that Odoo creates outside the chart template.
    acc(
        "liquidity",
        "Rahat ja pankkisaamiset",
        f"bale['|', '|', {tags('liquidity')}, ('account_type', '=', 'asset_cash'), "
        "'&', ('account_type', '=', 'asset_current'), ('tag_ids', '=', False)][]",
        HDI,
    ),
    grp(
        "liabilities_and_equity",
        "VASTATTAVAA",
        "capital+minority_interest+consolidation_reserve+accumulated_appropriations+provisions+liabilities",
        H,
    ),
    grp(
        "capital",
        "OMA PÄÄOMA",
        "stock_capital+share_premium_account+revaluation_reserve+other_reserves_header+retained_earnings+profit_loss",
        HI,
    ),
    acc(
        "stock_capital",
        "Osake-, osuus- tai muu vastaava pääoma",
        bs_liab("stock_capital"),
        HDI,
    ),
    acc(
        "share_premium_account",
        "Ylikurssirahasto",
        bs_liab("share_premium_account"),
        HDI,
    ),
    acc(
        "revaluation_reserve",
        "Arvonkorotusrahasto",
        bs_liab("revaluation_reserve"),
        HDI,
    ),
    grp(
        "other_reserves_header",
        "Muut rahastot",
        "invested_unrestriced_equity_fund+reserve+articles_of_association_or_fund_under_rules"
        "+fair_value_reserve+other_reserves",
        HDI,
    ),
    acc(
        "invested_unrestriced_equity_fund",
        "Sijoitetun vapaan oman pääoman rahasto",
        bs_liab("the_invested_unrestricted_equity_fund"),
        DTI,
    ),
    acc("reserve", "Vararahasto", bs_liab("reserve"), DTI),
    acc(
        "articles_of_association_or_fund_under_rules",
        "Yhtiöjärjestyksen tai sääntöjen mukaiset rahastot",
        bs_liab("the_articles_of_association_or_fund_under_the_rules"),
        DTI,
    ),
    acc(
        "fair_value_reserve",
        "Käyvän arvon rahasto",
        bs_liab("the_fair_value_reserve"),
        DTI,
    ),
    acc("other_reserves", "Muut rahastot", bs_liab("other_reserves"), DTI),
    # Prior years' result: retained earnings accounts, Odoo's undistributed
    # profits account and P&L of earlier fiscal years not closed to equity.
    acc(
        "retained_earnings",
        "Edellisten tilikausien voitto (tappio)",
        f"-bale['|', {tags('retained_earnings')}, "
        "('account_type', '=', 'equity_unaffected')][] - balu[]",
        HDI,
    ),
    # Current fiscal year result: P&L accounts from the start of the fiscal year.
    grp(
        "profit_loss",
        "Tilikauden voitto (tappio)",
        f"-bale[('account_type', 'in', {PL_TYPES})][]",
        HDI,
    ),
    acc("minority_interest", "VÄHEMMISTÖOSUUDET", bs_liab("minority_share"), HI),
    grp(
        "accumulated_appropriations",
        "TILINPÄÄTÖSSIIRTOJEN KERTYMÄ",
        "deprecation_difference+tax_based_reservations",
        HI,
    ),
    acc("deprecation_difference", "Poistoero", bs_liab("depreciation_difference"), HDI),
    acc(
        "tax_based_reservations",
        "Vapaaehtoiset varaukset",
        bs_liab("tax_based_reservations"),
        HDI,
    ),
    grp(
        "provisions",
        "PAKOLLISET VARAUKSET",
        "provisions_for_pensions+tax_provisions+other_provisions",
        HI,
    ),
    acc(
        "provisions_for_pensions",
        "Eläkevaraukset",
        bs_liab("provisions_for_pensions"),
        DTI,
    ),
    acc("tax_provisions", "Verovaraukset", bs_liab("tax_provisions"), DTI),
    acc(
        "other_provisions",
        "Muut pakolliset varaukset",
        bs_liab("other_provisions"),
        DTI,
    ),
    acc("consolidation_reserve", "KONSERNIRESERVI", bs_liab("badwill"), HDI),
    grp("liabilities", "VIERAS PÄÄOMA", "longterm_loans+shortterm_liabilities", HI),
    grp(
        "longterm_loans",
        "Pitkäaikainen",
        "longterm_subordinated_loans+longterm_bonds+longterm_convertible_bonds+longterm_loans_from_financial_institutions"
        "+longterm_pension_loans+longterm_received_advances+longterm_payables+longterm_bills"
        "+longterm_consolidated_liabilities+longterm_payables_to_associated_companies"
        "+longterm_deferred_tax_liabilities+longterm_liabilities_other+longterm_accrued",
        HDI,
    ),
    acc(
        "longterm_subordinated_loans",
        "Pääomalainat",
        bs_liab("subordinated_loans_long-term"),
        DTI,
    ),
    acc(
        "longterm_bonds", "Joukkovelkakirjalainat", bs_liab("bonds_are_long-term"), DTI
    ),
    acc(
        "longterm_convertible_bonds",
        "Vaihtovelkakirjalainat",
        bs_liab("convertible_bonds_long-term"),
        DTI,
    ),
    acc(
        "longterm_loans_from_financial_institutions",
        "Lainat rahoituslaitoksilta",
        bs_liab("loans_from_financial_institutions_long-term"),
        DTI,
    ),
    acc(
        "longterm_pension_loans", "Eläkelainat", bs_liab("pension_loans_long-term"), DTI
    ),
    acc(
        "longterm_received_advances",
        "Saadut ennakot",
        bs_liab("advances_received_long-term"),
        DTI,
    ),
    acc("longterm_payables", "Ostovelat", bs_liab("accounts_payable-current"), DTI),
    acc("longterm_bills", "Rahoitusvekselit", bs_liab("bills_of_long-term"), DTI),
    acc(
        "longterm_consolidated_liabilities",
        "Velat saman konsernin yrityksille",
        bs_liab("consolidated_long-term_liabilities"),
        DTI,
    ),
    acc(
        "longterm_payables_to_associated_companies",
        "Velat omistusyhteysyrityksille",
        bs_liab("payables_to_associated_companies_non-current"),
        DTI,
    ),
    acc(
        "longterm_deferred_tax_liabilities",
        "Laskennalliset verovelat",
        bs_liab("deferred_tax_liabilities_are_long-term"),
        DTI,
    ),
    acc(
        "longterm_liabilities_other",
        "Muut velat",
        bs_liab("other_long-term_liabilities"),
        DTI,
    ),
    acc("longterm_accrued", "Siirtovelat", bs_liab("accrued_long-term"), DTI),
    grp(
        "shortterm_liabilities",
        "Lyhytaikainen",
        "shortterm_capital_loans+shortterm_bonds+shortterm_convertible_bonds+shortterm_loans_from_financial_institutions"
        "+shortterm_pension_loans+shortterm_received_advances+shortterm_payables+shortterm_bills"
        "+shortterm_consolidated_liabilities+shortterm_payables_to_associated_companies"
        "+shortterm_deferred_tax_liabilities+shortterm_liabilities_other+shortterm_accrued_liabilities",
        HDI,
    ),
    acc(
        "shortterm_capital_loans",
        "Pääomalainat",
        bs_liab("capital_loans_are_short-term"),
        DTI,
    ),
    acc("shortterm_bonds", "Joukkovelkakirjalainat", bs_liab("bonds_short-term"), DTI),
    acc(
        "shortterm_convertible_bonds",
        "Vaihtovelkakirjalainat",
        bs_liab("convertible_bonds_short-term"),
        DTI,
    ),
    acc(
        "shortterm_loans_from_financial_institutions",
        "Lainat rahoituslaitoksilta",
        bs_liab("loans_from_financial_institutions_short-term"),
        DTI,
    ),
    acc(
        "shortterm_pension_loans",
        "Eläkelainat",
        bs_liab("pension_loans_short-term"),
        DTI,
    ),
    acc(
        "shortterm_received_advances",
        "Saadut ennakot",
        bs_liab("advances_received_short-term"),
        DTI,
    ),
    acc("shortterm_payables", "Ostovelat", bs_liab("accounts_payable"), DTI),
    acc("shortterm_bills", "Rahoitusvekselit", bs_liab("bills_of_short-term"), DTI),
    acc(
        "shortterm_consolidated_liabilities",
        "Velat saman konsernin yrityksille",
        bs_liab("trade_payables_to_group", "group_short-term_liabilities"),
        DTI,
    ),
    acc(
        "shortterm_payables_to_associated_companies",
        "Velat omistusyhteysyrityksille",
        bs_liab(
            "trade_payables_to_associated_companies", "payables_to_associated_companies"
        ),
        DTI,
    ),
    acc(
        "shortterm_deferred_tax_liabilities",
        "Laskennalliset verovelat",
        bs_liab("deferred_tax_liabilities_are_short-term"),
        DTI,
    ),
    acc(
        "shortterm_liabilities_other",
        "Muut velat",
        bs_liab("other_short-term_liabilities"),
        DTI,
    ),
    acc(
        "shortterm_accrued_liabilities",
        "Siirtovelat",
        bs_liab("accrued_liabilities_short-term"),
        DTI,
    ),
    # Shown only when non-zero (style hides empty rows): signals accounts
    # missing from the report or an otherwise unbalanced balance sheet.
    grp(
        "balance_check",
        "Tarkistus: vastaavaa - vastattavaa (pitää olla 0)",
        "assets-liabilities_and_equity",
        H,
    ),
]

PL = [
    acc("revenue", "LIIKEVAIHTO", pl("sales", "sales_adjustment_items"), H),
    acc(
        "account_tag_manufacturing_warehouse_change",
        "Valmiiden ja keskeneräisten tuotteiden varastojen muutos",
        pl("manufacturing_warehouse_change"),
        H,
    ),
    acc(
        "account_tag_manufacturing_own_use",
        "Valmistus omaan käyttöön",
        pl("manufacturing_own_use"),
        H,
    ),
    acc(
        "account_tag_income_other", "Liiketoiminnan muut tuotot", pl("income_other"), H
    ),
    grp(
        "materials_and_services",
        "Materiaalit ja palvelut",
        "materials_supplies_and_goods+account_tag_external_services",
        H,
    ),
    grp(
        "materials_supplies_and_goods",
        "Aineet, tarvikkeet ja tavarat",
        "account_tag_purchases+account_tag_change_in_inventories",
        HI,
    ),
    acc(
        "account_tag_purchases",
        "Ostot tilikauden aikana",
        pl("purchases", "purchases_adjustment_items"),
        DD,
    ),
    acc(
        "account_tag_change_in_inventories",
        "Varastojen muutos",
        pl("change_in_inventories"),
        DD,
    ),
    acc(
        "account_tag_external_services",
        "Ulkopuoliset palvelut",
        pl("external_services"),
        HI,
    ),
    grp(
        "personnel_expenses",
        "Henkilöstökulut",
        "account_tag_wages_and_salaries_in_production+indirect_personnel_expenses",
        H,
    ),
    acc(
        "account_tag_wages_and_salaries_in_production",
        "Palkat ja palkkiot",
        pl("wages_and_salaries_in_production", "wages_and_salaries_other"),
        HI,
    ),
    grp(
        "indirect_personnel_expenses",
        "Henkilösivukulut",
        "account_tag_pension_costs_of_production+account_tag_other_additional_expenses",
        HI,
    ),
    acc(
        "account_tag_pension_costs_of_production",
        "Eläkekulut",
        pl("pension_costs_of_production", "pension_costs_other"),
        DD,
    ),
    acc(
        "account_tag_other_additional_expenses",
        "Muut henkilösivukulut",
        pl("other_additional_expenses_other", "other_additional_expenses_production"),
        DD,
    ),
    grp(
        "depreciations_and_impairments",
        "Poistot ja arvonalentumiset",
        "account_tag_depreciation+account_tag_impairment+extra_ordinary_impairments_current_assets",
        H,
    ),
    acc(
        "account_tag_depreciation",
        "Suunnitelman mukaiset poistot",
        pl("depreciation_according_to_plan"),
        D,
    ),
    acc(
        "account_tag_impairment",
        "Arvonalentumiset pysyvien vastaavien hyödykkeistä",
        pl("impairment_of_fixed_assets"),
        D,
    ),
    acc(
        "extra_ordinary_impairments_current_assets",
        "Vaihtuvien vastaavien poikkeukselliset arvonalentumiset",
        pl("current_assets_extra_ordinary_deprecation"),
        D,
    ),
    acc(
        "account_tag_expenses_other",
        "Liiketoiminnan muut kulut",
        pl(
            "personnel_other",
            "premises_costs",
            "vehicle_expenses",
            "it_expenses",
            "in_machinery_and_equipment_expenses",
            "travelling",
            "representation",
            "selling_expenses",
            "marketing_expenses",
            "research_and_development",
            "administrative_services",
            "other_administrative_expenses",
            "other_fixed_costs",
        ),
        H,
    ),
    grp(
        "profit_loss",
        "LIIKEVOITTO (-TAPPIO)",
        "revenue+account_tag_manufacturing_warehouse_change+account_tag_manufacturing_own_use"
        "+account_tag_income_other+materials_and_services+personnel_expenses+depreciations_and_impairments"
        "+account_tag_expenses_other",
        H,
    ),
    grp(
        "finance_income_and_expenses",
        "Rahoitustuotot ja -kulut",
        "account_tag_income_from_shares_in_group+account_tag_income_from_associated_company"
        "+account_tag_investment_income_other_current_assets_group+account_tag_other_interest_and_financial_income"
        "+account_tag_impairment_losses+impairment_losses_financial_assets_current_assets"
        "+account_tag_financial_expenses",
        H,
    ),
    acc(
        "account_tag_income_from_shares_in_group",
        "Tuotot osuuksista saman konsernin yrityksissä",
        pl("income_from_shares_in_group"),
        D,
    ),
    acc(
        "account_tag_income_from_associated_company",
        "Tuotot osuuksista omistusyhteysyrityksissä",
        pl("income_from_associated_company"),
        D,
    ),
    grp(
        "account_tag_investment_income_other_current_assets_group",
        "Tuotot muista pysyvien vastaavien sijoituksista",
        "investment_income_other_current_assets_group+investment_income_other_current_assets_other",
        D,
    ),
    acc(
        "investment_income_other_current_assets_group",
        "Saman konsernin yrityksiltä",
        pl("investment_income_other_current_assets_group"),
        DD,
    ),
    acc(
        "investment_income_other_current_assets_other",
        "Muilta",
        pl("other_investment_income_from_fixed_assets"),
        DD,
    ),
    grp(
        "account_tag_other_interest_and_financial_income",
        "Muut korko- ja rahoitustuotot",
        "interest_and_financial_income_group+interest_and_financial_income_other",
        D,
    ),
    acc(
        "interest_and_financial_income_group",
        "Saman konsernin yrityksiltä",
        pl("interest_and_financial_income_from_group"),
        DD,
    ),
    acc(
        "interest_and_financial_income_other",
        "Muilta",
        pl("other_interest_and_financial_income"),
        DD,
    ),
    acc(
        "account_tag_impairment_losses",
        "Arvonalentumiset pysyvien vastaavien sijoituksista",
        pl("impairment_losses_on_investments_in_fixed_assets"),
        D,
    ),
    acc(
        "impairment_losses_financial_assets_current_assets",
        "Arvonalentumiset vaihtuvien vastaavien rahoitusarvopapereista",
        pl("impairment_losses_financial_assets_current_assets"),
        D,
    ),
    grp(
        "account_tag_financial_expenses",
        "Korkokulut ja muut rahoituskulut",
        "financial_expenses_group+financial_expenses_other",
        D,
    ),
    acc(
        "financial_expenses_group",
        "Saman konsernin yrityksille",
        pl("financial_expenses_group"),
        DD,
    ),
    acc("financial_expenses_other", "Muille", pl("other_financial_expenses"), DD),
    grp(
        "profit_loss_before_taxes",
        "VOITTO (TAPPIO) ENNEN TILINPÄÄTÖSSIIRTOJA JA VEROJA",
        "profit_loss+finance_income_and_expenses",
        H,
    ),
    grp(
        "appropriations",
        "Tilinpäätössiirrot",
        "account_tag_change_in_depreciation_difference+account_tag_change_in_provisions+account_tag_group_contribution",
        H,
    ),
    acc(
        "account_tag_change_in_depreciation_difference",
        "Poistoeron muutos",
        pl("change_in_depreciation_difference"),
        D,
    ),
    acc(
        "account_tag_change_in_provisions",
        "Vapaaehtoisten varausten muutos",
        pl("change_in_provisions"),
        D,
    ),
    acc(
        "account_tag_group_contribution", "Konserniavustus", pl("group_contribution"), D
    ),
    acc(
        "account_tag_taxes_for_the_period",
        "Tuloverot",
        pl("taxes_for_the_period", "deferred_taxes"),
        H,
    ),
    acc("account_tag_other_tax", "Muut välittömät verot", pl("other_tax"), H),
    acc("account_tag_minority_shares", "Vähemmistöosuudet", pl("minority_shares"), D),
    grp(
        "profit_loss_after_taxes",
        "TILIKAUDEN VOITTO (TAPPIO)",
        "profit_loss_before_taxes+appropriations+account_tag_taxes_for_the_period+account_tag_other_tax"
        "+account_tag_minority_shares",
        H,
    ),
]

STYLES = {
    H: "mis_template_financial_report.style_header",
    HI: "mis_template_financial_report.style_header_indent",
    HDI: "style_header_double_indent",
    D: "mis_template_financial_report.style_details",
    DD: "mis_template_financial_report.style_details_double_indent",
    DTI: "style_details_triple_indent",
}

HEADER = """<?xml version="1.0" encoding="utf-8" ?>
<!-- Generated from "{src}". Account tags: {mod}. -->
<odoo>
"""


def check(rows):
    acc_re = re.compile(
        r"\b(?:bal|pbal|nbal|crd|deb|fld)[piseu]?\s*(\[.*?\])\s*(\[.*?\])?"
    )
    for name, _d, expr, *_ in rows:
        for mo in acc_re.finditer(expr):
            dom = mo.group(1)
            # what mis_builder will see must be a complete domain
            compile(dom, name, "eval")
            assert dom.count("(") == dom.count(")"), (name, dom)
    names = [r[0] for r in rows]
    dup = {n for n in names if names.count(n) > 1}
    assert not dup, dup
    for name, _d, expr, *_ in rows:
        assert re.fullmatch(r"[a-z_][a-z0-9_]*", name), name
        if "[" not in expr:
            for ref in re.findall(r"[a-z_][a-z0-9_]*", expr):
                assert ref in names, (name, ref)


def kpis(prefix, report, rows):
    out = []
    for seq, (name, desc, expr, style) in enumerate(rows, start=1):
        out.append(f"""    <record id="kpi_{prefix}_{name}" model="mis.report.kpi">
        <field name="report_id" ref="{report}" />
        <field name="sequence">{seq * 10}</field>
        <field name="name">{name}</field>
        <field name="description">{escape(desc)}</field>
        <field name="expression">{escape(expr)}</field>
        <field name="style_id" ref="{STYLES[style]}" />
        <field name="auto_expand_accounts">false</field>
        <field name="type">num</field>
        <field name="compare_method">diff</field>
        <field name="accumulation_method">sum</field>
    </record>
""")
    return "".join(out)


def excel_desc(path):
    descs = {}
    for row in openpyxl.load_workbook(path).active.iter_rows(
        min_row=2, values_only=True
    ):
        if row[3]:
            descs.setdefault(row[3], row[4])  # first wins: 'assets' is duplicated
    return descs


# Corrections to Excel descriptions (typo / renamed KPIs keep own text)
DESC_FIX = {"shares_in_consolidated_securities": "Osuudet saman konsernin yrityksissä"}


def apply_desc(rows, descs):
    res = []
    for name, desc, *rest in rows:
        if name in DESC_FIX:
            desc = DESC_FIX[name]
        elif name in descs and name != "financial_securities":
            desc = descs[name]
        res.append((name, desc, *rest))
    return res


BS = apply_desc(BS, excel_desc(ARGS.bs_xlsx))
PL = apply_desc(PL, excel_desc(ARGS.pl_xlsx))
check(BS)
check(PL)

bs = (
    HEADER.format(src="Tase päivitetty v2.xlsx", mod=MOD)
    + """    <record id="style_header_double_indent" model="mis.report.style">
        <field name="name">PL/BS double indented header (FI)</field>
        <field name="font_weight_inherit">false</field>
        <field name="font_weight">bold</field>
        <field name="indent_level_inherit">false</field>
        <field name="indent_level">2</field>
    </record>
    <record id="style_details_triple_indent" model="mis.report.style">
        <field name="name">PL/BS triple indented details (FI)</field>
        <field name="indent_level_inherit">false</field>
        <field name="indent_level">3</field>
    </record>
    <record id="report_fi_bs" model="mis.report">
        <field name="name">Tase</field>
        <field name="style_id" ref="mis_template_financial_report.style_default" />
    </record>
"""
    + kpis("fi_bs", "report_fi_bs", BS)
    + "</odoo>\n"
)

pl_xml = (
    HEADER.format(src="Tulos päivitetty v2.xlsx", mod=MOD)
    + """    <record id="report_fi_pl" model="mis.report">
        <field name="name">Tuloslaskelma</field>
        <field name="style_id" ref="mis_template_financial_report.style_default" />
    </record>
"""
    + kpis("fi_pl", "report_fi_pl", PL)
    + "</odoo>\n"
)

open(f"{OUT}/mis_report_fi_bs.xml", "w", encoding="utf-8").write(bs)
open(f"{OUT}/mis_report_fi_pl.xml", "w", encoding="utf-8").write(pl_xml)
sys.stdout.write(f"BS kpis {len(BS)} PL kpis {len(PL)}\n")
