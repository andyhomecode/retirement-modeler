"""Default tax-law tables written into a new workbook's tax_tables tab.

After the workbook exists, its tax_tables tab is what counts: the spreadsheet formulas and the Python model
(model.py, via workbook.py) both read the tab, so to update the law (or switch state/province) edit the tab,
not this file. This file only seeds new workbooks.

Amounts are for TAX_YEAR and grow with inflation in the model where the law indexes them.

US federal: 2026 law (One Big Beautiful Bill Act, indexed by Rev. Proc. 2025-32).
Canada: 2026 federal and British Columbia amounts (CRA 2026 indexation, 2.0%; BC Budget 2026).
"""
TAX_YEAR = 2026

# ---------------------------------------------------------------------------------------------- United States
US_SCALARS = [
    ("TAX_YEAR", TAX_YEAR, "Year the amounts on this tab are for; indexed amounts grow with inflation after it"),
    ("FED_STD", 32200, "Standard deduction, married filing jointly"),
    ("FED_65_ADD", 1650, "Extra standard deduction per spouse 65+"),
    ("SENIOR_DED", 6000, "OBBBA senior deduction per person 65+ (not indexed)"),
    ("SENIOR_DED_FROM", 2025, "... first year"),
    ("SENIOR_DED_TO", 2028, "... last year (expires after 2028 under current law)"),
    ("SENIOR_PHASEOUT_START", 150000, "... reduced by 6% of MAGI above this (joint)"),
    ("SENIOR_PHASEOUT_RATE", 0.06, ""),
    ("NIIT_RATE", 0.038, "Net investment income tax"),
    ("NIIT_THRESHOLD", 250000, "Not indexed"),
    ("SS_TAXABLE_BASE", 32000, "Social Security taxation thresholds (joint, not indexed)"),
    ("SS_TAXABLE_ADJUSTED_BASE", 44000, ""),
    ("SALT_CAP_AFTER", 10000, "SALT cap after the years in the SALT table"),
    ("SALT_PHASEDOWN_START", 505000, "SALT cap phase-down starts at this MAGI (2026; grows 1%/yr through 2029)"),
    ("SALT_PHASEDOWN_RATE", 0.30, ""),
    ("UNRECAPTURED_1250_EXTRA", 0.10, "Depreciation recapture: 25% max rate vs 15%"),
    ("HOME_SALE_EXCLUSION", 500000, "Gain excluded on selling your home (joint)"),
    ("SS_WAGE_BASE", 184500, "Social Security payroll tax wage base"),
    ("ADDITIONAL_MEDICARE_THRESHOLD", 250000, "0.9% extra Medicare tax above this (joint, not indexed)"),
    ("MEDICAL_FLOOR", 0.075, "Medical expenses deductible above this share of AGI"),
    ("SS_FULL_RETIREMENT_AGE", 67, "Born 1960 or later"),
    ("STATE_NAME", "New York", "State used for state income tax (edit the brackets below to change it)"),
    ("STATE_STD", 16050, "State standard deduction (joint)"),
    ("STATE_INDEXED", 0, "1 = state brackets and deduction grow with inflation; 0 = fixed"),
    ("STATE_SS_EXEMPT", 1, "1 = state doesn't tax Social Security"),
    ("STATE_RET_EXCLUSION", 20000, "Per-person exclusion for IRA/401(k)/pension withdrawals"),
    ("STATE_RET_EXCLUSION_AGE", 59.5, "... from this age"),
    ("LOCAL_RATE", 0.0, "Flat city/county income tax on state taxable income (NYC is roughly 0.039)"),
]
US_BRACKETS = {
    "fed": ("Federal ordinary brackets (joint)", True,
            [(0, 0.10), (24800, 0.12), (100800, 0.22), (211400, 0.24), (403550, 0.32), (512450, 0.35), (768700, 0.37)]),
    "ltcg": ("Federal capital gains / qualified dividends (joint)", True, [(0, 0.0), (98900, 0.15), (613700, 0.20)]),
    # New York State (joint). NY's high-income "benefit recapture" isn't modeled.
    "state": ("State brackets (joint)", None,
              [(0, 0.04), (17150, 0.045), (23600, 0.0525), (27900, 0.055), (161550, 0.06), (323200, 0.0685),
               (2155350, 0.0965), (5000000, 0.103), (25000000, 0.109)]),
}
US_IRMAA = [  # MAGI over (joint), Part B add-on, Part D add-on ($/month per person)
    (218000, 81.20, 14.50), (274000, 202.90, 37.50), (342000, 324.60, 60.40), (410000, 446.30, 83.30),
    (750000, 487.00, 91.00)]
US_SALT = [(2026, 40400), (2027, 40804), (2028, 41212), (2029, 41624)]
UNIFORM_LIFETIME = {72: 27.4, 73: 26.5, 74: 25.5, 75: 24.6, 76: 23.7, 77: 22.9, 78: 22.0, 79: 21.1, 80: 20.2,
                    81: 19.4, 82: 18.5, 83: 17.7, 84: 16.8, 85: 16.0, 86: 15.2, 87: 14.4, 88: 13.7, 89: 12.9,
                    90: 12.2, 91: 11.5, 92: 10.8, 93: 10.1, 94: 9.5, 95: 8.9, 96: 8.4, 97: 7.8, 98: 7.3,
                    99: 6.8, 100: 6.4, 101: 6.0, 102: 5.6, 103: 5.2, 104: 4.9, 105: 4.6, 106: 4.3, 107: 4.1,
                    108: 3.9, 109: 3.7, 110: 3.5, 111: 3.4, 112: 3.3, 113: 3.1, 114: 3.0, 115: 2.9, 116: 2.8,
                    117: 2.7, 118: 2.5, 119: 2.3, 120: 2.0}

# ---------------------------------------------------------------------------------------------- Canada (BC)
CA_SCALARS = [
    ("TAX_YEAR", TAX_YEAR, "Year the amounts on this tab are for; indexed amounts grow with inflation after it"),
    ("FED_CREDIT_RATE", 0.14, "Federal non-refundable credits are worth this rate (lowest bracket, 2026+)"),
    ("FED_BPA_MAX", 16452, "Federal basic personal amount (full)"),
    ("FED_BPA_MIN", 14829, "... reduced to this at high income"),
    ("FED_BPA_PHASE_START", 181440, "... reduction starts at this net income"),
    ("FED_BPA_PHASE_END", 258482, "... and is complete here"),
    ("FED_AGE_AMOUNT", 9208, "Age amount (65+)"),
    ("FED_AGE_THRESHOLD", 46432, "... reduced by 15% of net income above this"),
    ("AGE_REDUCTION_RATE", 0.15, ""),
    ("FED_PENSION_AMOUNT", 2000, "Pension income amount (not indexed)"),
    ("FED_EMPLOYMENT_AMOUNT", 1501, "Canada employment amount"),
    ("FED_MEDICAL_THRESHOLD", 2891, "Medical expenses above the lesser of 3% of net income and this"),
    ("FED_ELIGIBLE_DTC", 0.150198, "Federal dividend tax credit, eligible dividends (share of grossed-up amount)"),
    ("ELIGIBLE_GROSSUP", 0.38, "Eligible dividend gross-up"),
    ("CG_INCLUSION", 0.5, "Share of capital gains that is taxable"),
    ("OAS_MONTHLY", 742.31, "Full OAS pension at 65-74, per month (Jan 2026)"),
    ("OAS_75_BOOST", 0.10, "OAS increase from age 75"),
    ("OAS_CLAWBACK_THRESHOLD", 95323, "OAS recovery tax starts at this net income"),
    ("OAS_CLAWBACK_RATE", 0.15, ""),
    ("CPP_MAX_MONTHLY", 1507.65, "Maximum CPP retirement pension at 65 (2026); caps survivor + own pension"),
    ("CPP_SURVIVOR_SHARE", 0.6, "Survivor pension (65+) as a share of the deceased's pension"),
    ("CPP_RATE", 0.0595, "Employee CPP contribution rate (base + first additional)"),
    ("CPP_EXEMPTION", 3500, "Basic exemption (not indexed)"),
    ("YMPE", 74600, "Year's maximum pensionable earnings"),
    ("CPP2_RATE", 0.04, "Second additional CPP rate"),
    ("YAMPE", 85000, "Year's additional maximum pensionable earnings"),
    ("EI_RATE", 0.0163, "Employee EI premium rate"),
    ("EI_MAX_INSURABLE", 68900, ""),
    ("TFSA_LIMIT", 7000, "New TFSA room per person per year"),
    ("PROV_NAME", "British Columbia", "Province used for provincial tax (edit the brackets and amounts to change it)"),
    ("PROV_CREDIT_RATE", 0.056, "Provincial credits are worth this rate (BC lowest rate, 5.6% from 2026)"),
    ("PROV_BPA", 13216, "Provincial basic personal amount"),
    ("PROV_AGE_AMOUNT", 5927, "Provincial age amount (65+)"),
    ("PROV_AGE_THRESHOLD", 44119, "... reduced by 15% of net income above this"),
    ("PROV_PENSION_AMOUNT", 1000, "Provincial pension income amount (not indexed)"),
    ("PROV_MEDICAL_THRESHOLD", 2748, "Provincial medical expense threshold"),
    ("PROV_ELIGIBLE_DTC", 0.12, "Provincial dividend tax credit, eligible dividends"),
    ("PROV_REDUCTION_MAX", 690, "BC tax reduction (low income)"),
    ("PROV_REDUCTION_THRESHOLD", 25570, "... reduced by the rate below above this net income"),
    ("PROV_REDUCTION_RATE", 0.0356, ""),
    ("PROV_INDEX_PAUSE_FROM", 2027, "BC isn't indexing its brackets and amounts from this year ..."),
    ("PROV_INDEX_PAUSE_TO", 2030, "... through this year (BC Budget 2026); indexing resumes after"),
    ("RRIF_START_AGE", 72, "Minimum RRIF withdrawals start the year you turn this (RRSP must convert by the end of 71)"),
]
CA_BRACKETS = {
    "fed": ("Federal brackets", True,
            [(0, 0.14), (58523, 0.205), (117045, 0.26), (181440, 0.29), (258482, 0.33)]),
    "prov": ("Provincial brackets", True,
             [(0, 0.056), (50363, 0.077), (100728, 0.105), (115648, 0.1229), (140430, 0.147), (190405, 0.168),
              (265545, 0.205)]),
}
# RRIF minimum withdrawal by age on January 1 (under 71: 1 / (90 - age))
RRIF_FACTORS = {71: 0.0528, 72: 0.0540, 73: 0.0553, 74: 0.0567, 75: 0.0582, 76: 0.0598, 77: 0.0617, 78: 0.0636,
                79: 0.0658, 80: 0.0682, 81: 0.0708, 82: 0.0738, 83: 0.0771, 84: 0.0808, 85: 0.0851, 86: 0.0899,
                87: 0.0955, 88: 0.1021, 89: 0.1099, 90: 0.1192, 91: 0.1306, 92: 0.1449, 93: 0.1634, 94: 0.1879}
RRIF_FACTORS.update({a: 0.20 for a in range(95, 121)})

BRACKET_ROWS = 12   # fixed room per bracket table so a state/province with more brackets fits


def rmd_start_age(born):
    """US required minimum distributions (SECURE 2.0): 72 if born before 1951, 73 for 1951-59, 75 from 1960."""
    return 72 if born < 1951 else (73 if born < 1960 else 75)
