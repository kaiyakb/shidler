# Microsoft Corporation FY2024 – Accounting & Performance Ratio Analysis

*Technical Specification | Stage 3*

|                |                                                         |
|----------------|---------------------------------------------------------|
|**Created by**  |Kaiya Baisa                                              |
|**Date Created**|April 2026                                               |
|**Version**     |1.0                                                      |
|**Role**        |Financial Analyst / FP&A Analyst                         |
|**Audience**    |CFO / Director of FP&A                                   |
|**Data Sources**|Microsoft FY2024 Annual Report (SEC EDGAR); Yahoo Finance|

-----

## 1. Problem Statement

Microsoft Corporation is a publicly traded multinational technology company running in the software and cloud computing industry with major product lines including Microsoft Azure, Microsoft Office, and Windows. This outlines the analytical framework for computing 25+ accounting and performance ratios from Microsoft’s FY2024 financial statements (fiscal year ending June 30, 2024), enabling management to assess financial health, operational efficiency, leverage, liquidity, and value creation.

The primary goal of this analysis is to convert Microsoft’s raw financial statement data into actionable performance indicators that allow the CFO to evaluate trends over time, benchmark against competitors, and name strengths or weaknesses requiring strategic attention.

-----

## 2. Inputs (Known Variables)

### Balance Sheet Items

|**Variable**                |**Named Range**                      |**Description**                   |**Year**       |**Value ($M)**   |
|----------------------------|-------------------------------------|----------------------------------|---------------|-----------------|
|Cash & marketable securities|BAL_cash_marketable_securities_[year]|Liquid assets                     |FY2024 / FY2023|75,543 / 111,262 |
|Receivables                 |BAL_receivables_[year]               |Accounts receivable               |FY2024 / FY2023|56,924 / 48,688  |
|Inventories                 |BAL_inventories_[year]               |Inventory balance                 |FY2024 / FY2023|1,246 / 2,500    |
|Total current assets        |BAL_assets_current_[year]            |Sum of current assets             |FY2024 / FY2023|159,734 / 184,257|
|Net tangible fixed assets   |BAL_fixed_assets_net_[year]          |PP&E less accumulated depreciation|FY2024         |135,591          |
|Total assets                |BAL_assets_total_[year]              |All assets                        |FY2024 / FY2023|512,163 / 411,976|
|Total current liabilities   |BAL_liabilities_current_[year]       |Short-term obligations            |FY2024         |125,286          |
|Long-term debt              |BAL_debt_long_term_[year]            |Non-current borrowings            |FY2024 / FY2023|42,688 / 41,990  |
|Total liabilities           |BAL_liabilities_total_[year]         |All liabilities                   |FY2024         |243,286          |
|Shareholders’ equity        |BAL_equity_shareholders_[year]       |Book value of equity              |FY2024 / FY2023|268,877 / 206,223|

### Income Statement Items

|**Variable**      |**Named Range**     |**Description**                   |**FY2024 ($M)**|
|------------------|--------------------|----------------------------------|---------------|
|Net sales         |INC_sales           |Total revenue                     |245,122        |
|Cost of goods sold|INC_cost_goods_sold |Direct costs                      |59,742         |
|SG&A expenses     |INC_sga             |Operating expenses                |61,575         |
|Depreciation      |INC_depreciation    |Non-cash expense                  |14,372         |
|EBIT              |INC_ebit            |Earnings before interest and taxes|109,433        |
|Other income      |INC_other_income    |Non-operating income/(loss)       |(1,646)        |
|Interest expense  |INC_interest_expense|Cost of debt                      |1,526          |
|Taxes             |INC_taxes           |Income tax expense                |19,651         |
|Net income        |INC_net             |Bottom line                       |86,610         |
|Dividends         |INC_dividends       |Shareholder distributions         |22,351         |

### Cash Flow Statement Items

|**Variable**         |**Named Range** |**Description**    |**FY2024 ($M)**|
|---------------------|----------------|-------------------|---------------|
|Cash from operations |CASH_operating  |Operating cash flow|117,022        |
|Cash from investments|CASH_investments|Investing cash flow|(96,970)       |

### Market / Analyst Inputs

|**Variable**      |**Named Range**   |**Description**                         |**Value**|
|------------------|------------------|----------------------------------------|---------|
|Share price       |share_price       |Market price per share (Yahoo Finance)  |$446.34  |
|Shares outstanding|shares_outstanding|Total shares (millions)                 |7,431    |
|Cost of capital   |cost_capital      |WACC / required return                  |8.9%     |
|Tax rate          |tax_rate          |Effective rate from financial statements|18.3%    |

-----

## 3. Assumptions & Constraints

- All numbers are in millions USD
- Fiscal year ends June 30, 2024 (FY2024); prior year is FY2023 (ending June 30, 2023).
- Tax rate uses the effective rate of 18.3% derived directly from the FY2024 income statement (taxes of $19,651M on taxable income of $106,261M).
- Cost of capital is set at 8.9%, from Yahoo Finance analyst estimates for Microsoft’s WACC.
- Depreciation figure is taken from the Income Statement ($14,372M). The cash flow statement includes a broader depreciation and amortization figure of $19,723M; the income statement figure is used for ratio calculations per standard practice.
- Start-of-year (FY2023) balance sheet values are used as denominators for ROA, ROC, ROE, and turnover ratios where indicated, with average-denominator variants calculated separately.
- Total capitalization is defined as long-term debt plus shareholders’ equity.
- No off-balance-sheet items, contingent liabilities, or minority interests are included.
- Market capitalization is computed as share price multiplied by shares outstanding: $446.34 × 7,431M = $3,316,752.54M.
- Interest rates and coverage ratios are made on a simple annual basis.

-----

## 4. Calculation Flow

The following steps describe the logic and sequencing of the Excel model, using named-range pseudocode consistent with the Stage 2 workbook.

### Step 1: Derived Inputs

- `market_capitalization` = share_price × shares_outstanding → **$3,316,752.54M**
- `currentYear_after_tax_operating_income` = INC_net + (1 − tax_rate) × INC_interest_expense → **$87,856.74M**
- `currentYear_daily_sales_average` = INC_sales / 365 → **$671.57M/day**
- `currentYear_cost_goods_sold_daily` = INC_cost_goods_sold / 365 → **$163.68M/day**
- `currentYear_working_capital_net` = BAL_assets_current_2024 − BAL_liabilities_current_2024 → **$34,448M**
- `startYear_total_capitalization` = BAL_debt_long_term_2023 + BAL_equity_shareholders_2023 → **$248,213M**
- `avg_equity` = AVERAGE(startYear_equity, currentYear_equity) → **$237,550M**
- `avg_total_assets` = AVERAGE(startYear_total_assets, currentYear_assets_total) → **$462,069.5M**
- `avg_total_capitalization` = AVERAGE(startYear_total_capitalization, currentYear_total_capitalization) → **$279,889M**

### Step 2: Performance Ratios

- **MVA** (Market Value Added) = market_capitalization − currentYear_equity → **$3,047,875.54M**
- **Market-to-Book Ratio** = market_capitalization / currentYear_equity → **12.34x**
- **EVA** (Economic Value Added) = currentYear_after_tax_operating_income − (cost_capital × startYear_total_capitalization) → **$65,765.79M**

### Step 3: Profitability Ratios

- **ROA** = currentYear_after_tax_operating_income / startYear_total_assets → **21.33%**
- **ROC** = currentYear_after_tax_operating_income / startYear_total_capitalization → **35.40%**
- **ROE** = INC_net / startYear_equity → **42.00%**
- **ROA [avg]** = currentYear_after_tax_operating_income / avg_total_assets → **19.01%**
- **ROC [avg]** = currentYear_after_tax_operating_income / avg_total_capitalization → **31.39%**
- **ROE [avg]** = INC_net / avg_equity → **36.46%**

### Step 4: Efficiency Ratios

- **Asset Turnover** = INC_sales / startYear_total_assets → **0.595** *(named: RATIO_asset_turnover)*
- **Receivables Turnover** = INC_sales / startYear_receivables → **5.03x**
- **Average Collection Period** = startYear_receivables / currentYear_daily_sales_average → **72.5 days**
- **Inventory Turnover** = INC_cost_goods_sold / startYear_inventory → **23.90x**
- **Days in Inventory** = startYear_inventory / currentYear_cost_goods_sold_daily → **15.3 days**
- **Profit Margin** = INC_net / INC_sales → **35.33%**
- **Operating Profit Margin** = currentYear_after_tax_operating_income / INC_sales → **35.84%** *(named: RATIO_operating_profit_margin)*

### Step 5: Leverage Ratios

- **Long-term Debt Ratio** = currentYear_debt_long_term / (currentYear_debt_long_term + currentYear_equity) → **13.70%**
- **Long-term Debt-Equity Ratio** = currentYear_debt_long_term / currentYear_equity → **15.88%**
- **Total Debt Ratio** = currentYear_liabilities_total / currentYear_assets_total → **47.50%**
- **Times Interest Earned** = INC_ebit / INC_interest_expense → **71.71x**
- **Cash Coverage Ratio** = (INC_ebit + INC_depreciation) / INC_interest_expense → **81.13x**
- **Debt Burden** = INC_net / currentYear_after_tax_operating_income → **0.9858** *(named: RATIO_debt_burden)*
- **Leverage Ratio** = currentYear_assets_total / currentYear_equity → **1.905x** *(named: RATIO_leverage)*

### Step 6: Liquidity Ratios

- **NWC-to-Assets** = currentYear_working_capital_net / currentYear_assets_total → **6.73%**
- **Current Ratio** = currentYear_assets_current / currentYear_liabilities_current → **1.275x**
- **Quick Ratio** = (currentYear_cash_marketable_securities + BAL_receivables_2024) / currentYear_liabilities_current → **1.057x**
- **Cash Ratio** = currentYear_cash_marketable_securities / currentYear_liabilities_current → **0.603x**

### Step 7: Du Pont Decomposition

- **Du Pont ROA** = RATIO_asset_turnover × RATIO_operating_profit_margin → **21.33%**
- **Du Pont ROE** = RATIO_leverage × RATIO_asset_turnover × RATIO_operating_profit_margin × RATIO_debt_burden → **40.05%**

-----

## 5. Outputs

|**Output**           |**Description**                                                                                   |**Format**    |**Purpose**                                   |
|---------------------|--------------------------------------------------------------------------------------------------|--------------|----------------------------------------------|
|Ratio summary table  |All 25+ ratios organized by category (Performance, Profitability, Efficiency, Leverage, Liquidity)|Table         |Core analytical output                        |
|Du Pont decomposition|ROA and ROE breakdown into component drivers                                                      |Table         |Identifies what is driving Microsoft’s returns|
|Formula documentation|Named-range formula for each ratio as a dedicated column in the Ratios sheet                      |Column        |Auditability and reproducibility              |
|Executive summary    |Key findings, ratio interpretation, and strategic recommendations                                 |1–2 paragraphs|Stage 5 final memo input                      |

-----

## 6. Model Review — What Worked & What to Improve

### What Worked Well

- The named range system was highly effective. Every input variable was assigned a descriptive named range (e.g., `BAL_equity_shareholders_2024`, `INC_net`), making all ratio formulas self-documenting and easy to audit.
- Separating the Ratios sheet into a ‘Start of Year,’ ‘Current Year,’ and ‘Mixed Year’ derived inputs section before the ratio outputs created a clean, logical flow that mirrors professional FP&A practice.
- The Balance Sheet check column (which confirmed total assets equaled total liabilities and equity, verifying a zero difference) provided immediate data integrity validation.
- Having both start-of-year and average-denominator variants for ROA, ROC, and ROE is analytically rigorous and gives the CFO multiple perspectives on return performance.

### What to Improve

- There is an error in `currentYear_total_capitalization`: the named range formula references `BAL_debt_long_term_2020` and `BAL_equity_shareholders_2020` (FY2020 data) instead of FY2024 values. This must be corrected to `BAL_debt_long_term_2024 + BAL_equity_shareholders_2024`.
- Similarly, the Quick Ratio formula references `BAL_receivables_2020` instead of `BAL_receivables_2024`, which understates the quick ratio. This should be corrected in the refined model.
- A color-coding convention consistent with industry standards would make the model look neater and cleaner.

### What Would Make the Model More Auditable

- Adding a dedicated ‘Inputs’ tab that isolates all raw financial statement data (balance sheet, income statement, cash flow) from the calculation logic would create a cleaner model architecture.
- A ratio interpretation column alongside each calculated ratio (e.g., ‘Above 1.0 indicates adequate liquidity’) would make the model more useful for non-technical stakeholders.

### Additional Analysis Worth Including

- Multi-year trend analysis (FY2022–FY2024) to show whether Microsoft’s profitability and efficiency are improving or plateauing.
- Industry peer comparison (e.g., vs. Alphabet or Apple) to contextualize the absolute ratio values.

-----

## 7. Limitations & Next Steps

This specification does not incorporate industry peer comparisons, multi-year trend analysis, or off-balance-sheet items. The quick ratio and total capitalization hold known referencing errors from Stage 2 that must be corrected before the model is considered final. Additionally, the analysis reflects a single fiscal year (FY2024) and cannot by itself indicate whether trends are improving or deteriorating.

With Stage 4, it will involve writing a structured AI prompt using this specification as its machine-readable input. That prompt will instruct an AI model to regenerate the corrected Excel workbook and produce a polished ratio summary. Stage 5 will deliver a final executive memo to the CFO interpreting the ratio results and providing strategic recommendations for Microsoft management.