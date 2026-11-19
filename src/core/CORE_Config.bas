Attribute VB_Name = "CORE_Config"
'==============================================================================
' MODULE: CORE_Config
'------------------------------------------------------------------------------
' PURPOSE
'   Hold every constant the K-FRTB-SSA engine depends on: sheet names, the
'   workbook names that locate each table, the column of each input field,
'   the width of each output table, the parameter codes with their indices
'   in the parameter vector and their MAR40 reference values, the default
'   MAR40 maturity ladder, row-status and severity labels, the Checks rule
'   identifiers and the error numbers. When a layout changes, it changes
'   here and nowhere else.
'
' PUBLIC SURFACE
'   None outside this VBA project. The constants are Public for in-project
'   use by the CORE_ modules, FRTBSSA_Formulas, M_Main and the test modules;
'   Option Private Module keeps them off the supported external surface.
'
' DEPENDENCIES
'   None. Every value is a literal.
'
' STATE OWNERSHIP
'   Constants only; no state.
'
' ERROR POLICY
'   Not applicable.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' UPDATED
'   2026-11-15
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the constants inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' PRODUCT
'------------------------------------------------------------------------------
        Public Const PRODUCT_NAME   As String = "K-FRTB-SSA"     'Name shown in messages
        Public Const ENGINE_VERSION As String = "0.2.0"     'Version of this engine

'------------------------------------------------------------------------------
' SHEET NAMES
'------------------------------------------------------------------------------
    'Worksheet tab names, as M_Main looks them up in ThisWorkbook.
        Public Const SH_PARAMS      As String = "Params"           'Parameters and tables
        Public Const SH_IR_POS      As String = "IRPositions"      'Interest rate inputs
        Public Const SH_EQ_POS      As String = "EQPositions"      'Equity inputs
        Public Const SH_CO_POS      As String = "COPositions"      'Commodity inputs
        Public Const SH_FX_POS      As String = "FXPositions"      'FX and gold inputs
        Public Const SH_OPTIONS     As String = "OptionCharges"    'Non-delta option charges
        Public Const SH_RESULTS     As String = "Results"          'Requirement per risk class
        Public Const SH_IR_ZONES    As String = "IR_Zones"         'IR charges per currency
        Public Const SH_IR_LADDER   As String = "IR_Ladder"        'IR ladder per currency and band
        Public Const SH_BUCKETS     As String = "Buckets"          'EQ markets, commodities, FX
        Public Const SH_IR_CALC     As String = "IRCalc"           'One row per IR input row
        Public Const SH_EQ_CALC     As String = "EQCalc"           'One row per EQ input row
        Public Const SH_CO_CALC     As String = "COCalc"           'One row per CO input row
        Public Const SH_FX_CALC     As String = "FXCalc"           'One row per FX input row
        Public Const SH_CHECKS      As String = "Checks"           'Rules and findings

'------------------------------------------------------------------------------
' TABLE LOCATION
'------------------------------------------------------------------------------
    'Every table is located through the workbook name of its first header
    'cell, in column C; its data start on the next row. Totals and single
    'outputs are located through their own names. InputRows can therefore move
    'in the template without a change here.
        Public Const NM_IR_POS_HEADER    As String = "IRPositions_Header"     'Header "PositionID"
        Public Const NM_EQ_POS_HEADER    As String = "EQPositions_Header"     'Header "PositionID"
        Public Const NM_CO_POS_HEADER    As String = "COPositions_Header"     'Header "PositionID"
        Public Const NM_FX_POS_HEADER    As String = "FXPositions_Header"     'Header "PositionID"
        Public Const NM_IR_CALC_HEADER   As String = "IRCalc_Header"          'Header "PositionID"
        Public Const NM_EQ_CALC_HEADER   As String = "EQCalc_Header"          'Header "PositionID"
        Public Const NM_CO_CALC_HEADER   As String = "COCalc_Header"          'Header "PositionID"
        Public Const NM_FX_CALC_HEADER   As String = "FXCalc_Header"          'Header "PositionID"
        Public Const NM_LADDER_HEADER    As String = "IR_Ladder_Header"       'Header "Ccy"
        Public Const NM_ZONES_HEADER     As String = "IR_Zones_Header"        'Header "Ccy"
        Public Const NM_BK_EQ_HEADER     As String = "Buckets_EQ_Header"      'Header "#", equity block
        Public Const NM_BK_CO_HEADER     As String = "Buckets_CO_Header"      'Header "#", commodity block
        Public Const NM_BK_FX_HEADER     As String = "Buckets_FX_Header"      'Header "#", FX block
        Public Const NM_RESULTS_HEADER   As String = "Results_Header"         'Header "Risk class"
        Public Const NM_RESULTS_STATUS   As String = "Results_Status"         'Run status line
        Public Const NM_FINGERPRINT      As String = "Results_Fingerprint"    'Fingerprint of the run's inputs
        Public Const NM_CHECKS_HEADER    As String = "Checks_Header"          'Header "Rule"
        Public Const NM_FINDINGS_HEADER  As String = "Checks_Findings_Header" 'Header "#", findings
        Public Const NM_CHECKS_ERRORS    As String = "Checks_Errors"          'Error count cell
        Public Const NM_CHECKS_WARNINGS  As String = "Checks_Warnings"        'Warning count cell
        Public Const NM_IR_GMR_TOTAL     As String = "IR_GMR_Total"           'IR_Zones total row, GMR
        Public Const NM_EQ_GENERAL       As String = "EQ_General"             'Buckets equity total row
        Public Const NM_CO_DIRECTIONAL   As String = "CO_Directional"         'Buckets commodity total row
        Public Const NM_FX_ONOP          As String = "FX_ONOP"                'Overall net open position
        Public Const NM_FX_CHARGE        As String = "FX_Charge"              'FX requirement
        Public Const NM_LADDER_ZONE      As String = "Ladder_Zone"            'Ladder: zone column
        Public Const NM_LADDER_UP_HIGH   As String = "Ladder_UpperHigh"       'Ladder: bound, coupon >= 3%
        Public Const NM_LADDER_UP_LOW    As String = "Ladder_UpperLow"        'Ladder: bound, coupon < 3%
        Public Const NM_LADDER_LBL_HIGH  As String = "Ladder_LabelHigh"       'Ladder: label, coupon >= 3%
        Public Const NM_LADDER_LBL_LOW   As String = "Ladder_LabelLow"        'Ladder: label, coupon < 3%
        Public Const NM_LADDER_RW        As String = "Ladder_RW"              'Ladder: risk weight
        Public Const NM_SR_KEY           As String = "SR_Key"                 'Specific-risk keys
        Public Const NM_SR_CHARGE        As String = "SR_Charge"              'Specific-risk charges
        Public Const NM_FX_CCY           As String = "FX_Ccy"                 'FX table currencies
        Public Const NM_FX_RATE          As String = "FX_Rate"                'FX table rates
        Public Const NM_OPTIONS_PREFIX   As String = "Options"                '+ IR, EQ, CO, FX

'------------------------------------------------------------------------------
' VOLUME
'------------------------------------------------------------------------------
    'Rows of a position sheet that the workbook supports, counted from the row
    'below the header to the last row with a PositionID: the template formats,
    'validates and colours this many rows, and the formula workbook calculates
    'them. A run on a sheet that holds more stops (rule C14).
        Public Const MAX_POSITION_ROWS As Long = 500

'------------------------------------------------------------------------------
' INPUT COLUMNS
'------------------------------------------------------------------------------
    'Column of each field, counted from the first table column (1 = C).
    'Interest rate positions (IRPositions).
        Public Const IR_ID           As Long = 1     'PositionID
        Public Const IR_INSTRUMENT   As Long = 3     'Bond, FRN, Swap leg, Future leg, FRA leg, Forward leg or Other
        Public Const IR_CCY          As Long = 4     'Currency
        Public Const IR_DIRECTION    As Long = 5     'Long or Short
        Public Const IR_VALUE        As Long = 6     'Position value, currency units
        Public Const IR_RATE_TYPE    As Long = 7     'Fixed or Floating
        Public Const IR_COUPON       As Long = 8     'Coupon, decimal
        Public Const IR_MATURITY     As Long = 9     'Final maturity date
        Public Const IR_REPRICING    As Long = 10    'Next repricing date, floating only
        Public Const IR_CATEGORY     As Long = 11    'GOV, QUAL, OTHER or NONE
        Public Const IR_GRADE        As Long = 12    'Rating grade
        Public Const IR_ISSUE        As Long = 13    'Issue ID for specific-risk netting
        Public Const IR_INPUT_WIDTH  As Long = 14    'Columns read
    'Equity positions (EQPositions).
        Public Const EQ_ID           As Long = 1     'PositionID
        Public Const EQ_MARKET       As Long = 3     'Market
        Public Const EQ_TYPE         As Long = 4     'Single name or Index
        Public Const EQ_IDENTIFIER   As Long = 5     'Stock or index identifier
        Public Const EQ_CCY          As Long = 6     'Currency
        Public Const EQ_DIRECTION    As Long = 7     'Long or Short
        Public Const EQ_VALUE        As Long = 8     'Market value, currency units
        Public Const EQ_INPUT_WIDTH  As Long = 9     'Columns read
    'Commodity positions (COPositions).
        Public Const CO_ID           As Long = 1     'PositionID
        Public Const CO_COMMODITY    As Long = 3     'Commodity name
        Public Const CO_DIRECTION    As Long = 4     'Long or Short
        Public Const CO_QUANTITY     As Long = 5     'Quantity, units of the commodity
        Public Const CO_PRICE        As Long = 7     'Spot price, currency per unit
        Public Const CO_CCY          As Long = 8     'Currency of the price
        Public Const CO_INPUT_WIDTH  As Long = 10    'Columns read
    'FX and gold positions (FXPositions).
        Public Const FX_ID           As Long = 1     'PositionID
        Public Const FX_CCY          As Long = 3     'Currency (gold: GoldCode)
        Public Const FX_DIRECTION    As Long = 4     'Long or Short
        Public Const FX_AMOUNT       As Long = 5     'Amount, currency units
        Public Const FX_INPUT_WIDTH  As Long = 6     'Columns read

'------------------------------------------------------------------------------
' OUTPUT WIDTHS
'------------------------------------------------------------------------------
    'Number of columns each core procedure returns, from column C (Results
    'from column E). They match the template headers.
        Public Const IR_CALC_WIDTH   As Long = 22    'IRCalc C:X
        Public Const EQ_CALC_WIDTH   As Long = 12    'EQCalc C:N
        Public Const CO_CALC_WIDTH   As Long = 8     'COCalc C:J
        Public Const FX_CALC_WIDTH   As Long = 7     'FXCalc C:I
        Public Const LADDER_WIDTH    As Long = 13    'IR_Ladder C:O
        Public Const ZONES_WIDTH     As Long = 30    'IR_Zones C:AF
        Public Const BK_EQ_WIDTH     As Long = 10    'Buckets equity block C:L
        Public Const BK_CO_WIDTH     As Long = 10    'Buckets commodity block C:L
        Public Const BK_FX_WIDTH     As Long = 9     'Buckets FX block C:K
        Public Const RESULTS_WIDTH   As Long = 11    'Results E:O
        Public Const RESULTS_COLUMN  As Long = 3     'Results data column (E) from the header (C)
        Public Const FINDINGS_WIDTH  As Long = 7     'Checks findings C:I

'------------------------------------------------------------------------------
' TOTAL VECTORS
'------------------------------------------------------------------------------
    'Index of each figure in the Totals() vector a risk-class procedure
    'returns. Amounts in reporting currency; counts as whole numbers.
        Public Const TOT_SPECIFIC    As Long = 1     'Specific risk (IR, EQ)
        Public Const TOT_GENERAL     As Long = 2     'General market risk / directional / FX charge
        Public Const TOT_OTHER       As Long = 3     'EQ index charge, CO gross charge
        Public Const TOT_USED        As Long = 4     'InputRows with status OK
        Public Const TOT_REJECTED    As Long = 5     'InputRows with status INVALID
        Public Const TOT_EXCLUDED    As Long = 6     'InputRows with status EXCLUDED (FX)
        Public Const TOT_MIXED       As Long = 7     'IR rows flagged MIXED
        Public Const TOT_BUCKETS     As Long = 8     'EQ markets or commodities found
        Public Const TOT_LONG        As Long = 9     'FX: sum of net long currencies
        Public Const TOT_SHORT       As Long = 10    'FX: sum of net short currencies
        Public Const TOT_GOLD        As Long = 11    'FX: |net gold|
        Public Const TOT_ONOP        As Long = 12    'FX: overall net open position
        Public Const TOT_COUNT       As Long = 12    'Size of the vector

'------------------------------------------------------------------------------
' ENGINE BUNDLES
'------------------------------------------------------------------------------
    'CORE_Engine.Calculate takes one Variant vector of inputs and returns one
    'of outputs, each element holding a value or an array. VBA user-defined
    'types are avoided so that the core also runs under LibreOffice's VBA
    'compatibility mode, which the test harness uses.
    'Inputs (raw cell values unless stated).
        Public Const IN_NUMERIC        As Long = 1     '(1 To PI_COUNT) numeric parameters
        Public Const IN_ASOF           As Long = 2     'AsOfDate cell
        Public Const IN_REPCCY         As Long = 3     'ReportingCcy cell
        Public Const IN_GOLD           As Long = 4     'GoldCode cell
        Public Const IN_LADDER         As Long = 5     '(1 To BAND_COUNT, 1 To LD_WIDTH)
        Public Const IN_LADDER_LABELS  As Long = 6     'String (1 To BAND_COUNT, 1 To 2)
        Public Const IN_SR_TABLE       As Long = 7     '(1 To n, 1 To 2) key, charge
        Public Const IN_FX_TABLE       As Long = 8     '(1 To m, 1 To 2) code, rate
        Public Const IN_OPTIONS        As Long = 9     '(1 To RC_COUNT) option charges
        Public Const IN_IR_ROWS        As Long = 10    'IRPositions cells
        Public Const IN_EQ_ROWS        As Long = 11    'EQPositions cells
        Public Const IN_CO_ROWS        As Long = 12    'COPositions cells
        Public Const IN_FX_ROWS        As Long = 13    'FXPositions cells
        Public Const IN_IR_FIRST_ROW   As Long = 14    'Sheet row of the first IR row
        Public Const IN_EQ_FIRST_ROW   As Long = 15    'Sheet row of the first EQ row
        Public Const IN_CO_FIRST_ROW   As Long = 16    'Sheet row of the first CO row
        Public Const IN_FX_FIRST_ROW   As Long = 17    'Sheet row of the first FX row
        Public Const IN_EQ_CAPACITY    As Long = 18    'Market rows of the Buckets block
        Public Const IN_CO_CAPACITY    As Long = 19    'Commodity rows of the Buckets block
        Public Const IN_COUNT          As Long = 19    'Size of the input vector
    'Outputs.
        Public Const OUT_IR_CALC       As Long = 1     'IRCalc table
        Public Const OUT_LADDER        As Long = 2     'IR_Ladder table
        Public Const OUT_ZONES         As Long = 3     'IR_Zones table with total row
        Public Const OUT_EQ_CALC       As Long = 4     'EQCalc table
        Public Const OUT_MARKETS       As Long = 5     'Buckets equity rows and total row
        Public Const OUT_CO_CALC       As Long = 6     'COCalc table
        Public Const OUT_COMMODITIES   As Long = 7     'Buckets commodity rows and total row
        Public Const OUT_FX_CALC       As Long = 8     'FXCalc table
        Public Const OUT_CURRENCIES    As Long = 9     'Buckets FX rows and total row
        Public Const OUT_FX_SUMMARY    As Long = 10    '(1 To 2, 1 To 1): ONOP, FX charge
        Public Const OUT_RESULTS       As Long = 11    'Results E:O, 4 classes and total
        Public Const OUT_RULES         As Long = 12    '(1 To RULE_COUNT, 1 To 2): rule, count
        Public Const OUT_FINDINGS      As Long = 13    'Checks findings C:I, or Empty
        Public Const OUT_STATUS        As Long = 14    'Status line for Results
        Public Const OUT_STOPPED       As Long = 15    'True when no results were calculated
        Public Const OUT_REQUIREMENT   As Long = 16    'Own funds requirement
        Public Const OUT_RWA           As Long = 17    'Risk-weighted assets
        Public Const OUT_ERRORS        As Long = 18    'Errors on Checks
        Public Const OUT_WARNINGS      As Long = 19    'Warnings on Checks
        Public Const OUT_USED          As Long = 20    'Positions used
        Public Const OUT_COUNT         As Long = 20    'Size of the output vector

'------------------------------------------------------------------------------
' RISK CLASSES
'------------------------------------------------------------------------------
        Public Const RC_IR           As Long = 1     'Interest rate
        Public Const RC_EQ           As Long = 2     'Equity
        Public Const RC_CO           As Long = 3     'Commodity
        Public Const RC_FX           As Long = 4     'Foreign exchange
        Public Const RC_COUNT        As Long = 4     'Number of risk classes
        Public Const RC_CODES        As String = "IR|EQ|CO|FX"    'Suffix of the Options names

'------------------------------------------------------------------------------
' PARAMETERS
'------------------------------------------------------------------------------
    'Text parameters, read through the workbook name equal to the code.
        Public Const PRM_ASOF        As String = "AsOfDate"        'Reporting date
        Public Const PRM_REPCCY      As String = "ReportingCcy"    'Reporting currency
        Public Const PRM_GOLD        As String = "GoldCode"        'Currency code of gold
    'Numeric parameters: index in the parameter vector, in the order of
    'PRM_NUMERIC_CODES; PRM_REFERENCE_VALUES gives the MAR40 value of each
    '(decimal point, read with Val so that the Windows locale is irrelevant).
        Public Const PI_DAYS_PER_YEAR    As Long = 1     'Days per year, residual maturity
        Public Const PI_COUPON_THRESHOLD As Long = 2     'Low-coupon threshold, decimal
        Public Const PI_SCALING_IR       As Long = 3     'Scaling factor, interest rate
        Public Const PI_SCALING_EQ       As Long = 4     'Scaling factor, equity
        Public Const PI_SCALING_CO       As Long = 5     'Scaling factor, commodity
        Public Const PI_SCALING_FX       As Long = 6     'Scaling factor, FX
        Public Const PI_RWA_MULTIPLIER   As Long = 7     'RWA multiplier
        Public Const PI_VERTICAL         As Long = 8     'Vertical disallowance
        Public Const PI_ZONE1            As Long = 9     'Horizontal, within zone 1
        Public Const PI_ZONE2            As Long = 10    'Horizontal, within zone 2
        Public Const PI_ZONE3            As Long = 11    'Horizontal, within zone 3
        Public Const PI_ADJACENT         As Long = 12    'Horizontal, adjacent zones
        Public Const PI_ZONE13           As Long = 13    'Horizontal, zones 1 and 3
        Public Const PI_EQ_SPECIFIC      As Long = 14    'Equity specific, single names
        Public Const PI_EQ_INDEX         As Long = 15    'Equity further charge, indices
        Public Const PI_EQ_GENERAL       As Long = 16    'Equity general market risk
        Public Const PI_FX_CHARGE        As Long = 17    'FX and gold charge
        Public Const PI_CO_DIRECTIONAL   As Long = 18    'Commodity directional charge
        Public Const PI_CO_GROSS         As Long = 19    'Commodity gross charge
        Public Const PI_COUNT            As Long = 19    'Size of the parameter vector
        Public Const PRM_NUMERIC_CODES   As String = _
            "DaysPerYear|CouponThreshold|ScalingIR|ScalingEQ|ScalingCO|ScalingFX|" & _
            "RWAMultiplier|VerticalDisallowance|HorizontalZone1|HorizontalZone2|" & _
            "HorizontalZone3|HorizontalAdjacent|HorizontalZone13|EQSpecificRate|" & _
            "EQIndexRate|EQGeneralRate|FXChargeRate|COMDirectionalRate|COMGrossRate"
        Public Const PRM_REFERENCE_VALUES As String = _
            "365|0.03|1.3|3.5|1.9|1.2|12.5|0.1|0.4|0.3|0.3|0.4|1|0.08|0.02|0.08|0.08|0.15|0.03"
        Public Const PRM_REFERENCE_GOLD  As String = "XAU"     'MAR40.60 reference gold code
    'Kind of check applied to each numeric parameter, in the same order:
    'P = > 0, R = rate in [0, 1], N = >= 0.
        Public Const PRM_DOMAINS         As String = "PNPPPPPRRRRRRRRRRRR"

'------------------------------------------------------------------------------
' MATURITY LADDER (MAR40.26 Table 4)
'------------------------------------------------------------------------------
    'Default ladder used by the SSA_ worksheet functions; RunSSA reads the
    'ladder from Params. Upper bounds in years, inclusive; OPEN_BOUND marks
    'an open-ended band and a band that does not exist for that coupon.
    'ReferenceLadder sets the one-month bound of band 1 to 1/12 exactly.
        Public Const BAND_COUNT          As Long = 15            'Time bands
        Public Const ZONE_COUNT          As Long = 3             'Zones
        Public Const OPEN_BOUND          As Double = 1E+300      'No upper bound
        Public Const LADDER_ZONES        As String = "1|1|1|1|2|2|2|3|3|3|3|3|3|3|3"
        Public Const LADDER_UPPER_HIGH   As String = _
            "0.0833333333333333|0.25|0.5|1|2|3|4|5|7|10|15|20|||"
        Public Const LADDER_UPPER_LOW    As String = _
            "0.0833333333333333|0.25|0.5|1|1.9|2.8|3.6|4.3|5.7|7.3|9.3|10.6|12|20|"
        Public Const LADDER_WEIGHTS      As String = _
            "0|0.002|0.004|0.007|0.0125|0.0175|0.0225|0.0275|0.0325|0.0375|0.045|0.0525|0.06|0.08|0.125"
    'Columns of the ladder matrix passed to the core.
        Public Const LD_ZONE             As Long = 1     'Zone 1 to 3
        Public Const LD_UPPER_HIGH       As Long = 2     'Upper bound, coupon >= threshold
        Public Const LD_UPPER_LOW        As Long = 3     'Upper bound, coupon < threshold
        Public Const LD_WEIGHT           As Long = 4     'Risk weight, decimal
        Public Const LD_WIDTH            As Long = 4     'Columns of the matrix
    'Index of each figure in the ladder-charge vector (IR_Zones E:AD).
        Public Const ZC_VERTICAL         As Long = 1     'Vertical disallowance
        Public Const ZC_ZONE_FIRST       As Long = 2     'Zone 1 block; each zone has 5 figures:
                                                         'net long, net short, matched, charge, net
        Public Const ZC_MATCHED12        As Long = 17    'Offset between zones 1 and 2
        Public Const ZC_Z1_AFTER12       As Long = 18    'Zone 1 net after it
        Public Const ZC_Z2_AFTER12       As Long = 19    'Zone 2 net after it
        Public Const ZC_MATCHED23        As Long = 20    'Offset between zones 2 and 3
        Public Const ZC_Z2_AFTER23       As Long = 21    'Zone 2 net after it
        Public Const ZC_Z3_AFTER23       As Long = 22    'Zone 3 net after it
        Public Const ZC_MATCHED13        As Long = 23    'Offset between zones 1 and 3
        Public Const ZC_BETWEEN          As Long = 24    'Between-zone charge
        Public Const ZC_NET              As Long = 25    'Residual net position charge
        Public Const ZC_GENERAL          As Long = 26    'General market risk
        Public Const ZC_COUNT            As Long = 26    'Size of the vector

'------------------------------------------------------------------------------
' DATES
'------------------------------------------------------------------------------
    'Serial numbers a cell can hold as a date, in Excel and in VBA. A larger
    'number, such as a date typed as 20311231, is not a date.
        Public Const DATE_SERIAL_MIN As Double = 1#          '1 January 1900
        Public Const DATE_SERIAL_MAX As Double = 2958465#    '31 December 9999

'------------------------------------------------------------------------------
' LABELS
'------------------------------------------------------------------------------
        Public Const ST_OK           As String = "OK"            'Row used
        Public Const ST_INVALID      As String = "INVALID: "     'Prefix, row rejected
        Public Const ST_EXCLUDED     As String = "EXCLUDED: "    'Prefix, row out of scope
        Public Const SEV_ERROR       As String = "ERROR"         'Checks severity
        Public Const SEV_WARNING     As String = "WARNING"       'Checks severity
        Public Const SEV_INFO        As String = "INFO"          'Checks severity
        Public Const SR_NONE         As String = "NONE"          'No specific risk
        Public Const SR_NO_MATCH     As String = "NO MATCH"      'No Table 1 charge
        Public Const SR_CATEGORIES   As String = "|GOV|QUAL|OTHER|NONE|"    'Valid categories
        Public Const LBL_MIXED       As String = "MIXED"         'Issue with mixed keys
        Public Const LBL_NOT_APPLY   As String = "n/a"           'Band absent for this coupon
        Public Const FX_REPORTING    As String = "Reporting ccy" 'FX block treatment
        Public Const FX_GOLD_LABEL   As String = "Gold"          'FX block treatment
        Public Const FX_CURRENCY     As String = "Currency"      'FX block treatment
        Public Const LBL_TOTAL_ZONES As String = "Total"              'IR_Zones total row
        Public Const LBL_TOTAL_EQ    As String = "Total equity"       'Buckets equity total row
        Public Const LBL_TOTAL_CO    As String = "Total commodities"  'Buckets commodity total row
        Public Const LBL_TOTAL_FX    As String = "Totals"             'Buckets FX total row

'------------------------------------------------------------------------------
' CHECKS RULES
'------------------------------------------------------------------------------
        Public Const RULE_IR_ROWS    As String = "C01"    'IR rows rejected
        Public Const RULE_EQ_ROWS    As String = "C02"    'EQ rows rejected
        Public Const RULE_CO_ROWS    As String = "C03"    'CO rows rejected
        Public Const RULE_FX_ROWS    As String = "C04"    'FX rows rejected
        Public Const RULE_FX_REPCCY  As String = "C05"    'Reporting currency not at rate 1
        Public Const RULE_FX_DUP     As String = "C06"    'Duplicate FX codes
        Public Const RULE_FX_RATE    As String = "C07"    'Invalid FX rates
        Public Const RULE_CHANGED    As String = "C08"    'Constants differ from MAR40
        Public Const RULE_MIXED      As String = "C09"    'IR issue with mixed keys
        Public Const RULE_EQ_CAP     As String = "C10"    'Equity markets beyond Buckets rows
        Public Const RULE_CO_CAP     As String = "C11"    'Commodities beyond Buckets rows
        Public Const RULE_FX_EXCL    As String = "C12"    'FX rows in reporting currency
        Public Const RULE_OPTIONS    As String = "C13"    'Option charges entered
        Public Const RULE_PARAMS     As String = "C14"    'Unusable parameter or table value, or a position sheet over MAX_POSITION_ROWS
        Public Const RULE_COUNT      As Long = 14         'Rules C01 to C14
        Public Const RULE_SEVERITIES As String = _
            "ERROR|ERROR|ERROR|ERROR|ERROR|ERROR|ERROR|WARNING|WARNING|WARNING|WARNING|INFO|INFO|ERROR"

'------------------------------------------------------------------------------
' ERROR NUMBERS
'------------------------------------------------------------------------------
    'Raised by the core and the macros; range vbObjectError + 2048 to 4095.
    'Written as literals equal to vbObjectError + n: LibreOffice, the optional
    'smoke host, cannot evaluate vbObjectError in a constant expression.
        Public Const ERR_LAYOUT          As Long = &H80040801    'vbObjectError + 2049: Template name or layout missing
        Public Const ERR_ARGUMENT        As Long = &H80040802    'vbObjectError + 2050: Core called with inconsistent arrays
        Public Const ERR_CLEANUP_FAILED  As Long = &H80040803    'vbObjectError + 2051: Excel settings not restored
