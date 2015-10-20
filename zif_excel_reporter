*"* components of interface ZIF_EXCEL_REPORTER
interface ZIF_EXCEL_REPORTER
  public .


  types ROW_ID type I .
  types COL_ID type I .
  types:
    rows_range TYPE RANGE OF row_id .
  types:
    cols_range TYPE RANGE OF col_id .

  class-methods ADDR2STR
    importing
      !I_COL type COL_ID optional
      !I_ROW type ROW_ID optional
      !I_COL_E type COL_ID optional
      !I_ROW_E type ROW_ID optional
    returning
      value(R_ADDRESS) type STRING .
  class-methods STR2ADDR
    importing
      value(I_ADDRESS) type STRING
    exporting
      !E_SHEET type STRING
      !E_ROW type ROW_ID
      !E_COL type COL_ID
      !E_ROW_E type ROW_ID
      !E_COL_E type COL_ID
      !ET_COLS_RNG type COLS_RANGE
      !ET_ROWS_RNG type ROWS_RANGE .
  class-methods COL_ID2STR
    importing
      value(I_COL) type COL_ID
    returning
      value(R_STR) type STRING .
  class-methods STR2COL_ID
    importing
      !I_ADDRESS type ANY
    returning
      value(R_ID) type I .
  methods OUT_SECTION
    importing
      !NAME type STRING
      !DATA type ANY optional
      !MERGE type STRING optional .
  methods JOIN_SECTION
    importing
      !NAME type STRING
      !DATA type ANY optional
      !MERGE type STRING optional .
  methods ADD_SHEET
    importing
      !SHEET_NAME type STRING default ''
      !TITLE type STRING default ''
    preferred parameter SHEET_NAME
    returning
      value(SHEET) type ref to ZIF_EXCEL_REPORTER .
  methods NEW_PAGE .
  methods PAGE_SETUP
    importing
      !BOTTOM_MARGIN type F default -1
      !FIT_TO_PAGES_TALL type I default -1
      !FIT_TO_PAGES_WIDE type I default -1
      !LEFT_MARGIN type F default -1
      !ORIENTATION type I default -1
      !RIGHT_MARGIN type F default -1
      !TOP_MARGIN type F default -1
      !HEADER_MARGIN type F default -1
      !FOOTER_MARGIN type F default -1
      !ZOOM type I default -1
      !TITLE_ROWS_FROM type I default -1
      !TITLE_ROWS_TO type I default -1
      !TITLE_COLS_FROM type I default -1
      !TITLE_COLS_TO type I default -1
      !FIX_ROWS type I default 0
      !FIX_COLS type I default 0
      !LEFT_HEADER type STRING default ''
      !CENTER_HEADER type STRING default ''
      !RIGHT_HEADER type STRING default ''
      !LEFT_FOOTER type STRING default ''
      !CENTER_FOOTER type STRING default ''
      !RIGHT_FOOTER type STRING default ''
      !PRINT_AREA type STRING default '' .
  methods SAVE_AS
    importing
      !FILENAME type STRING .
  methods SHOW
    importing
      !BOTTOM_MARGIN type F default -1
      !FIT_TO_PAGES_TALL type I default -1
      !FIT_TO_PAGES_WIDE type I default -1
      !LEFT_MARGIN type F default -1
      !ORIENTATION type I default -1
      !RIGHT_MARGIN type F default -1
      !TOP_MARGIN type F default -1
      !HEADER_MARGIN type F default -1
      !FOOTER_MARGIN type F default -1
      !ZOOM type I default -1
      !TITLE_ROWS_FROM type I default -1
      !TITLE_ROWS_TO type I default -1
      !TITLE_COLS_FROM type I default -1
      !TITLE_COLS_TO type I default -1
      !FIX_ROWS type I default 0
      !FIX_COLS type I default 0
      !LEFT_HEADER type STRING default ''
      !CENTER_HEADER type STRING default ''
      !RIGHT_HEADER type STRING default ''
      !LEFT_FOOTER type STRING default ''
      !CENTER_FOOTER type STRING default ''
      !RIGHT_FOOTER type STRING default ''
      !PRINT_AREA type STRING default '' .
  methods GROUP_BEGIN .
  methods GROUP_END .
  methods COL_GROUP
    importing
      value(COLUMNS) type STRING default ''
      !FROM type I default 0
      !TO type I default 0 .
  methods FREE .
endinterface.
