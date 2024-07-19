*---------------------------------------------------------------------*
*       FORM FRM_CREATE_BASIC_OBJECTS                                 *
*---------------------------------------------------------------------*
FORM FRM_CREATE_BASIC_OBJECTS.
  DATA: LO_BDS_INSTANCE       TYPE REF TO CL_BDS_DOCUMENT_SET,
        LT_DOC_SIGNATURE      TYPE SBDST_SIGNATURE,
        LS_DOC_SIGNATURE      LIKE LINE OF LT_DOC_SIGNATURE,
        LT_DOC_COMPONENTS     TYPE SBDST_COMPONENTS,
        LT_DOC_URIS           TYPE SBDST_URI,
        LS_DOC_URIS           LIKE LINE OF LT_DOC_URIS,
        LV_ITEM_URL(256)      TYPE C,
        LGO_DOCUMENT_TYPE(80) TYPE C,
        LV_FILESIZE           TYPE I,
        LT_DATA_TABLE         TYPE SBDST_CONTENT,
        LV_PROTECT            TYPE C,
        LT_SHEETS             TYPE SOI_SHEETS_TABLE.

  CLEAR LV_ITEM_URL.
  LS_DOC_SIGNATURE-PROP_VALUE = 'ZPDND002MT_' && GV_EXCEL_LANG.
  LS_DOC_SIGNATURE-PROP_NAME  = 'DESCRIPTION'.
  LGO_DOCUMENT_TYPE            = C_EXCEL.

  APPEND LS_DOC_SIGNATURE TO LT_DOC_SIGNATURE.

  CREATE OBJECT LO_BDS_INSTANCE.

  CALL METHOD LO_BDS_INSTANCE->GET_INFO
    EXPORTING
      CLASSNAME  = C_DOC_CLASSNAME
      CLASSTYPE  = C_DOC_CLASSTYPE
      OBJECT_KEY = C_DOC_OBJECT_KEY
    CHANGING
      COMPONENTS = LT_DOC_COMPONENTS
      SIGNATURE  = LT_DOC_SIGNATURE.

  CALL METHOD LO_BDS_INSTANCE->GET_WITH_URL
    EXPORTING
      CLASSNAME  = C_DOC_CLASSNAME
      CLASSTYPE  = C_DOC_CLASSTYPE
      OBJECT_KEY = C_DOC_OBJECT_KEY
    CHANGING
      URIS       = LT_DOC_URIS
      SIGNATURE  = LT_DOC_SIGNATURE.

  FREE LO_BDS_INSTANCE.

  READ TABLE LT_DOC_URIS INTO LS_DOC_URIS INDEX 1.
  IF SY-SUBRC <> 0.
    CLEAR LS_DOC_URIS.
  ENDIF.
  LV_ITEM_URL = LS_DOC_URIS-URI.

  CALL METHOD GO_CONTROL->GET_DOCUMENT_PROXY
    EXPORTING
      DOCUMENT_TYPE  = LGO_DOCUMENT_TYPE
    IMPORTING
      DOCUMENT_PROXY = GO_DOCUMENT
      ERROR          = GV_ERROR
      RETCODE        = GV_RETCODE.

  CALL METHOD GO_DOCUMENT->OPEN_DOCUMENT
    EXPORTING
      DOCUMENT_URL = LV_ITEM_URL
      OPEN_INPLACE = C_INPLACE
    IMPORTING
      ERROR        = GV_ERROR
      RETCODE      = GV_RETCODE.


  IF GV_RETCODE = 'OK'.

    CALL METHOD GO_DOCUMENT->GET_APPLICATION_PROPERTY
      EXPORTING
        NO_FLUSH         = ' '
        PROPERTY_NAME    = 'INTERNATIONAL'
        SUBPROPERTY_NAME = 'DECIMAL_SEPARATOR'
      IMPORTING
        ERROR            = GV_ERROR
        RETCODE          = GV_RETCODE
      CHANGING
        RETVALUE         = GV_DECIMAL_SEP.

    DATA: HAS TYPE I.

    CALL METHOD GO_DOCUMENT->HAS_SPREADSHEET_INTERFACE
      IMPORTING
        IS_AVAILABLE = HAS.
*
    IF NOT HAS IS INITIAL.
      CALL METHOD GO_DOCUMENT->GET_SPREADSHEET_INTERFACE
        IMPORTING
          SHEET_INTERFACE = GO_HANDLE.

      CALL METHOD GO_HANDLE->GET_SHEETS
        EXPORTING
          NO_FLUSH = SPACE
        IMPORTING
          SHEETS   = LT_SHEETS
          ERROR    = GV_ERROR
          RETCODE  = GV_RETCODE.

    ENDIF. "IF RETCODE = 'OK'
  ENDIF.

ENDFORM.                    "create_basic_objects
*&---------------------------------------------------------------------*
*& Form FRM_DISPLAY_ORDERS_DATA
*&---------------------------------------------------------------------*
FORM FRM_DISPLAY_DATA.

  " check the document has not been created yet
  CHECK GV_FIRST_SET IS INITIAL.

  IF GO_HANDLE IS NOT INITIAL.

    " Insert the data into the excel document
    PERFORM FRM_DISPLAY_EXCEL_DATA.

    " Protect the sheet from editing
    PERFORM FRM_PROTECT_SHEET.

    GV_FIRST_SET = 'X'.
  ENDIF.
ENDFORM. " FRM_DISPLAY_ORDERS_DATA

*&---------------------------------------------------------------------*
*& Form FRM_INSERT_TABLE
*&---------------------------------------------------------------------*
FORM FRM_INSERT_TABLE USING   UT_TABLE TYPE STANDARD TABLE
                              UP_TABNAME       TYPE  X030L-TABNAME
                              US_RANGE_DETAILS TYPE  SOI_FULL_RANGE_ITEM.
  DATA: LT_FIELDS_TABLE TYPE TABLE OF RFC_FIELDS,
        LV_ROWS         TYPE I,
        LV_COLUMNS      TYPE I.

  IF UT_TABLE[] IS NOT INITIAL.
    "Get number of columns
    TRY.
        LO_TABLE_DESCR ?= CL_ABAP_TABLEDESCR=>DESCRIBE_BY_DATA( P_DATA = UT_TABLE ).
        LO_STRUCT_DESCR ?= LO_TABLE_DESCR->GET_TABLE_LINE_TYPE( ).
        LV_COLUMNS = LINES( LO_STRUCT_DESCR->COMPONENTS ).
      CATCH CX_SY_MOVE_CAST_ERROR.
    ENDTRY.
    "Get number of rows
    LV_ROWS = LINES( UT_TABLE ).


    CALL METHOD GO_HANDLE->INSERT_RANGE_DIM
      EXPORTING
        NAME     = US_RANGE_DETAILS-NAME
        NO_FLUSH = ''
        TOP      = GV_LINE_COUNT
        LEFT     = US_RANGE_DETAILS-LEFT
        ROWS     = LV_ROWS
        COLUMNS  = LV_COLUMNS
      IMPORTING
        ERROR    = GV_ERROR
        RETCODE  = GV_RETCODE.


    CALL FUNCTION 'DP_GET_FIELDS_FROM_TABLE'
      EXPORTING
        TABNAME          = UP_TABNAME
      TABLES
        DATA             = UT_TABLE
        FIELDS           = LT_FIELDS_TABLE
      EXCEPTIONS
        DP_INVALID_TABLE = 1
        OTHERS           = 2.


    CALL METHOD GO_HANDLE->INSERT_ONE_TABLE
      EXPORTING
        DATA_TABLE   = UT_TABLE                  " Data
        FIELDS_TABLE = LT_FIELDS_TABLE                " The Fields of the Table
        RANGENAME    = US_RANGE_DETAILS-NAME              " The Name of the Range
        NO_FLUSH     = 'X'              " Flush?
        WHOLETABLE   = 'X'               " Inserts Whole Table
*       UPDATING     = -1               " Screen Updating
      IMPORTING
        ERROR        = GV_ERROR
        RETCODE      = GV_RETCODE.
**  Bordering the table
    CALL METHOD GO_HANDLE->SET_FRAME
      EXPORTING
        RANGENAME = US_RANGE_DETAILS-NAME
        TYP       = 191
        COLOR     = 16.

    CALL METHOD GO_HANDLE->SET_COLOR
      EXPORTING
        RANGENAME = US_RANGE_DETAILS-NAME
        FRONT     = -1
        BACK      = 15 "20
        NO_FLUSH  = SPACE
      IMPORTING
        ERROR     = GV_ERROR
        RETCODE   = GV_RETCODE.

  ENDIF. "IF UT_TABLE[] IS NOT INITIAL
ENDFORM. " FRM_INSERT_TABLE

*&---------------------------------------------------------------------*
*&      Form  FRM_SET_EDITABLE
*&---------------------------------------------------------------------*
FORM FRM_SET_EDITABLE  TABLES PT_DATA
                       USING   UP_EDITABLE_COLS TYPE STRING
                             UP_TAB_TYPE TYPE X030L-TABNAME
                              UP_START_ROW     TYPE I.

  DATA: LV_TABIX    TYPE SY-TABIX,
        LINETYPE    TYPE STRING,
        ITABREF     TYPE REF TO DATA,
        LINE        TYPE REF TO DATA,
        LT_EDIT_COL TYPE STANDARD TABLE OF TY_EDIT_COL,
        LS_EDIT_COL TYPE TY_EDIT_COL.

  CLEAR T_SOI_CELL_TABLE[].

  SPLIT UP_EDITABLE_COLS AT '-' INTO TABLE LT_EDIT_COL.

  CREATE DATA ITABREF TYPE STANDARD TABLE OF (UP_TAB_TYPE)."ZPDND003_MULTI_ALV1.
  ASSIGN ITABREF->* TO <FS>.

*Importing content from memory id
  <FS> = PT_DATA[].

* CREATION OF A STRUCTURE :

  CREATE DATA LINE LIKE LINE OF <FS>.
  ASSIGN LINE->* TO <FL>.
* loop at t_exceldata.
  LOOP AT LT_EDIT_COL ASSIGNING FIELD-SYMBOL(<FS_EDIT_COL>).
    LV_TABIX = UP_START_ROW ." GV_LINE_COUNT.
    LOOP AT <FS> ASSIGNING <FL>.

*   POC, forcasst date and actual date editable
      S_SOI_CELL_ITEM-TOP        = LV_TABIX .
      S_SOI_CELL_ITEM-LEFT       = <FS_EDIT_COL>-COL.
      S_SOI_CELL_ITEM-ROWS       = 1.
      S_SOI_CELL_ITEM-COLUMNS    = 1.
      S_SOI_CELL_ITEM-BACK       = 2.
      S_SOI_CELL_ITEM-FRAMETYP   = -1.
      S_SOI_CELL_ITEM-SIZE       = -1.
      S_SOI_CELL_ITEM-BOLD       = -1.
      S_SOI_CELL_ITEM-ITALIC     = -1.
      S_SOI_CELL_ITEM-ALIGN      = -1.
      S_SOI_CELL_ITEM-FRAMECOLOR = -1.
      S_SOI_CELL_ITEM-DECIMALS = 3.
      S_SOI_CELL_ITEM-NUMBER = 1.

      APPEND S_SOI_CELL_ITEM TO T_SOI_CELL_TABLE.
      LV_TABIX = LV_TABIX  + 1.
    ENDLOOP.
  ENDLOOP.

  IF NOT T_SOI_CELL_TABLE[] IS INITIAL.
    CALL METHOD GO_HANDLE->CELL_FORMAT
      EXPORTING
        CELLS   = T_SOI_CELL_TABLE
      IMPORTING
        ERROR   = GV_ERROR
        RETCODE = GV_RETCODE.
  ENDIF.
  CLEAR: S_SOI_CELL_ITEM,T_SOI_CELL_TABLE[] .
ENDFORM.                    " FRM_SET_EDITABLE

*&---------------------------------------------------------------------*
*& Form FRM_PROTECT_SHEET
*&---------------------------------------------------------------------*
FORM FRM_PROTECT_SHEET .
  DATA: LV_PROTECT TYPE C.
  DATA: LT_SHEETS TYPE SOI_SHEETS_TABLE .
  DATA: LS_SHEETS TYPE SOI_SHEETS.
  IF GO_HANDLE IS NOT INITIAL.
    CALL METHOD GO_HANDLE->SET_ZOOM
      EXPORTING
        ZOOM    = 85
      IMPORTING
        ERROR   = GV_ERROR
        RETCODE = GV_RETCODE.
    CALL METHOD GO_HANDLE->GET_SHEETS
      IMPORTING
        SHEETS  = LT_SHEETS
        ERROR   = GV_ERROR
        RETCODE = GV_RETCODE.

    LOOP AT LT_SHEETS INTO LS_SHEETS.

      CALL METHOD GO_HANDLE->GET_PROTECTION
        EXPORTING
          SHEETNAME = LS_SHEETS-SHEET_NAME
        IMPORTING
          ERROR     = GV_ERROR
          RETCODE   = GV_RETCODE
          PROTECT   = LV_PROTECT.

* If not protected, protect the sheet
      IF LV_PROTECT NE 'X'.
        CALL METHOD GO_HANDLE->PROTECT
          EXPORTING
            PROTECT = 'X'
          IMPORTING
            ERROR   = GV_ERROR
            RETCODE = GV_RETCODE.
      ENDIF.
    ENDLOOP.
  ENDIF.
ENDFORM.

*---------------------------------------------------------------------*
*       FORM FRM_GET_EXCEL_DATA                                       *
*---------------------------------------------------------------------*
FORM FRM_GET_EXCEL_DATA TABLES PT_DATA
                               UT_RANGE_ITEMS TYPE SOI_GENERIC_TABLE
                        USING  US_RANGE TYPE SOI_FULL_RANGE_ITEM
                               UP_TAB_TYPE TYPE X030L-TABNAME.
  DATA L_TYPE           TYPE STRING.
  DATA WO_OREF          TYPE REF TO CX_ROOT.
  DATA L_FLOAT          TYPE F.
  DATA L_TEXT           TYPE STRING.
  DATA LV_TYPE TYPE  X030L.
  DATA LT_DATA  TYPE REF TO DATA.
  DATA L_DATA TYPE REF TO DATA.
  FIELD-SYMBOLS: <FS>    TYPE STANDARD TABLE,
                 <FL>    TYPE ANY,
                 <FIELD> TYPE ANY.
  CREATE DATA LT_DATA TYPE STANDARD TABLE OF (UP_TAB_TYPE).
  ASSIGN LT_DATA->* TO <FS>.
  <FS> = PT_DATA[].
  CREATE DATA L_DATA LIKE LINE OF <FS>.
  ASSIGN L_DATA->* TO <FL>.

  CHECK NOT GO_HANDLE IS INITIAL.
  CLEAR PT_DATA[].
  CLEAR <FL>.
  DATA: W_COLUMN         TYPE I.
  LOOP AT UT_RANGE_ITEMS INTO DATA(WS_CONTENT). " HERE
    W_COLUMN = WS_CONTENT-COLUMN.
    ASSIGN COMPONENT W_COLUMN OF STRUCTURE <FL> TO  <FIELD>.
    IF SY-SUBRC EQ 0.
      TRY.
          DESCRIBE FIELD <FIELD> TYPE   L_TYPE.
          CASE L_TYPE.

            WHEN 'D'.
              CONCATENATE  WS_CONTENT-VALUE+6(4)
                           WS_CONTENT-VALUE+3(2)
                           WS_CONTENT-VALUE(2)
                           INTO <FIELD>.
            WHEN 'P'.
              CLEAR L_FLOAT.
              REPLACE ',' IN WS_CONTENT-VALUE WITH '.'.
              L_FLOAT = WS_CONTENT-VALUE.
              <FIELD> = L_FLOAT.

            WHEN OTHERS.
              REPLACE '''=' IN WS_CONTENT-VALUE WITH '='.
              <FIELD> = WS_CONTENT-VALUE.

          ENDCASE.

        CATCH CX_SY_ARITHMETIC_ERROR INTO WO_OREF.
          " Alimenter une table des erreurs
          L_TEXT = WO_OREF->GET_TEXT( ).
          EXIT.
        CATCH CX_SY_CONVERSION_NO_NUMBER INTO WO_OREF.
          " Alimenter une table des erreurs
          L_TEXT = WO_OREF->GET_TEXT( ).
          EXIT.
        CATCH CX_SY_CONVERSION_OVERFLOW INTO WO_OREF. "+CBIT 07.02.2014
          " Alimenter une table des erreurs
          L_TEXT = WO_OREF->GET_TEXT( ).   "+CBIT 07.02.2014
          EXIT.
      ENDTRY.
    ENDIF.
    AT END OF ROW.
*      IF <FL> IS NOT INITIAL.
      APPEND <FL> TO PT_DATA.
      CLEAR <FL>.
*      ENDIF.
      "$$
    ENDAT.
  ENDLOOP.
ENDFORM. " FRM_GET_EXCEL_DATA

*&---------------------------------------------------------------------*
*& Form FRM_INSERT_DATA
*&---------------------------------------------------------------------*
FORM FRM_INSERT_DATA  TABLES  UT_TABLE              TYPE STANDARD TABLE
                      USING   UP_TABNAME            TYPE  X030L-TABNAME
                              US_RANGE_DETAILS      TYPE  SOI_FULL_RANGE_ITEM
                              UP_HIDE               TYPE ABAP_BOOL
                              UP_HIDE_TRAINS        TYPE STRING.
  DATA: LT_FIELDS_TABLE TYPE TABLE OF RFC_FIELDS,
        LT_PATTERN      TYPE TABLE OF TY_PATTERN,
        LV_HIDE_TOP     TYPE I,
        LV_HIDE_ROWS    TYPE I.

  DATA: LT_TABLE TYPE TABLE OF ZPDND002MT_ALV4_C.
  FIELD-SYMBOLS: <FS_TABLE> TYPE  ZPDND002MT_ALV4_C.

  IF UP_HIDE_TRAINS IS NOT INITIAL.
    """"" HIDE TRAINS
    DATA: LT_TRAINS_HIDE TYPE STANDARD TABLE OF TY_EDIT_COL,
          LS_TRAINS_HIDE TYPE TY_EDIT_COL.
    SPLIT UP_HIDE_TRAINS   AT '-' INTO TABLE LT_TRAINS_HIDE.
  ENDIF.

  IF UT_TABLE[] IS NOT INITIAL .
    LV_HIDE_TOP =  US_RANGE_DETAILS-TOP  + LINES( UT_TABLE ).
    LV_HIDE_ROWS = US_RANGE_DETAILS-ROWS - LINES( UT_TABLE ).
    CALL FUNCTION 'DP_GET_FIELDS_FROM_TABLE'
      EXPORTING
        TABNAME          = UP_TABNAME
      TABLES
        DATA             = UT_TABLE
        FIELDS           = LT_FIELDS_TABLE
      EXCEPTIONS
        DP_INVALID_TABLE = 1
        OTHERS           = 2.
    IF UP_HIDE_TRAINS IS NOT INITIAL.
      """"" HIDE TRAINS
      DATA LV_TRAIN_HIDE TYPE ABAP_BOOL VALUE ABAP_FALSE.
      LOOP AT LT_TRAINS_HIDE INTO LS_TRAINS_HIDE.
        DELETE LT_FIELDS_TABLE WHERE FIELDNAME CP LS_TRAINS_HIDE.
        LV_TRAIN_HIDE = ABAP_TRUE.
      ENDLOOP.

" Modify the formula of quantity AC utilities
      DATA(LV_COL) =   LINES(  LT_FIELDS_TABLE ).
      READ TABLE LT_FIELDS_TABLE INTO DATA(LS_FIELDS) INDEX ( LV_COL - 1 ).
      IF SY-SUBRC = 0.
        DATA(LV_FIELD) = LS_FIELDS-FIELDNAME.
        READ TABLE UT_TABLE INDEX 8 ASSIGNING <FS_TABLE>.
        IF SY-SUBRC = 0.
         ASSIGN COMPONENT LV_FIELD OF STRUCTURE <FS_TABLE> TO FIELD-SYMBOL(<FS_FIELD>).
         DATA(LV_LETTER) = SY-ABCDE+LV_COL(1).
          REPLACE 'K' WITH LV_LETTER INTO <FS_TABLE>-UTILITES.
        ENDIF.
      ENDIF.
    ENDIF.

    CALL METHOD GO_HANDLE->INSERT_ONE_TABLE
      EXPORTING
        DATA_TABLE   = UT_TABLE[]
        FIELDS_TABLE = LT_FIELDS_TABLE
        RANGENAME    = US_RANGE_DETAILS-NAME
        NO_FLUSH     = 'X'
        WHOLETABLE   = 'X'
      IMPORTING
        ERROR        = GV_ERROR
        RETCODE      = GV_RETCODE.

**    Hide extra rows in excel

    IF  UP_HIDE = ABAP_TRUE.

      CALL METHOD GO_HANDLE->INSERT_RANGE_DIM
        EXPORTING
          NAME     = 'HIDE' && US_RANGE_DETAILS-NAME
          NO_FLUSH = 'X'
          TOP      = LV_HIDE_TOP
          LEFT     = US_RANGE_DETAILS-LEFT
          ROWS     = LV_HIDE_ROWS
          COLUMNS  = US_RANGE_DETAILS-COLUMNS
        IMPORTING
          ERROR    = GV_ERROR
          RETCODE  = GV_RETCODE.

      CALL METHOD GO_HANDLE->HIDE_ROWS
        EXPORTING
          NAME     = 'HIDE' && US_RANGE_DETAILS-NAME
          NO_FLUSH = 'X'
        IMPORTING
          ERROR    = GV_ERROR
          RETCODE  = GV_RETCODE.
      CALL METHOD GO_HANDLE->CLEAR_CONTENT_RANGE
        EXPORTING
          NAME     = 'HIDE'
          NO_FLUSH = 'X'
        IMPORTING
          ERROR    = GV_ERROR
          RETCODE  = GV_RETCODE.
    ENDIF. " IF  UP_HIDE = ABAP_TRUE.
  ENDIF. "IF UT_TABLE[] IS NOT INITIAL
ENDFORM. " FRM_INSERT_DATA

FORM FRM_HIDE_TRAINS USING UP_RANGE TYPE SOI_FULL_RANGE_ITEM
                            UP_HIDE_TRAINS TYPE STRING.
  """"" HIDE TRAINS
  DATA: LT_TRAINS_HIDE TYPE STANDARD TABLE OF TY_EDIT_COL,
        LS_TRAINS_HIDE TYPE TY_EDIT_COL.
  SPLIT UP_HIDE_TRAINS   AT '-' INTO TABLE LT_TRAINS_HIDE.

  CALL METHOD GO_HANDLE->INSERT_RANGE_DIM
    EXPORTING
      NAME     = 'HIDE_TRAIN'
      NO_FLUSH = 'X'
      TOP      = UP_RANGE-TOP
      LEFT     = 3 + ( 10 - LINES( LT_TRAINS_HIDE ) )
      ROWS     = 8
      COLUMNS  = LINES( LT_TRAINS_HIDE )
    IMPORTING
      ERROR    = GV_ERROR
      RETCODE  = GV_RETCODE.

  GO_HANDLE->CLEAR_RANGE(
 EXPORTING
   NAME     =  'HIDE_TRAIN'   ).

  DATA: LT_FORMAT_TABLE TYPE SOI_FORMAT_TABLE,
        LS_FORMAT_LINE  TYPE LINE OF SOI_FORMAT_TABLE.
  DATA TYP TYPE I.
  FIELD-SYMBOLS <FS1> TYPE X.

  LS_FORMAT_LINE-NAME = 'HIDE_TRAIN'.
  LS_FORMAT_LINE-BACK       = 2.
  LS_FORMAT_LINE-FRAMETYP   = 129.
  LS_FORMAT_LINE-SIZE       = -1.
  LS_FORMAT_LINE-BOLD       = -1.
  LS_FORMAT_LINE-ITALIC     = -1.
  LS_FORMAT_LINE-ALIGN      = -1.
  APPEND LS_FORMAT_LINE TO LT_FORMAT_TABLE.

  GO_HANDLE->SET_RANGES_FORMAT(
    EXPORTING
      FORMATTABLE = LT_FORMAT_TABLE  ).


  DATA: LT_RANGE_LIST TYPE SOI_RANGE_LIST,
        LS_RANGE_LIST TYPE LINE OF SOI_RANGE_LIST.
  LS_RANGE_LIST-NAME = 'HIDE_TRAIN'.
  APPEND LS_RANGE_LIST TO LT_RANGE_LIST.

  GO_HANDLE->DELETE_RANGES(
    EXPORTING
      RANGES   =  LT_RANGE_LIST    ).


ENDFORM.
*&---------------------------------------------------------------------*
*& Form FRM_DISPLAY_EXCEL_DATA
*&---------------------------------------------------------------------*
FORM FRM_DISPLAY_EXCEL_DATA.
  DATA: LS_RANGE1        TYPE  SOI_FULL_RANGE_ITEM,
        LS_RANGE         TYPE  SOI_FULL_RANGE_ITEM,
        GT_RANGE_CONTENT TYPE SOI_GENERIC_TABLE.
  IF GO_HANDLE IS NOT INITIAL.

    LOOP AT GT_TAB_BORDERS ASSIGNING FIELD-SYMBOL(<FS_TAB_BORDERS>).
      LS_RANGE-NAME    = <FS_TAB_BORDERS>-CODE_ZONE.
      LS_RANGE-TOP     = <FS_TAB_BORDERS>-TOP.
      LS_RANGE-LEFT    = <FS_TAB_BORDERS>-LEFT.
      LS_RANGE-ROWS    = <FS_TAB_BORDERS>-ROWS.
      LS_RANGE-COLUMNS = <FS_TAB_BORDERS>-COLS.

      CLEAR: GT_RANGE_CONTENT[].
      REFRESH T_RANGES.
      CLEAR T_RANGES.
      T_RANGES-NAME = LS_RANGE-NAME.
      APPEND T_RANGES.
      CALL METHOD GO_HANDLE->GET_RANGES_DATA
        EXPORTING
          ALL      = ''
        IMPORTING
          CONTENTS = GT_RANGE_CONTENT[]
          ERROR    = GV_ERROR
          RETCODE  = GV_RETCODE
        CHANGING
          RANGES   = T_RANGES[].

      CASE <FS_TAB_BORDERS>-CODE_ZONE.
        WHEN 'ZPDND002MT_TAB1'.
*****************************************************************************
**  Insert Table 1  ALV 1
*****************************************************************************
          PERFORM FRM_GET_EXCEL_DATA  TABLES   GT_EXCEL_DATA_1[]
                                               GT_RANGE_CONTENT[]
                                      USING    LS_RANGE
                                               'ZPDND002MT_ALV1'.

          LOOP AT GT_EXCEL_DATA_1 ASSIGNING FIELD-SYMBOL(<FS_EXCEL_1>).
            READ TABLE GT_ZPDND002MT_ALV1-TAB1 ASSIGNING FIELD-SYMBOL(<FS_TAB1>) INDEX SY-TABIX.
            IF SY-SUBRC = 0.
              <FS_TAB1>-QTY_PCS = <FS_EXCEL_1>-QTY_PCS.
              <FS_TAB1>-BLANK4  = <FS_EXCEL_1>-BLANK4.
            ENDIF.
          ENDLOOP.
          PERFORM FRM_INSERT_DATA     TABLES   GT_ZPDND002MT_ALV1-TAB1[]
                                      USING    'ZPDND002MT_ALV1'
                                               LS_RANGE
                                               'X'
                                               ''.
        WHEN 'ZPDND002MT_TAB1_TT'.
********************************************************************************
** Insert totals Table 1
********************************************************************************
          PERFORM FRM_GET_EXCEL_DATA  TABLES   GT_EXCEL_DATA_TOTAL1[]
                                               GT_RANGE_CONTENT[]
                                      USING    LS_RANGE
                                               'ZPDND002MT_ALV1_HEAD'.

          PERFORM FRM_INSERT_DATA     TABLES   GT_EXCEL_DATA_TOTAL1[]
                                      USING    'ZPDND002MT_ALV1_HEAD'
                                               LS_RANGE
                                               ''
                                               ''.
        WHEN 'ZPDND002MT_TAB2'.
*****************************************************************************
**  Insert Table 2 ALV 1
*****************************************************************************
          PERFORM FRM_GET_EXCEL_DATA  TABLES   GT_EXCEL_DATA_2[]
                                               GT_RANGE_CONTENT[]
                                      USING    LS_RANGE
                                               'ZPDND002MT_ALV2'.
          DATA LV_CHAR TYPE CHAR40.
          LOOP AT GT_EXCEL_DATA_2 ASSIGNING FIELD-SYMBOL(<FS_EXCEL_2>).
            READ TABLE GT_ZPDND002MT_ALV1-TAB2 ASSIGNING FIELD-SYMBOL(<FS_TAB2>) INDEX SY-TABIX.
            IF SY-SUBRC = 0.
              <FS_TAB2>-QTY_PCS = <FS_EXCEL_2>-QTY_PCS.
            ENDIF.
          ENDLOOP.

          PERFORM FRM_INSERT_DATA     TABLES   GT_ZPDND002MT_ALV1-TAB2[]
                                      USING    'ZPDND002MT_ALV2'
                                              LS_RANGE
                                               'X'
                                               ''.


        WHEN 'ZPDND002MT_TAB2_TT'.
********************************************************************************
**   Insert totals 2 Table 1
********************************************************************************
          PERFORM FRM_GET_EXCEL_DATA  TABLES   GT_EXCEL_DATA_TOTAL2[]
                                               GT_RANGE_CONTENT[]
                                      USING    LS_RANGE
                                               'ZPDND002MT_ALV1_HEAD'.
          PERFORM FRM_INSERT_DATA     TABLES   GT_EXCEL_DATA_TOTAL2[]
                                      USING    'ZPDND002MT_ALV1_HEAD'
                                              LS_RANGE
                                               ''
                                               ''.
        WHEN 'ZPDND002MT_TAB3'.
********************************************************************************
**  Insert Table 3 ALV 1
********************************************************************************
          PERFORM FRM_GET_EXCEL_DATA  TABLES   GT_EXCEL_DATA_3[]
                                               GT_RANGE_CONTENT[]
                                      USING    LS_RANGE
                                               'ZPDND002MT_ALV3'.

          LOOP AT GT_EXCEL_DATA_3 ASSIGNING FIELD-SYMBOL(<FS_EXCEL_3>).
            READ TABLE GT_ZPDND002MT_ALV1-TAB3 ASSIGNING FIELD-SYMBOL(<FS_TAB3>) INDEX SY-TABIX.
            IF SY-SUBRC = 0.
              <FS_TAB3>-QTY_CALC    = <FS_EXCEL_3>-QTY_CALC.
              <FS_TAB3>-ENERGY      = <FS_EXCEL_3>-ENERGY.
              <FS_TAB3>-QTY_PCS     = <FS_EXCEL_3>-QTY_PCS.
              <FS_TAB3>-TAUX_CALC   = <FS_EXCEL_3>-TAUX_CALC.
              <FS_TAB3>-TAUX_REL    = <FS_EXCEL_3>-TAUX_REL.
              <FS_TAB3>-ECART       = <FS_EXCEL_3>-ECART.
            ENDIF.
          ENDLOOP.

          PERFORM FRM_INSERT_DATA   TABLES    GT_ZPDND002MT_ALV1-TAB3[]
                                    USING     'ZPDND002MT_ALV3'
                                              LS_RANGE
                                              ''
                                              ''.

          CLEAR: GT_RANGE_CONTENT[].
          REFRESH T_RANGES.
          CLEAR T_RANGES.
          T_RANGES-NAME = LS_RANGE-NAME.
          APPEND T_RANGES.
          CALL METHOD GO_HANDLE->GET_RANGES_DATA
            EXPORTING
              ALL      = ''
            IMPORTING
              CONTENTS = GT_RANGE_CONTENT[]
              ERROR    = GV_ERROR
              RETCODE  = GV_RETCODE
            CHANGING
              RANGES   = T_RANGES[].


          PERFORM FRM_GET_EXCEL_DATA  TABLES   GT_ALV3[]
                                     GT_RANGE_CONTENT[]
                            USING    LS_RANGE
                                     'ZPDND002MT_ALV1'.

        WHEN 'ZPDND002MT_TAB4_HEADER'. " Orders
******************************************************************************
**  Insert Table ALV 2 : Orders
******************************************************************************


          PERFORM FRM_GET_EXCEL_DATA  TABLES  GT_EXCEL_DATA_4_HEAD[]
                                              GT_RANGE_CONTENT[]
                                       USING  LS_RANGE
                                              'ZPDND002MT_ALV4_C'.

          PERFORM FRM_INSERT_DATA   TABLES    GT_EXCEL_DATA_4_HEAD[]
                                    USING     'ZPDND002MT_ALV4_C'
                                              LS_RANGE
                                              ''
                                              GV_TRAINS_HIDE.

        WHEN 'ZPDND002MT_TAB_TRAIN'.

          GS_RANGE_TAB4 =  LS_RANGE.

          CLEAR GT_EXCEL_DATA_4_HEAD[].
          DATA(LS_RANGE_NEW) = LS_RANGE.
          LS_RANGE_NEW-TOP = LS_RANGE_NEW-TOP + 1.
          LS_RANGE_NEW-ROWS = LS_RANGE_NEW-ROWS - 1.

          PERFORM FRM_GET_EXCEL_DATA  TABLES  GT_EXCEL_DATA_4_HEAD[]
                        GT_RANGE_CONTENT[]
                USING  LS_RANGE_NEW
                        'ZPDND002MT_ALV4_C'.

          LOOP AT  GT_EXCEL_DATA_4_HEAD ASSIGNING FIELD-SYMBOL(<FS_EXCEL>).

            CASE SY-TABIX.

              WHEN '6'.
                READ TABLE  GT_ZPDND002MT_ALV2_C INDEX SY-TABIX ASSIGNING FIELD-SYMBOL(<FS_TT>).
                IF SY-SUBRC = 0.
                  <FS_TT> = <FS_EXCEL>.
                ENDIF.
              WHEN '7'.
                READ TABLE  GT_ZPDND002MT_ALV2_C INDEX SY-TABIX ASSIGNING <FS_TT>.
                IF SY-SUBRC = 0.
                  <FS_TT> = <FS_EXCEL>.
                ENDIF.
              WHEN '8'.
                READ TABLE  GT_ZPDND002MT_ALV2_C INDEX SY-TABIX ASSIGNING <FS_TT>.
                IF SY-SUBRC = 0.
                  <FS_TT> = <FS_EXCEL>.

                  APPEND <FS_TT> TO GT_ZPDND002MT_ALV2_AC[].

                ENDIF.
            ENDCASE.

          ENDLOOP.


          PERFORM FRM_INSERT_DATA   TABLES    GT_ZPDND002MT_ALV2_C[]
                                    USING     'ZPDND002MT_ALV4_C'
                                              LS_RANGE
                                              ' '
                                              GV_TRAINS_HIDE."here

          DATA(LV_EDIT_TOP) =  LS_RANGE-TOP + 7.


          PERFORM FRM_SET_EDITABLE    TABLES  GT_ZPDND002MT_ALV2_AC[]
                                      USING   GC_EDIT_COLS_AC
                                              'ZPDND002MT_ALV4_C'
                                               LV_EDIT_TOP.

*          ENDIF.
          PERFORM FRM_HIDE_TRAINS   USING  LS_RANGE  GV_TRAINS_HIDE.


        WHEN 'ZPDND002MT_TAB_CHART1'.
          PERFORM FRM_GET_EXCEL_DATA  TABLES  GT_EXCEL_CHART[]
                                              GT_RANGE_CONTENT[]
                                      USING  LS_RANGE
                                             'ZPDND002MT_CHART'.

          PERFORM FRM_INSERT_DATA   TABLES    GT_EXCEL_CHART[]
                                    USING     'ZPDND002MT_CHART'
                                              LS_RANGE
                                              ' '
                                              ''.

        WHEN 'ZPDND002MT_TAB_CHART2'.
          PERFORM FRM_GET_EXCEL_DATA  TABLES  GT_EXCEL_CHART[]
                                    GT_RANGE_CONTENT[]
                            USING  LS_RANGE
                                   'ZPDND002MT_CHART'.

          PERFORM FRM_INSERT_DATA   TABLES    GT_EXCEL_CHART[]
                                    USING     'ZPDND002MT_CHART'
                                              LS_RANGE
                                              ' '
                                              ''.
      ENDCASE.

    ENDLOOP.


    READ TABLE GT_ALV3 INTO GS_ALV3 INDEX 3.
    IF SY-SUBRC = 0.

      GV_QTY_FG = GS_ALV3-QTY_CONS.
      GV_UNIT = GS_ALV3-BSTME.

    ENDIF.

  ENDIF.
ENDFORM. " FRM_DISPLAY_EXCEL_DATA
*&---------------------------------------------------------------------*
*& Form FRM_EXPORT_EXCEL_FILE
*&---------------------------------------------------------------------*
FORM FRM_EXPORT_EXCEL_FILE .
  GO_DOCUMENT->SAVE_COPY_AS(
    EXPORTING
        FILE_NAME   = 'Calcul_AC'
        PROMPT_USER = 'X'
      IMPORTING
        ERROR       = GV_ERROR
        RETCODE     = GV_RETCODE
  ).
ENDFORM.
*&---------------------------------------------------------------------*
*&      Form  FRM_GET_LANG
*&---------------------------------------------------------------------*
FORM FRM_GET_LANG CHANGING CP_LANG TYPE SY-LANGU.


  DATA: LT_SHEETS            TYPE SOI_SHEETS_TABLE,
        LS_SHEETS            TYPE SOI_SHEETS,
        LV_DOCUMENT_TYPE(80) TYPE C,
        LV_DUMMYDOCUMENT     TYPE REF TO I_OI_DOCUMENT_PROXY,
        LO_DUMMYHANDLE       TYPE REF TO I_OI_SPREADSHEET..
  " Create the Container
  CALL METHOD C_OI_CONTAINER_CONTROL_CREATOR=>GET_CONTAINER_CONTROL
    IMPORTING
      CONTROL = GO_CONTROL
      ERROR   = GV_ERROR.

  CREATE OBJECT GO_CONTAINER
    EXPORTING
      CONTAINER_NAME = 'C_CONTAINER'.

  CALL METHOD GO_CONTROL->INIT_CONTROL
    EXPORTING
      R3_APPLICATION_NAME      = TEXT-R01
      INPLACE_ENABLED          = C_INPLACE
      INPLACE_SCROLL_DOCUMENTS = 'X'
      PARENT                   = GO_CONTAINER
      REGISTER_ON_CLOSE_EVENT  = 'X'
      REGISTER_ON_CUSTOM_EVENT = 'X'
      NO_FLUSH                 = 'X'
    IMPORTING
      ERROR                    = GV_ERROR.

  CHECK CP_LANG IS INITIAL.

  LV_DOCUMENT_TYPE            = C_EXCEL.

  CALL METHOD GO_CONTROL->GET_DOCUMENT_PROXY
    EXPORTING
      DOCUMENT_TYPE  = LV_DOCUMENT_TYPE
    IMPORTING
      DOCUMENT_PROXY = LV_DUMMYDOCUMENT
      ERROR          = GV_ERROR.

  CLEAR GV_RETCODE.
  CALL METHOD LV_DUMMYDOCUMENT->CREATE_DOCUMENT
    EXPORTING
      OPEN_INPLACE = C_INPLACE
    IMPORTING
      RETCODE      = GV_RETCODE
      ERROR        = GV_ERROR.
  CHECK GV_RETCODE = 'OK'.

  DATA: HAS TYPE I.
  CALL METHOD LV_DUMMYDOCUMENT->HAS_SPREADSHEET_INTERFACE
    IMPORTING
      IS_AVAILABLE = HAS.

  IF NOT HAS IS INITIAL.
    CALL METHOD LV_DUMMYDOCUMENT->GET_SPREADSHEET_INTERFACE
      IMPORTING
        SHEET_INTERFACE = LO_DUMMYHANDLE.
    CALL METHOD LO_DUMMYHANDLE->GET_SHEETS
      EXPORTING
        NO_FLUSH = SPACE
      IMPORTING
        SHEETS   = LT_SHEETS
        ERROR    = GV_ERROR
        RETCODE  = GV_RETCODE.
    LOOP AT LT_SHEETS INTO LS_SHEETS.
      CASE LS_SHEETS.
        WHEN 'Sheet1'.
          CP_LANG = 'E'.
        WHEN OTHERS.
          CP_LANG = 'F'.
      ENDCASE.
    ENDLOOP.
  ENDIF.
  CALL METHOD LV_DUMMYDOCUMENT->CLOSE_DOCUMENT.
  FREE LV_DUMMYDOCUMENT.
ENDFORM. " FRM_GET_LANG

*&---------------------------------------------------------------------*
*&      Form  FRM_CREATE_RANGES
*&---------------------------------------------------------------------*
FORM FRM_CREATE_RANGES.
  CLEAR: GT_RANGES_LIST[].
  LOOP AT GT_TAB_BORDERS ASSIGNING FIELD-SYMBOL(<FS_TAB_BORDER>).
    CLEAR GS_RANGES_LIST.
    GS_RANGES_LIST-NAME    = <FS_TAB_BORDER>-CODE_ZONE. " TABNAME
    GS_RANGES_LIST-TOP     = <FS_TAB_BORDER>-TOP.
    GS_RANGES_LIST-LEFT    = <FS_TAB_BORDER>-LEFT.
    GS_RANGES_LIST-ROWS    = <FS_TAB_BORDER>-ROWS.
    GS_RANGES_LIST-COLUMNS = <FS_TAB_BORDER>-COLS.
    APPEND GS_RANGES_LIST TO GT_RANGES_LIST.
    IF GS_RANGES_LIST-NAME = 'ZPDND002MT_TAB4_HEADER'.
      MOVE-CORRESPONDING GS_RANGES_LIST TO  GS_RANGE_HEAD.
    ENDIF.
  ENDLOOP. " LOOP AT GT_TAB_BORDERS ASSIGNING FIELD-SYMBOL(<FS_TAB_BORDER>).

  IF GO_HANDLE IS NOT INITIAL.
    GO_HANDLE->INSERT_RANGES(
    EXPORTING
      RANGES   =  GT_RANGES_LIST
    IMPORTING
      ERROR    = GV_ERROR
      RETCODE  = GV_RETCODE ).
  ENDIF.
ENDFORM.