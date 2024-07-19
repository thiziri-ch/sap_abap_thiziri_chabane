*&---------------------------------------------------------------------*
*& Report ZSDF001
*&---------------------------------------------------------------------*
* Description : Smartform program for "Sales order printout"
*&---------------------------------------------------------------------*
* Created by : Thiziri Chabane
*&---------------------------------------------------------------------*
REPORT ZSDF001.

TYPE-POOLS : ABAP, SLIS, ICON, LIST.
TABLES: ZSDF001.
DATA: P_VBELN TYPE VBELN_VA.
** Constants declaration
CONSTANTS:
  LC_ZRC   TYPE C LENGTH 6 VALUE 'ZRC',
  LC_ZNIF  TYPE C LENGTH 4 VALUE 'ZNIF',
  LC_ZNAI  TYPE C LENGTH 4 VALUE 'ZNAI',
  LC_ZNIS  TYPE C LENGTH 4 VALUE 'ZNIS',
  LC_KNTYP TYPE C LENGTH 1 VALUE 'D',
  LC_KINAK TYPE C LENGTH 1 VALUE ' ',
  LC_KSTAT TYPE C LENGTH 1 VALUE 'X',
  LC_ZRC_C TYPE C LENGTH 6 VALUE 'SAPI01',
  LC_ZIF_C TYPE C LENGTH 6 VALUE 'CGIID',
  LC_ZAI_C TYPE C LENGTH 6 VALUE 'CGIIR',
  L_SPRAS  TYPE SY-LANGU VALUE 'F'
  .
** Data declaration for local use
DATA GT_HEAD              TYPE ZSDF001_S_HEAD.
DATA GT_ITEM              TYPE STANDARD TABLE OF ZSDF001_S_ITEM.
DATA LS_ITEM              TYPE                   ZSDF001_S_ITEM.

TABLES: NAST,                          "Messages
        TNAPR,                         "Programs & Forms
        ITCPO,                         "Communicationarea for Spool
        ARC_PARAMS,                    "Archive parameters
        TOA_DARA,                      "Archive parameters
        ADDR_KEY.                      "Adressnumber for ADDRESS

DATA: RETCODE   LIKE SY-SUBRC.         "Returncode
DATA: XSCREEN(1) TYPE C.               "Output on printer or screen

FORM ENTRY USING RETURN_CODE TYPE I
                 US_SCREEN TYPE C.

  XSCREEN = US_SCREEN.
  IF NAST-OBJKY IS NOT INITIAL.
    PERFORM FRM_PRINT_SO USING NAST TNAPR. "IN PROGRAM ZHSCF003 IF FOUND
    IF RETCODE NE 0.
      RETURN_CODE = 1.
    ELSE.
      RETURN_CODE = 0.
    ENDIF.
  ENDIF.

ENDFORM.                    "ENTRY

FORM FRM_PRINT_SO USING PS_NAST TYPE NAST PS_TNAPR TYPE TNAPR .
  DATA : XT685T        TYPE TABLE OF T685T,
         LT_HEADER     TYPE TABLE OF BAPISDHD WITH HEADER LINE,
*         LT_ITEMS      TYPE TABLE OF BAPISDITBOS WITH HEADER LINE,
         LT_ITEMS      TYPE TABLE OF BAPISDIT WITH HEADER LINE,
         LT_PARTNERS   TYPE TABLE OF BAPISDPART WITH HEADER LINE,
         LT_BUSINESS   TYPE TABLE OF BAPISDBUSI WITH HEADER LINE,
         LT_CONDITIONS TYPE TABLE OF BAPISDCOND WITH HEADER LINE,
         LT_ADDRESS    TYPE TABLE OF BAPISDCOAD WITH HEADER LINE,
         LT_ID         TYPE TABLE OF BAPIBUS1006_ID_DETAILS WITH HEADER LINE,
         LT_FLOW       TYPE TABLE OF BAPISDFLOW WITH HEADER LINE,
         STRUCT        TYPE TABLE OF T001 WITH HEADER LINE,
         T_T001Z       TYPE TABLE OF T001Z WITH HEADER LINE,
         E_ADRC        TYPE TABLE OF ADRC WITH HEADER LINE,
         E_ADR6        TYPE TABLE OF ADR6,
         LT_RETURN     TYPE TABLE OF BAPIRET2 WITH HEADER LINE.

  DATA : LV_TAX_PERCENT TYPE CHAR20,
         COND_TYPE      TYPE KSCHA,
         COND_TYPE_TXT  TYPE VTEXT,
         ES_T005T       TYPE  T005T,
         E_T052         TYPE T052,
         E_T685T        TYPE  T685T,
         TVBUR_I        TYPE TVBUR,
         Z_ZSDF001      TYPE ZSDF001.

  DATA : LS_VIEW      LIKE ORDER_VIEW,
         LF_MEMREAD   TYPE  ORDER_READ-MEM_ACCESS,
         LT_SALESDOCS TYPE STANDARD TABLE OF SALES_KEY,
         LS_SALESDOCS LIKE SALES_KEY,
         LF_SALESDOC  LIKE BAPIVBELN-VBELN.
  DATA  LV_OUTPUT_CAT TYPE CHAR01.
**********************************************************************
** Read sales document data
**********************************************************************
  CLEAR GT_ITEM.
  CLEAR GT_HEAD.
**Sales order number
  P_VBELN = PS_NAST-OBJKY.

* read only from SD memory
  MOVE 'A' TO LF_MEMREAD.
  MOVE 'X' TO LS_VIEW-HEADER.
  MOVE 'X' TO LS_VIEW-ITEM.
  MOVE 'X' TO LS_VIEW-BUSINESS.
  MOVE 'X' TO LS_VIEW-PARTNER.
  MOVE 'X' TO LS_VIEW-ADDRESS.
  MOVE 'X' TO LS_VIEW-SDCOND.
  MOVE 'X' TO LS_VIEW-FLOW.
  MOVE P_VBELN  TO LF_SALESDOC.
* call conversion exits for salesorder number
  CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
    EXPORTING
      INPUT  = LF_SALESDOC
    IMPORTING
      OUTPUT = LF_SALESDOC.

  MOVE LF_SALESDOC TO LS_SALESDOCS.
  APPEND LS_SALESDOCS TO LT_SALESDOCS.

  CALL FUNCTION 'BAPISDORDER_GETDETAILEDLIST'
    EXPORTING
      I_BAPI_VIEW          = LS_VIEW
      I_MEMORY_READ        = LF_MEMREAD
    TABLES
      SALES_DOCUMENTS      = LT_SALESDOCS
      ORDER_HEADERS_OUT    = LT_HEADER
      ORDER_ITEMS_OUT      = LT_ITEMS
      ORDER_BUSINESS_OUT   = LT_BUSINESS
      ORDER_PARTNERS_OUT   = LT_PARTNERS
      ORDER_ADDRESS_OUT    = LT_ADDRESS
      ORDER_CONDITIONS_OUT = LT_CONDITIONS
      ORDER_FLOWS_OUT      = LT_FLOW.
  IF SY-SUBRC <> 0.
    MESSAGE ID SY-MSGID TYPE 'E' NUMBER SY-MSGNO
                  WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4 DISPLAY LIKE SY-MSGTY.
  ELSEIF LT_HEADER IS NOT INITIAL.

    GT_HEAD-SALES_DOC     = LT_HEADER-DOC_NUMBER.
    GT_HEAD-DOC_TYPE      = LT_HEADER-DOC_TYPE.
    GT_HEAD-DOC_CONDITION = LT_HEADER-CONDITIONS.
    GT_HEAD-DOC_DATE      = LT_HEADER-REC_DATE.
    " REFERENCE DOCUMENT V1.5
    IF LT_HEADER-DOC_TYPE EQ 'ZDR' OR LT_HEADER-DOC_TYPE EQ 'ZCR'.
      LOOP AT LT_FLOW WHERE DOC_CAT_SD EQ 'M'.
        MOVE LT_FLOW-PRECSDDOC TO GT_HEAD-REFERENCE_DOC.
      ENDLOOP.
    ELSE.
      CLEAR GT_HEAD-REFERENCE_DOC.
    ENDIF.

** Sales Document Title
**V1.7
    CASE LT_HEADER-DOC_TYPE.
      WHEN 'ZCR' OR 'ZDR' OR 'ZRE'.
        SELECT SINGLE DOC_TYPE_TXT
         FROM ZSDF001
         INTO  GT_HEAD-DOC_TYPE_TXT_F
         WHERE LANGUAGE  EQ L_SPRAS
          AND SALES_ORG EQ ' '
          AND DOC_TYPE  EQ LT_HEADER-DOC_TYPE.
        IF SY-SUBRC <> 0.
          CLEAR  GT_HEAD-DOC_TYPE_TXT_F.
          MESSAGE ID 'ZSDF001' TYPE 'W' NUMBER '001' WITH  LT_HEADER-DOC_TYPE. "Title not available
        ENDIF.

      WHEN 'ZFD'.
        SELECT SINGLE DOC_TYPE_TXT
          FROM ZSDF001
          INTO  GT_HEAD-DOC_TYPE_TXT_F
          WHERE LANGUAGE  EQ L_SPRAS
           AND SALES_ORG EQ ' '
           AND DOC_TYPE  EQ LT_HEADER-DOC_TYPE.
        IF SY-SUBRC <> 0.
          MESSAGE ID 'ZSDF001' TYPE 'W' NUMBER '001' WITH LT_HEADER-DOC_TYPE. "Title not available
          GT_HEAD-DOC_TYPE_TXT_F = 'COMMANDE DE VENTE'.

        ENDIF.

    ENDCASE.
**********************************************************************
** read customer's name | addresses | identification numbers (NIF | NRC | NAI | NIS)
**********************************************************************
    READ TABLE LT_PARTNERS
             WITH KEY PARTN_ROLE = 'RE' "Bill to party
             INTO LT_PARTNERS.
    IF SY-SUBRC = 0.
      GT_HEAD-CUS_NUMBER      = LT_PARTNERS-CUSTOMER.

**Get customer address details
      IF LT_ADDRESS IS  NOT INITIAL.
        READ TABLE LT_ADDRESS
          WITH KEY ADDRESS = LT_PARTNERS-ADDRESS
          INTO LT_ADDRESS.
        IF SY-SUBRC <> 0.
          CLEAR LT_ADDRESS.
        ELSE.
          GT_HEAD-CUS_NAME1       = LT_ADDRESS-NAME.
          GT_HEAD-CUS_NAME2       = LT_ADDRESS-NAME_2.
          GT_HEAD-CUS_NAME3       = LT_ADDRESS-NAME_3.
          GT_HEAD-CUS_CITY1       = LT_ADDRESS-CITY.
          GT_HEAD-CUS_POST_CODE1  = LT_ADDRESS-POSTL_CODE.
          GT_HEAD-CUS_COUNTRY     = LT_ADDRESS-COUNTRY.
          GT_HEAD-CUS_STREET      = LT_ADDRESS-STREET.
        ENDIF.
      ENDIF.
** Get Identification numbers
      CLEAR LT_ID.
      CALL FUNCTION 'BAPI_IDENTIFICATIONDETAILS_GET'
        EXPORTING
          BUSINESSPARTNER      = LT_PARTNERS-CUSTOMER
        TABLES
          IDENTIFICATIONDETAIL = LT_ID
          RETURN               = LT_RETURN.
      IF SY-SUBRC <> 0.
        CLEAR : LT_ID , LT_RETURN.
      ELSE.
        IF LT_RETURN IS NOT INITIAL.
          " Error reading identification details
        ELSE.
          IF LT_ID IS NOT INITIAL.
            LOOP AT LT_ID.
              CASE LT_ID-IDENTIFICATIONTYPE.
                WHEN LC_ZRC.
                  GT_HEAD-ID_NUMBER_RC  = LT_ID-IDENTIFICATIONNUMBER.
                WHEN LC_ZNIF.
                  GT_HEAD-ID_NUMBER_IF = LT_ID-IDENTIFICATIONNUMBER.
                WHEN LC_ZNAI.
                  GT_HEAD-ID_NUMBER_AI  = LT_ID-IDENTIFICATIONNUMBER.
                WHEN LC_ZNIS.
                  GT_HEAD-ID_NUMBER_IS  = LT_ID-IDENTIFICATIONNUMBER.
                WHEN OTHERS.
              ENDCASE.
            ENDLOOP.
          ENDIF.
        ENDIF.
      ENDIF.
    ENDIF.
**********************************************************************
** Comapany code, name, address, identification numbers , tel, fax, email
**********************************************************************
    " Add condition to check company code is not initial
    CHECK LT_HEADER-COMP_CODE IS NOT INITIAL.
    IF SY-SUBRC = 0.
      GT_HEAD-COMPANY_CODE  = LT_HEADER-COMP_CODE.

      SELECT SINGLE  BUTXT
        FROM  T001
        INTO  GT_HEAD-COMPANY_NAME
        WHERE BUKRS = LT_HEADER-COMP_CODE
        AND   SPRAS = L_SPRAS.
      IF SY-SUBRC <> 0.
        CLEAR GT_HEAD-COMPANY_NAME.
      ENDIF.

      CALL FUNCTION 'T001_READ'
        EXPORTING
          BUKRS    = LT_HEADER-COMP_CODE
        IMPORTING
          STRUCT   = STRUCT
        EXCEPTIONS
          NO_ENTRY = 1
          OTHERS   = 2.
      IF SY-SUBRC <> 0 .
        CLEAR STRUCT.
      ELSE.
        IF STRUCT IS NOT INITIAL.
          CLEAR E_ADRC.
          CALL FUNCTION 'RTP_US_DB_ADRC_READ'
            EXPORTING
              I_ADDRESS_NUMBER = STRUCT-ADRNR
            IMPORTING
              E_ADRC           = E_ADRC
            EXCEPTIONS
              NOT_FOUND        = 1
              OTHERS           = 2.
          IF SY-SUBRC <> 0.
            CLEAR E_ADRC.
          ELSE.
            GT_HEAD-COMPANY_NAME1         = E_ADRC-NAME1.
            GT_HEAD-COMPANY_NAME2         = E_ADRC-NAME2.
            GT_HEAD-COMPANY_NAME3         = E_ADRC-NAME3.
            GT_HEAD-COMPANY_CITY1         = E_ADRC-CITY1.
            GT_HEAD-COMPANY_POST_CODE1    = E_ADRC-POST_CODE1.
            GT_HEAD-COMPANY_COUNTRY       = E_ADRC-COUNTRY.
            GT_HEAD-COMPANY_STREET        = E_ADRC-STREET.
            GT_HEAD-COMPANY_LOCATION      = E_ADRC-LOCATION.
            GT_HEAD-COMPANY_TEL_NUMBER    = E_ADRC-TEL_NUMBER.
            GT_HEAD-COMPANY_FAX_NUMBER    = E_ADRC-FAX_NUMBER.
          ENDIF.
*            ENDIF.
** Company country name
          CLEAR ES_T005T.
          CALL FUNCTION 'READ_T005T'
            EXPORTING
              IC_SPRAS = L_SPRAS
              IC_LAND1 = E_ADRC-COUNTRY
            IMPORTING
              ES_T005T = ES_T005T.
          IF SY-SUBRC <> 0.
            CLEAR ES_T005T.
          ELSE.
            GT_HEAD-COMPANY_COUNTRY_NAME  = ES_T005T-LANDX.
          ENDIF.

**Company email address
          CLEAR E_ADR6.
          CALL FUNCTION 'ADDR_SELECT_ADR6_SINGLE'
            EXPORTING
              ADDRNUMBER          = STRUCT-ADRNR
            TABLES
              ET_ADR6             = E_ADR6
            EXCEPTIONS
              COMM_DATA_NOT_EXIST = 1
              PARAMETER_ERROR     = 2
              INTERNAL_ERROR      = 3
              OTHERS              = 4.
          IF  SY-SUBRC <> 0.
*            MESSAGE ID SY-MSGID TYPE 'W' NUMBER SY-MSGNO
*                          WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4 DISPLAY LIKE SY-MSGTY.
            CLEAR E_ADR6.
          ELSE.
            IF E_ADR6 IS NOT INITIAL.
              GT_HEAD-SMTP_ADDR  = E_ADR6[ 1 ]-SMTP_ADDR.
            ENDIF.
          ENDIF.
** Company NIF | NRC | NAI
          SELECT * FROM T001Z
            INTO TABLE T_T001Z
            WHERE  BUKRS = LT_HEADER-COMP_CODE.

          IF SY-SUBRC = 0.
            LOOP AT T_T001Z .
              CASE T_T001Z-PARTY.
                WHEN LC_ZIF_C .
                  GT_HEAD-COMPANY_ID_NUMBER_IF = T_T001Z-PAVAL.
                WHEN LC_ZAI_C .
                  GT_HEAD-COMPANY_ID_NUMBER_AI = T_T001Z-PAVAL.
                WHEN LC_ZRC_C .
                  GT_HEAD-COMPANY_ID_NUMBER_RC = T_T001Z-PAVAL.
              ENDCASE.
            ENDLOOP.
          ENDIF.
        ENDIF.
      ENDIF.
    ENDIF.

********************************************************
**Sales Office city
********************************************************
    GT_HEAD-SALES_OFFICE = LT_HEADER-SALES_OFF.

    CLEAR TVBUR_I.
    CALL FUNCTION 'ISP_SELECT_SINGLE_TVBUR'
      EXPORTING
        VKBUR          = LT_HEADER-SALES_OFF
*       MSGTY          = '*'
      IMPORTING
*       meldung        = meldung
        TVBUR_I        = TVBUR_I
      EXCEPTIONS
        NO_ENTRY_FOUND = 1
        OTHERS         = 2.
    IF SY-SUBRC <> 0.
      CLEAR TVBUR_I.
    ENDIF.
    IF TVBUR_I IS NOT INITIAL.
      CLEAR E_ADRC.
      CALL FUNCTION 'RTP_US_DB_ADRC_READ'
        EXPORTING
          I_ADDRESS_NUMBER = TVBUR_I-ADRNR
        IMPORTING
          E_ADRC           = E_ADRC
        EXCEPTIONS
          NOT_FOUND        = 1
          OTHERS           = 2.
      IF SY-SUBRC <> 0 .
        CLEAR E_ADRC.
      ELSE.
        GT_HEAD-SALES_OFFICE_CITY1 = E_ADRC-CITY1.
      ENDIF.
    ENDIF.
********************************************************************
** Sales document items
********************************************************************
    DATA: SUM TYPE ZSDF001_S_HEAD-TAX_AMOUNT .
*        VALUE IS INITIAL.
    CLEAR SUM.

    LOOP AT LT_ITEMS.
      "      do 20 times.
      "[] INTO LT_ITEMS.
      LS_ITEM-MATERIAL_NUM    = LT_ITEMS-MATERIAL.
      LS_ITEM-SALES_DOC_ITEM  = LT_ITEMS-ITM_NUMBER.
      LS_ITEM-MATERIAL_TXT    = LT_ITEMS-SHORT_TEXT.
      LS_ITEM-SALES_UNIT      = LT_ITEMS-SALES_UNIT.
      LS_ITEM-NET_PRICE       = LT_ITEMS-NET_PRICE.
      LS_ITEM-NET_VALUE       = LT_ITEMS-NET_VALUE.
      LS_ITEM-TAX_AMOUNT      = LT_ITEMS-TAX_AMOUNT.
      SUM = SUM + LS_ITEM-TAX_AMOUNT.
**Quantity
      IF LT_HEADER-DOC_TYPE EQ 'ZCR' OR LT_HEADER-DOC_TYPE EQ 'ZDR'.
        LS_ITEM-ORDER_QTY       = LT_ITEMS-TARGET_QTY.
      ELSE.
        LS_ITEM-ORDER_QTY       = LT_ITEMS-REQ_QTY.
      ENDIF.


      IF LT_CONDITIONS IS NOT INITIAL.
        READ TABLE LT_CONDITIONS
        WITH KEY
          CONDTYPE = LC_KNTYP
*          APPLICATIO = 'TX'
          CONDISACTI = LC_KINAK
          STAT_CON =  LC_KSTAT
          GROUPCOND = 'X'
*          CONDCLASS = 'D'
          INTO LT_CONDITIONS .

        IF SY-SUBRC EQ 0.
          COND_TYPE = LT_CONDITIONS-COND_TYPE.
          WRITE: LT_CONDITIONS-COND_VALUE TO LV_TAX_PERCENT EXPONENT 0 DECIMALS 2.
          CONDENSE LV_TAX_PERCENT.
          CONCATENATE LV_TAX_PERCENT(2) SYM_FILLED_SQUARE INTO LS_ITEM-TAX_PERCENT.
        ENDIF.

      ENDIF.
**Add lines
      APPEND LS_ITEM TO GT_ITEM.
      CLEAR  LS_ITEM.
*       enddo.
    ENDLOOP.

** Totals amount
    GT_HEAD-TAX_AMOUNT = SUM.
    GT_HEAD-NET_VALUE = LT_HEADER-NET_VAL_HD.
    GT_HEAD-TOTAL_VALUE = GT_HEAD-TAX_AMOUNT + GT_HEAD-NET_VALUE .

**Payment term
    READ  TABLE LT_BUSINESS WITH KEY  ITM_NUMBER = '000000' INTO LT_BUSINESS.
*      GT_HEAD-PAYMENT_TERMS_CODE  =  LT_BUSINESS-PMNTTRMS.
    SELECT SINGLE VTEXT
        FROM TVZBT
        INTO GT_HEAD-PAYMENT_TERMS
        WHERE SPRAS = L_SPRAS
        AND   ZTERM =  LT_BUSINESS-PMNTTRMS.
    IF SY-SUBRC <> 0 .
      CLEAR GT_HEAD-PAYMENT_TERMS.
    ELSE.
      CALL FUNCTION 'WCB_READ_T052'
        EXPORTING
          I_ZTERM   = LT_BUSINESS-PMNTTRMS
        IMPORTING
          E_T052    = E_T052
        EXCEPTIONS
          NOT_FOUND = 1
          OTHERS    = 2.
      IF SY-SUBRC <> 0.
        CLEAR E_T052.
      ELSE.
        IF E_T052 IS NOT INITIAL.
          GT_HEAD-NO_OF_DAYS = E_T052-ZTAG1.
          GT_HEAD-DUE_DATE = E_T052-ZTAG1 + LT_BUSINESS-BILL_DATE.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDIF.
  IF GT_HEAD IS NOT INITIAL.
    PERFORM CALL_SMARTFORM USING PS_TNAPR-SFORM.
  ELSE.
    EXIT.
  ENDIF.


ENDFORM.
*&---------------------------------------------------------------------*
*& Form CALL_SMARTFORM
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM CALL_SMARTFORM USING PS_TNAPR-SFORM.

***CALL SMARTFORM
  DATA : LS_CONTROL_PARAMETERS TYPE SSFCTRLOP,                       "Smart Forms: Control structure
         LS_OPTION             TYPE SSFCRESOP,                         "Smart Forms: Return value at start of form printing
         LV_FORMNAME           TYPE TDSFNAME,                          "Smart Forms: Form Name
         LV_FM_NAME            TYPE RS38L_FNAM,                        "Name of Function Module
         LS_OUTPUT_OPTIONS     TYPE SSFCRESOP,
         LS_JOB_OUTPUT_INFO	   TYPE	SSFCRESCL.
  CLEAR: LS_OPTION.
  LV_FORMNAME = PS_TNAPR-SFORM.
  CALL FUNCTION 'SSF_FUNCTION_MODULE_NAME'
    EXPORTING
      FORMNAME = LV_FORMNAME
    IMPORTING
      FM_NAME  = LV_FM_NAME.

  CLEAR LS_CONTROL_PARAMETERS.

  LS_CONTROL_PARAMETERS-NO_DIALOG  = ABAP_OFF.
  LS_CONTROL_PARAMETERS-PREVIEW    = ABAP_ON.
  LS_CONTROL_PARAMETERS-NO_OPEN    = ABAP_ON.
  LS_CONTROL_PARAMETERS-NO_CLOSE   = ABAP_ON.
  LS_CONTROL_PARAMETERS-LANGU      = SY-LANGU.

  CALL FUNCTION 'SSF_OPEN'
    EXPORTING
      CONTROL_PARAMETERS = LS_CONTROL_PARAMETERS
      USER_SETTINGS      = ' '
    IMPORTING
      JOB_OUTPUT_OPTIONS = LS_OPTION
    EXCEPTIONS
      FORMATTING_ERROR   = 1
      INTERNAL_ERROR     = 2
      SEND_ERROR         = 3
      USER_CANCELED      = 4
      OTHERS             = 5.
  IF SY-SUBRC <> 0.
    MESSAGE ID SY-MSGID TYPE 'E' NUMBER SY-MSGNO
                    WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4 DISPLAY LIKE SY-MSGTY.
  ENDIF.

  IF SY-SUBRC <> 0.
*Implement suitable error handling here
  ENDIF.

  CALL FUNCTION LV_FM_NAME
    EXPORTING
      CONTROL_PARAMETERS = LS_CONTROL_PARAMETERS
      USER_SETTINGS      = 'X'
      SF_HEAD            = GT_HEAD
    IMPORTING
*     DOCUMENT_OUTPUT_INFO       =
*     JOB_OUTPUT_INFO    =
      JOB_OUTPUT_OPTIONS = LS_OPTION
    TABLES
      SF_ITEM            = GT_ITEM
    EXCEPTIONS
      FORMATTING_ERROR   = 1
      INTERNAL_ERROR     = 2
      SEND_ERROR         = 3
      USER_CANCELED      = 4
      OTHERS             = 5.
  IF SY-SUBRC <> 0.
    MESSAGE ID SY-MSGID TYPE 'E' NUMBER SY-MSGNO
                  WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4 DISPLAY LIKE SY-MSGTY.
  ENDIF.

  CALL FUNCTION 'SSF_CLOSE'
*   IMPORTING
*     JOB_OUTPUT_INFO        = LS_JOB_OUTPUT_INFO
    EXCEPTIONS
      FORMATTING_ERROR = 1
      INTERNAL_ERROR   = 2
      SEND_ERROR       = 3
      OTHERS           = 4.
  IF SY-SUBRC <> 0.
* Implement suitable error handling here
    MESSAGE ID SY-MSGID TYPE 'E' NUMBER SY-MSGNO
                 WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4 DISPLAY LIKE SY-MSGTY.
  ENDIF.
ENDFORM.