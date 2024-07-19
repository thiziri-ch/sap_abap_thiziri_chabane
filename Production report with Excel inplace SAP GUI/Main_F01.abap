*&---------------------------------------------------------------------*
*& Include          ZPDND002M_F01
*&---------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*& Form LEAVE_PROGRAM
*&---------------------------------------------------------------------*
FORM LEAVE_PROGRAM .
  GO_DOCUMENT->CLOSE_DOCUMENT(  ).
  FREE  GO_DOCUMENT.
  LEAVE PROGRAM.
ENDFORM. " LEAVE_PROGRAM
*&---------------------------------------------------------------------*
*& Form LEAVE_SCREEN
*&---------------------------------------------------------------------*
FORM LEAVE_SCREEN.
  GO_DOCUMENT->CLOSE_DOCUMENT(  ).
  FREE  GO_DOCUMENT.
  LEAVE TO SCREEN 0.
ENDFORM. " LEAVE_SCREEEN
*&---------------------------------------------------------------------*
*& Form FRM_REFRESH_DATE
*&---------------------------------------------------------------------*
FORM FRM_REFRESH_DATE.

  IF S_BUDAT-HIGH IS INITIAL.
    S_BUDAT-HIGH = S_BUDAT-LOW.
    S_BUDAT-SIGN = 'I'.
    S_BUDAT-OPTION = 'BT'.
    REFRESH S_BUDAT.
    APPEND S_BUDAT.
  ENDIF.
ENDFORM. " FRM_REFRESH_DATE
*&---------------------------------------------------------------------*
*& Form FRM_SET_STATUS
*&---------------------------------------------------------------------*
FORM FRM_SET_STATUS.

  SET PF-STATUS 'STATUS_0100'.

  IF SY-LANGU = 'F'.
    SET TITLEBAR 'TITLE_100_FR'.
  ELSE.
    SET TITLEBAR 'TITLE_100_EN'.
  ENDIF.

ENDFORM.
*&---------------------------------------------------------------------*
*& Form FRM_ACTIVE_BUTTON
*&---------------------------------------------------------------------*
FORM FRM_ACTIVE_BUTTON.
  CHECK GV_FIRST_SET = ABAP_FALSE.
  LOOP AT SCREEN.
    CASE SCREEN-GROUP1.

      WHEN  'GP1'.
        PERFORM FRM_CHECK_POSTING_AUTH  USING 'AC'
                                       P_WERKS
                                       'ZI01'
                                CHANGING
                                        GV_POSTING_AUTHORIZED.
        IF GV_POSTING_AUTHORIZED = ABAP_FALSE.
          SCREEN-ACTIVE =  '0'.
        ENDIF.

      WHEN 'GP2'.
        PERFORM FRM_CHECK_POSTING_AUTH  USING 'FG'
                                             P_WERKS
                                             ''
                                      CHANGING
                                              GV_POSTING_AUTHORIZED.
        IF GV_POSTING_AUTHORIZED = ABAP_FALSE.
          SCREEN-ACTIVE =  '0'.
        ENDIF.

    ENDCASE.
    MODIFY SCREEN.
  ENDLOOP.
ENDFORM. "FRM_ACTIVE_BUTTON
*&---------------------------------------------------------------------*
*& Form FRM_INIT_VALUES
*&---------------------------------------------------------------------*
FORM FRM_INIT_PBO_VALUES.

  DATA: LT_BORDERS     TYPE TABLE OF TY_BORDER,
        LS_TAB_BORDERS TYPE TY_TAB_BORDERS.

  PERFORM FRM_GET_LANG CHANGING GV_EXCEL_LANG.


  " Set date in selection screen
  S_BUDAT-LOW = SY-DATUM - 1.
  S_BUDAT-HIGH = SY-DATUM - 1.
  S_BUDAT-SIGN = 'I'.
  S_BUDAT-OPTION = 'BT'.
  APPEND S_BUDAT.

  " Select order types from TVARV
  SELECT *
   FROM TVARVC
   INTO TABLE @DATA(LT_AUART)
  WHERE NAME = 'ZPDND002MT_AUART'
    AND TYPE = 'S'.
  IF SY-SUBRC <> 0.
    MESSAGE E015(ZPDN002) DISPLAY LIKE 'E'.
  ELSE.
    GT_AUART-SIGN = 'I'.
    GT_AUART-OPTION = 'EQ'.
    LOOP AT LT_AUART INTO DATA(LS_AUART).
      GT_AUART-LOW = LS_AUART-LOW.
      APPEND GT_AUART.
    ENDLOOP.
  ENDIF. " IF SY-SUBRC <> 0.

  ""HM

  SELECT *
   FROM TVARVC
   INTO TABLE @DATA(LT_RANGES_INFO)
  WHERE NAME LIKE 'ZPDND002MT_TAB%'
    AND LOW  LIKE 'GT_ZPDND002MT_%'
    AND TYPE = 'S'.
  IF SY-SUBRC <> 0.
    MESSAGE E002(ZPDN002) DISPLAY LIKE 'W'.
  ENDIF. " IF SY-SUBRC <> 0.

  LOOP AT LT_RANGES_INFO ASSIGNING FIELD-SYMBOL(<FS_RANGE_INFO>).
    CLEAR: LT_BORDERS[].
    SPLIT <FS_RANGE_INFO>-HIGH AT '_' INTO TABLE LT_BORDERS.
    DATA(LV_LINES) = LINES( LT_BORDERS[] ).

    IF LV_LINES <> 4.
      MESSAGE E003(ZPDN002) WITH <FS_RANGE_INFO>-NAME DISPLAY LIKE 'I'.
    ENDIF. " IF LV_LINES <> 4.

    LS_TAB_BORDERS-TABNAME   = <FS_RANGE_INFO>-NAME+11.
    LS_TAB_BORDERS-STRUCTURE = <FS_RANGE_INFO>-LOW+3.
    LS_TAB_BORDERS-CODE_ZONE =  LS_TAB_BORDERS-STRUCTURE(10) && '_' && <FS_RANGE_INFO>-NAME+11.
    LOOP AT LT_BORDERS ASSIGNING FIELD-SYMBOL(<FS_BORDERS>).
      CASE SY-TABIX.
        WHEN 1.
          LS_TAB_BORDERS-TOP  = <FS_BORDERS>-VAL.
        WHEN 2.
          LS_TAB_BORDERS-LEFT = <FS_BORDERS>-VAL.
        WHEN 3.
          LS_TAB_BORDERS-ROWS = <FS_BORDERS>-VAL.
        WHEN 4.
          LS_TAB_BORDERS-COLS = <FS_BORDERS>-VAL.
      ENDCASE. " CASE SY-TABIX.
    ENDLOOP. " LOOP AT LT_BORDERS ASSIGNING FIELD-SYMBOL(<FS_BORDERS>).
    APPEND LS_TAB_BORDERS TO GT_TAB_BORDERS.
  ENDLOOP. " LOOP AT LT_RANGES_INFO ASSIGNING FIELD-SYMBOL(<FS_RANGE_INFO>).
*  ENDIF.
ENDFORM. " FRM_INIT_VALUES

*&---------------------------------------------------------------------*
*& Form FRM_GET_ORDERS_DATA
*&---------------------------------------------------------------------*
FORM FRM_GET_ORDERS_DATA CHANGING CP_CHECK_OK TYPE ABAP_BOOL.
  DATA: LS_OUT      TYPE ZPDND002MT_ALV4,
        LV_TRAIN(3) TYPE C,
        LV_NUM      TYPE I VALUE 1,
        LV_ENERGY   TYPE  ZPCS23_6,
        LV_QUANTITY TYPE ZPDND002MT_ALV1-QTY_IN,
        LV_ORDER1   TYPE AUFNR,
        LV_ORDER2   TYPE AUFNR.

  PERFORM GET_ORDERS_LIST TABLES GT_ORDERS
                                 GT_OUT_ORDERS[].

  PERFORM FRM_CHECK_ORDERS TABLES GT_ORDERS
                       CHANGING CP_CHECK_OK
                                LV_ORDER1
                                LV_ORDER2.
  IF CP_CHECK_OK = ABAP_FALSE.
    MESSAGE S016(ZPDN002) WITH LV_ORDER1 LV_ORDER2 DISPLAY LIKE 'E'.
    EXIT.
  ENDIF.

  DATA: LS_RANGE         TYPE  SOI_FULL_RANGE_ITEM,
        GT_RANGE_CONTENT TYPE SOI_GENERIC_TABLE.

  CLEAR: GT_RANGE_CONTENT[].
  REFRESH T_RANGES.
  CLEAR T_RANGES.
  MOVE-CORRESPONDING GS_RANGE_HEAD TO T_RANGES.
  APPEND T_RANGES.
  IF GO_HANDLE IS NOT INITIAL.
    CALL METHOD GO_HANDLE->GET_RANGES_DATA
      EXPORTING
        ALL      = ''
      IMPORTING
        CONTENTS = GT_RANGE_CONTENT[]
        ERROR    = GV_ERROR
        RETCODE  = GV_RETCODE
      CHANGING
        RANGES   = T_RANGES[].
  ENDIF.
  PERFORM FRM_GET_EXCEL_DATA  TABLES  GT_EXCEL_DATA_4_HEAD[]
                            GT_RANGE_CONTENT[]
                    USING  GS_RANGE_HEAD
                            'ZPDND002MT_ALV4_C'.


  GT_ZPDND002MT_ALV2_C[] = GT_EXCEL_DATA_4_HEAD[].

  READ TABLE GT_ZPDND002MT_ALV2_C ASSIGNING FIELD-SYMBOL(<FS_HEAD_TRAIN>) INDEX 1.
  IF SY-SUBRC = 0.
    <FS_HEAD_TRAIN>-T1 = <FS_HEAD_TRAIN>-T1 && ` ` && '(Available)'(016).
    <FS_HEAD_TRAIN>-T2 = <FS_HEAD_TRAIN>-T2 && ` ` && '(Available)'(016).
    <FS_HEAD_TRAIN>-T3 = <FS_HEAD_TRAIN>-T3 && ` ` && '(Available)'(016).
    <FS_HEAD_TRAIN>-T4 = <FS_HEAD_TRAIN>-T4 && ` ` && '(Available)'(016).
    <FS_HEAD_TRAIN>-T5 = <FS_HEAD_TRAIN>-T5 && ` ` && '(Available)'(016).
    <FS_HEAD_TRAIN>-T6 = <FS_HEAD_TRAIN>-T6 && ` ` && '(Available)'(016).
    <FS_HEAD_TRAIN>-T7 = <FS_HEAD_TRAIN>-T7 && ` ` && '(Available)'(016).
    <FS_HEAD_TRAIN>-T8 = <FS_HEAD_TRAIN>-T8 && ` ` && '(Available)'(016).
    <FS_HEAD_TRAIN>-T9 = <FS_HEAD_TRAIN>-T9 && ` ` && '(Available)'(016).
  ENDIF.

  DO 7 TIMES.
    APPEND INITIAL LINE TO GT_ZPDND002MT_ALV2_C.
  ENDDO.


  LOOP AT GT_ORDERS INTO DATA(LS_TRAINS).
    DATA(LV_FIELD) = LS_TRAINS-TRAIN.

    IF LS_TRAINS-STATUS = 'WORKING'.
      DATA(LV_FIELD2) = LV_FIELD && ` ` && '(Working)'(017).
    ELSEIF LS_TRAINS-STATUS = 'SHUTDOWN'.
      LV_FIELD2 = LV_FIELD && ` ` && '(Shutdown)'(018).
    ENDIF.

    READ TABLE GT_ZPDND002MT_ALV2_C INDEX 1 ASSIGNING FIELD-SYMBOL(<HEADER>).
    IF SY-SUBRC = 0.
      ASSIGN COMPONENT LV_FIELD OF STRUCTURE <HEADER> TO FIELD-SYMBOL(<FS_FIELD_2>).
      <FS_FIELD_2> = LV_FIELD2.
    ENDIF.


    READ TABLE GT_ZPDND002MT_ALV2_C INDEX 2 ASSIGNING FIELD-SYMBOL(<FS>). "HERE
    IF SY-SUBRC = 0.
      ASSIGN COMPONENT LV_FIELD OF STRUCTURE <FS> TO FIELD-SYMBOL(<FS_FIELD>).


      <FS_FIELD> = LS_TRAINS-AUFNR.

      " Cost center of orders
      SELECT SINGLE ZFGCC FROM @GT_ZPDNAC AS A
        WHERE A~ZTYPE = 4
        INTO @<FS>-UTILITES.
      GV_KOSTL_AC = <FS>-UTILITES.
    ENDIF.

    "Energy in
    PERFORM FRM_GET_ENERGY USING    '261'
                                    '262'
                                    GV_FIELD
                                    GV_MATNR_GN
                                    LS_TRAINS-AUFNR
                       	   CHANGING LV_ENERGY.

    READ TABLE GT_ZPDND002MT_ALV2_C INDEX 3 ASSIGNING <FS>. "HERE
    IF SY-SUBRC = 0.
      ASSIGN COMPONENT LV_FIELD OF STRUCTURE <FS> TO <FS_FIELD>.
      <FS_FIELD> = LV_ENERGY.
      REPLACE '.' WITH ',' INTO <FS_FIELD>.
    ENDIF.

    " Energy out
    PERFORM FRM_GET_ENERGY USING    '101'
                                    '102'
                                    GV_FIELD
                                    '' "MATNR
                                    LS_TRAINS-AUFNR
               	           CHANGING LV_ENERGY.

    READ TABLE GT_ZPDND002MT_ALV2_C INDEX 4 ASSIGNING <FS>. "HERE
    IF SY-SUBRC = 0.
      ASSIGN COMPONENT LV_FIELD OF STRUCTURE <FS> TO <FS_FIELD>.
      <FS_FIELD> = LV_ENERGY.
      REPLACE '.' WITH ',' INTO <FS_FIELD>.
    ENDIF.

    " Quantity confirmed

    READ TABLE GT_ZPDND002MT_ALV2_C INDEX 5 ASSIGNING <FS>. "HERE
    IF SY-SUBRC = 0.
      ASSIGN COMPONENT LV_FIELD OF STRUCTURE <FS> TO <FS_FIELD>.

      PERFORM GET_CONF_QTY_ORDER USING    '261'
                                          '262'
                                          LS_TRAINS-AUFNR
                                 CHANGING LV_QUANTITY.

      <FS_FIELD> = LV_QUANTITY.
      REPLACE '.' WITH ',' INTO <FS_FIELD>.


      PERFORM GET_CONF_QTY_UTIL  USING    '201'
                                          '202'
                                           GV_KOSTL_AC
                                 CHANGING LV_QUANTITY.
      <FS>-UTILITES    = LV_QUANTITY.
      REPLACE '.' WITH ',' INTO <FS>-UTILITES.

    ENDIF.
  ENDLOOP.
ENDFORM. " FRM_GET_ORDERS_DATA
*&---------------------------------------------------------------------*
*& Form FRM_INITIALIZE_GLOBAL_VALUES
*&---------------------------------------------------------------------*
FORM FRM_INITIALIZE_GLOBAL_VALUES.

  " Date confirmation AC
  GV_DATE_AC = S_BUDAT-LOW.

  " Date confirmation FG
  GV_DATE_FG = S_BUDAT-LOW.

  " Get KTH index
  PERFORM FRM_GET_INDEX CHANGING GV_FIELD.

  "Get trains to be hidden
  DATA LV_NAME TYPE STRING.
  LV_NAME = 'ZPDND002MT_' && P_WERKS.
  SELECT SINGLE *
   FROM TVARVC
   INTO @DATA(LS_TRAINS)
  WHERE NAME LIKE @LV_NAME
*    AND LOW  LIKE @P_WERKS
    AND TYPE = 'S'.
  IF SY-SUBRC = 0.
    GV_TRAINS_HIDE = LS_TRAINS-LOW.
  ENDIF. " IF SY-SUBRC <> 0.

  """"" HIDE TRAINS
  DATA: LT_TRAINS_HIDE TYPE STANDARD TABLE OF TY_EDIT_COL,
        LS_TRAINS_HIDE TYPE TY_EDIT_COL.
  SPLIT GV_TRAINS_HIDE   AT '-' INTO TABLE LT_TRAINS_HIDE.

  DATA(LV_COLS) = LINES( LT_TRAINS_HIDE ).

ENDFORM. " FRM_INITIALIZE_GLOBAL_VALUES

*&---------------------------------------------------------------------*
*& Form FRM_GET_MATERIALS_DATA
*&---------------------------------------------------------------------*
FORM FRM_GET_MATERIALS_DATA .
  DATA: LS_ALV1 TYPE ZPDND002MT_ALV1,
        LS_ALV2 TYPE ZPDND002MT_ALV2,
        LS_ALV3 TYPE ZPDND002MT_ALV3.

  DATA: LT_MATDOC TYPE TABLE OF MATDOC.
  SELECT * FROM ZPDNAC INTO TABLE @GT_ZPDNAC
    WHERE WERKS = @P_WERKS.
  IF SY-SUBRC <> 0.
    CLEAR GT_ZPDNAC[].
  ENDIF.

  SORT GT_ZPDNAC ASCENDING BY DISPLAY_ORDER.
  LOOP AT GT_ZPDNAC ASSIGNING FIELD-SYMBOL(<FS_ZPDNAC>).
    GT_MATNR-SIGN = 'I'.
    GT_MATNR-OPTION = 'EQ'.
    CASE <FS_ZPDNAC>-ZTYPE.
      WHEN '1'.
        IF <FS_ZPDNAC>-DISPLAY_ORDER = 10.
          GV_MATNR_GN = <FS_ZPDNAC>-MATNR.
        ENDIF.
        LS_ALV1-MATNR = <FS_ZPDNAC>-MATNR.
        LS_ALV1-ZDESC = <FS_ZPDNAC>-ZDESC.
        LS_ALV1-BSTME = <FS_ZPDNAC>-BSTME.
        APPEND LS_ALV1 TO GT_ZPDND002MT_ALV1-TAB1.
      WHEN '2'.
        LS_ALV2-MATNR = <FS_ZPDNAC>-MATNR.
        LS_ALV2-ZDESC = <FS_ZPDNAC>-ZDESC.
        LS_ALV2-BSTME = <FS_ZPDNAC>-BSTME.
        APPEND LS_ALV2 TO GT_ZPDND002MT_ALV1-TAB2.
      WHEN '3' OR '4' OR '5'.
        LS_ALV3-MATNR = <FS_ZPDNAC>-MATNR.
        LS_ALV3-ZDESC = <FS_ZPDNAC>-ZDESC.
        LS_ALV3-BSTME = <FS_ZPDNAC>-BSTME.
        APPEND LS_ALV3 TO GT_ZPDND002MT_ALV1-TAB3.

        IF <FS_ZPDNAC>-ZTYPE = '5'.
          GS_PDNAC_TYPE_5 = <FS_ZPDNAC>.
          GV_COST_CENTER = <FS_ZPDNAC>-ZFGCC.
        ELSEIF <FS_ZPDNAC>-ZTYPE = '3'.
          GS_PDNAC_TYPE_3 = <FS_ZPDNAC>.
        ELSEIF <FS_ZPDNAC>-ZTYPE = '4'.
          GS_PDNAC_TYPE_4 = <FS_ZPDNAC>.
        ENDIF.

    ENDCASE.
    GT_MATNR-LOW = <FS_ZPDNAC>-MATNR.
    APPEND GT_MATNR.
  ENDLOOP.
  " Get materials list
  SORT GT_MATNR ASCENDING.
  DELETE ADJACENT DUPLICATES FROM GT_MATNR.

  SELECT * FROM MATDOC
    INTO TABLE @GT_MATDOC
    WHERE BUDAT IN @S_BUDAT
      AND WERKS  = @P_WERKS
      AND MATNR  IN @GT_MATNR
       AND ( BWART = '101' OR BWART = '102' OR BWART = '261' OR BWART = '262' OR
    BWART = '551' OR BWART = '552' OR BWART = '201' OR BWART = '202'  ).
  IF SY-SUBRC <> 0.
    CLEAR GT_MATDOC[].
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form FRM_CHECK_POSTING_AUTH
*&---------------------------------------------------------------------*
FORM FRM_CHECK_POSTING_AUTH  USING UP_OBJECT TYPE CHAR10
                                   UP_WERKS      TYPE WERKS_D
                                   UP_AUFART     TYPE AUFK-AUART
                           CHANGING CP_AUTHORIZED TYPE ABAP_BOOL.

*  CP_AUTHORIZED = ABAP_TRUE.
  CP_AUTHORIZED = ABAP_FALSE.
  IF UP_OBJECT = 'AC'.

    AUTHORITY-CHECK OBJECT 'ZPDN_AC'
     ID 'WERKS' FIELD UP_WERKS
     ID 'AUFART' FIELD UP_AUFART.
    IF SY-SUBRC = 0.
      CP_AUTHORIZED = ABAP_TRUE.
    ENDIF. " IF sy-subrc <> 0.

  ELSEIF UP_OBJECT = 'FG'.

    AUTHORITY-CHECK OBJECT 'ZPDN_FG'
      ID 'WERKS' FIELD UP_WERKS.
    IF SY-SUBRC = 0.
      CP_AUTHORIZED = ABAP_TRUE.
    ENDIF.

  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form FRM_CHECK_PLANT_AUTH
*&---------------------------------------------------------------------*
FORM FRM_CHECK_PLANT_AUTH  USING UP_WERKS      TYPE WERKS_D
                           CHANGING CP_AUTHORIZED TYPE ABAP_BOOL.

  CP_AUTHORIZED = ABAP_TRUE.
  AUTHORITY-CHECK OBJECT 'ZPDNENHIS'
   ID 'WERKS' FIELD UP_WERKS.
  IF SY-SUBRC <> 0.
    CP_AUTHORIZED = ABAP_FALSE.
    MESSAGE E810(CP) WITH UP_WERKS DISPLAY LIKE 'E'.
  ENDIF. " IF sy-subrc <> 0.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form FRM_GET_MATDOCOIL_DATA
*&---------------------------------------------------------------------*
FORM FRM_GET_MATDOCOIL_DATA.

  SELECT  SINGLE MSEHI_INDEX
  FROM  MATDOCOIL_INDEX
  INTO @DATA(LV_MSEHI_INDEX)
  WHERE MSEHI = 'KTH'. "
  IF SY-SUBRC <> 0.
    MESSAGE E020(ZPDN001) WITH 'KTH'.
  ELSE.
    GV_FIELD =  'ADQNTP' &&  LV_MSEHI_INDEX.
  ENDIF.

ENDFORM. " FRM_GET_MATDOCOIL_DATA

**&---------------------------------------------------------------------*
**& Form FRM_GET_MAT_DOC_QTY
**&---------------------------------------------------------------------*
FORM FRM_GET_MAT_DOC_QTY USING UP_MATNR TYPE MARA-MATNR
                               UP_BWART_P TYPE CHAR20
                               UP_BWART_N TYPE CHAR20
                        CHANGING CP_ERFMG TYPE MENGE_D.
  DATA: LT_BWART_P TYPE STANDARD TABLE OF CHAR03,
        LT_BWART_N TYPE STANDARD TABLE OF CHAR03,
        LS_BWART_P TYPE CHAR03,
        LS_BWART_N TYPE CHAR03,
        LV_ERFMG   TYPE F.

  SPLIT  UP_BWART_P  AT '-' INTO TABLE LT_BWART_P.
  LOOP AT LT_BWART_P INTO LS_BWART_P.
    CLEAR: LV_ERFMG.
    SELECT SUM( CASE SHKZG
                   WHEN 'S'
                     THEN ERFMG
                   WHEN 'H'
                     THEN
                     ERFMG * ( -1 )
                 END ) AS ERFMG, ERFME
     FROM @GT_MATDOC AS A
*     FROM MATDOC AS A
      WHERE BUDAT IN @S_BUDAT
        AND WERKS  = @P_WERKS
        AND MATNR  = @UP_MATNR
        AND BWART = @LS_BWART_P
      GROUP BY ERFME
      INTO TABLE @DATA(LT_RESULT_P).
    IF SY-SUBRC <> 0.
      CLEAR: LV_ERFMG.
    ELSE.
      LOOP AT LT_RESULT_P ASSIGNING FIELD-SYMBOL(<FS_RESULT_P>) .
        CASE <FS_RESULT_P>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + (  ABS( <FS_RESULT_P>-ERFMG ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG +   ABS( <FS_RESULT_P>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG + ABS( <FS_RESULT_P>-ERFMG ) .
        ENDCASE.
      ENDLOOP.
      CP_ERFMG = CP_ERFMG + LV_ERFMG.
    ENDIF.
  ENDLOOP.

  SPLIT  UP_BWART_N  AT '-' INTO TABLE LT_BWART_N.
  LOOP AT LT_BWART_N INTO LS_BWART_N.
    CLEAR: LV_ERFMG.
    SELECT SUM( CASE SHKZG
                   WHEN 'S'
                     THEN ERFMG
                   WHEN 'H'
                     THEN
                     ERFMG * ( -1 )
                 END ) AS ERFMG, ERFME
     FROM @GT_MATDOC AS A
*                 FROM MATDOC AS A
      WHERE BUDAT IN @S_BUDAT
        AND WERKS  = @P_WERKS
        AND MATNR  = @UP_MATNR
        AND BWART = @LS_BWART_N
     GROUP BY ERFME
     INTO TABLE @DATA(LT_RESULT_N).
    IF SY-SUBRC <> 0.
      CLEAR: LV_ERFMG.
    ELSE.
      LOOP AT LT_RESULT_N ASSIGNING FIELD-SYMBOL(<FS_RESULT_N>).
        CASE <FS_RESULT_N>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + ( ABS( <FS_RESULT_N>-ERFMG ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG + ABS( <FS_RESULT_N>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG + ABS( <FS_RESULT_N>-ERFMG ).
        ENDCASE.
      ENDLOOP.
      CP_ERFMG = CP_ERFMG - LV_ERFMG.
    ENDIF.
  ENDLOOP.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form FRM_GET_ENERGY
*&---------------------------------------------------------------------*
FORM FRM_GET_ENERGY  USING    PV_BWART_P
                              PV_BWART_N
                              PV_FIELD
                              PV_MATNR
                              PV_AUFNR
                     CHANGING PV_ENERGY TYPE ZPCS23_6.
  DATA LV_ADQNTP TYPE OIB_01_ADQNTP.
  DATA GT_ADQNTP TYPE STANDARD TABLE OF OIB_01_ADQNTP.
  FIELD-SYMBOLS: <FS_ADQNTP>.
  CLEAR PV_ENERGY.

  IF PV_MATNR IS INITIAL.
    RANGES: LR_MATNR FOR  MARA-MATNR.
    LR_MATNR-OPTION = 'EQ'.
    LR_MATNR-SIGN = 'I'.

    SELECT DISTINCT MATNR FROM @GT_OUT_ORDERS AS A INTO @DATA(LV_MATNR).
      LR_MATNR-LOW = LV_MATNR.
      APPEND LR_MATNR.
    ENDSELECT.
    SELECT (PV_FIELD)
      FROM @GT_MATDOC AS C
*      FROM MATDOC AS C
      INNER JOIN MATDOCOIL AS A
       ON A~MBLNR = C~MBLNR
      AND A~MJAHR = C~MJAHR
      AND A~ZEILE = C~ZEILE
    WHERE C~WERKS = @P_WERKS
      AND C~AUFNR = @PV_AUFNR
      AND C~BUDAT IN @S_BUDAT
      AND C~BWART = @PV_BWART_P
      AND C~MATNR  IN @LR_MATNR
      INTO TABLE @GT_ADQNTP.
    IF SY-SUBRC = 0.
      LOOP AT GT_ADQNTP ASSIGNING <FS_ADQNTP>.
        LV_ADQNTP = LV_ADQNTP + ABS( <FS_ADQNTP> ).
      ENDLOOP.
      PV_ENERGY = PV_ENERGY + LV_ADQNTP.
    ENDIF.

    CLEAR: LV_ADQNTP, GT_ADQNTP[].
    SELECT (PV_FIELD)
        FROM @GT_MATDOC AS C
*      FROM MATDOC AS C
        INNER JOIN MATDOCOIL AS A
        ON A~MBLNR = C~MBLNR
        AND A~MJAHR = C~MJAHR
        AND A~ZEILE = C~ZEILE
      WHERE C~WERKS = @P_WERKS
        AND C~AUFNR = @PV_AUFNR
        AND C~BUDAT IN @S_BUDAT
        AND C~BWART = @PV_BWART_N
        AND C~MATNR   IN @LR_MATNR
    INTO TABLE @GT_ADQNTP.
    IF SY-SUBRC = 0.
      LOOP AT GT_ADQNTP ASSIGNING <FS_ADQNTP>.
        LV_ADQNTP = LV_ADQNTP + ABS( <FS_ADQNTP> ).
      ENDLOOP.
      PV_ENERGY = PV_ENERGY - LV_ADQNTP.
    ENDIF.

  ELSE. "   IF PV_MATNR IS INITIAL.

    SELECT (PV_FIELD)
      FROM @GT_MATDOC AS A
*      FROM MATDOC AS A
      INNER JOIN MATDOCOIL AS B
        ON A~MBLNR = B~MBLNR
        AND A~MJAHR = B~MJAHR
        AND A~ZEILE = B~ZEILE
        WHERE A~WERKS = @P_WERKS
          AND  A~AUFNR = @PV_AUFNR
          AND MATNR = @PV_MATNR
          AND A~BUDAT IN @S_BUDAT
          AND A~BWART = @PV_BWART_P
      INTO TABLE @GT_ADQNTP.
    IF SY-SUBRC = 0.
      LOOP AT GT_ADQNTP ASSIGNING <FS_ADQNTP>.           "#EC CI_NESTED
        LV_ADQNTP = LV_ADQNTP + ABS( <FS_ADQNTP> ).
      ENDLOOP.
      PV_ENERGY = LV_ADQNTP.
    ENDIF.

    CLEAR: LV_ADQNTP, GT_ADQNTP[].
    SELECT  (PV_FIELD)
       FROM @GT_MATDOC AS A
*      FROM MATDOC AS A
       INNER JOIN MATDOCOIL AS B
       ON A~MBLNR = B~MBLNR
        AND A~MJAHR = B~MJAHR
        AND A~ZEILE = B~ZEILE
      WHERE A~WERKS = @P_WERKS
        AND  A~AUFNR = @PV_AUFNR
        AND MATNR = @PV_MATNR
        AND A~BUDAT IN @S_BUDAT
        AND  A~BWART = @PV_BWART_N
      INTO TABLE @GT_ADQNTP.
    IF SY-SUBRC = 0.
      LOOP AT GT_ADQNTP ASSIGNING <FS_ADQNTP>.           "#EC CI_NESTED
        LV_ADQNTP = LV_ADQNTP + ABS( <FS_ADQNTP> ).
      ENDLOOP.
      PV_ENERGY = PV_ENERGY - LV_ADQNTP.
    ENDIF.
  ENDIF. " IF PV_MATNR IS INITIAL.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form FRM_GET_DATA
*&---------------------------------------------------------------------*
FORM FRM_GET_DATA.

  PERFORM FRM_INITIALIZE_GLOBAL_VALUES.

  PERFORM FRM_GET_MATERIALS_DATA.

  PERFORM FRM_GET_QUANTITIES_DATA.

  PERFORM FRM_GET_ORDERS_DATA CHANGING GV_CHECK_OK.

ENDFORM." FRM_GET_DATA


*&---------------------------------------------------------------------*
*& Form GET_ORDERS_LIST
*&---------------------------------------------------------------------*
FORM GET_ORDERS_LIST TABLES PT_ORDERS TYPE TT_ORDERS
                            PT_OUT_ORDERS TYPE TT_OUT_ORDERS.

  DATA: LR_AUFNR TYPE RANGE OF AUFNR.

  " Get All Released process Orders
  SELECT  ZPDNTRAIN~TRAIN, AUFK~AUFNR, CRHD~ARBPL, ( CASE AUART WHEN 'ZI01' THEN 'WORKING' ELSE 'SHUTDOWN' END ) AS STAT
    FROM AFKO
   INNER JOIN AUFK
      ON AFKO~AUFNR = AUFK~AUFNR
   INNER JOIN AFVC
      ON AFKO~AUFPL = AFVC~AUFPL
   INNER JOIN CRHD
      ON CRHD~OBJID = AFVC~ARBID
   INNER JOIN AFPO
      ON AFKO~AUFNR = AFPO~AUFNR
   INNER JOIN JEST
      ON AUFK~OBJNR = JEST~OBJNR
   INNER JOIN ZPDNTRAIN
      ON ZPDNTRAIN~WERKS = AUFK~WERKS
     AND ZPDNTRAIN~ARBPL = CRHD~ARBPL
   WHERE AUFK~WERKS  = @P_WERKS
     AND AFKO~GSTRP <= @S_BUDAT-HIGH
     AND AFKO~GLTRP >= @S_BUDAT-LOW
     AND AUFK~AUART  IN @GT_AUART
     AND JEST~INACT  = ''
     AND JEST~STAT   = 'I0002' "(Released)
    INTO TABLE @PT_ORDERS.
  IF SY-SUBRC <> 0.
    CLEAR PT_ORDERS[].
  ENDIF. " IF SY-SUBRC <> 0.

  "Needed to declare LS_AUFNR
  LOOP AT LR_AUFNR INTO DATA(LS_AUFNR).
  ENDLOOP.

  LS_AUFNR-OPTION = 'EQ'.
  LS_AUFNR-SIGN   = 'I'.
  LOOP AT PT_ORDERS ASSIGNING FIELD-SYMBOL(<FS_ORDER>).
    LS_AUFNR-LOW = <FS_ORDER>-AUFNR.
    APPEND LS_AUFNR TO LR_AUFNR.
  ENDLOOP.

  " Get all SHUTDWON Process orders
  SELECT AUFK~AUFNR
    FROM AUFK
   INNER JOIN JEST
      ON AUFK~OBJNR = JEST~OBJNR
   WHERE AUFK~AUFNR IN @LR_AUFNR
     AND JEST~INACT  = ''
     AND JEST~STAT   = 'E0001' " (SHUTDOWN)
     INTO TABLE @DATA(LT_SHUDOWN_ORDERS).
  IF SY-SUBRC = 0.

    " DELETE orders having SHUTDOWN STATUS : E0001-Active
    LOOP AT LT_SHUDOWN_ORDERS ASSIGNING FIELD-SYMBOL(<FS_SHT_ORDER>).
      DELETE PT_ORDERS WHERE AUFNR = <FS_SHT_ORDER>-AUFNR.
    ENDLOOP. " LOOP AT LT_SHUDOWN_ORDERS ASSIGNING FIELD-SYMBOL(<FS_SHT_ORDER>).
  ENDIF. " IF SY-SUBRC = 0.


  SELECT  * FROM  MATDOC AS A
    FOR ALL ENTRIES IN @GT_ORDERS
   WHERE AUFNR = @GT_ORDERS-AUFNR
    AND WERKS = @P_WERKS
  INTO CORRESPONDING FIELDS OF TABLE @PT_OUT_ORDERS.
  IF SY-SUBRC <> 0.
    CLEAR PT_OUT_ORDERS[].
  ENDIF.

  SORT PT_ORDERS BY TRAIN AUFNR.
  DELETE ADJACENT DUPLICATES FROM PT_ORDERS COMPARING TRAIN AUFNR.

ENDFORM. " GET_ORDERS_LIST

*&---------------------------------------------------------------------*
*& Form FRM_CHECK_ORDERS
*&---------------------------------------------------------------------*
FORM FRM_CHECK_ORDERS  TABLES   PT_ORDERS   TYPE TT_ORDERS
                       CHANGING CP_CHECK_OK TYPE ABAP_BOOL
                                CP_ORDER1   TYPE AUFNR
                                CP_ORDER2   TYPE AUFNR.

  CLEAR: CP_ORDER1, CP_ORDER2.
  CP_CHECK_OK = ABAP_TRUE.

  LOOP AT PT_ORDERS INTO DATA(LS_ODERS)
  GROUP BY ( TRAIN  = LS_ODERS-TRAIN
             SIZE  = GROUP SIZE
             INDEX = GROUP INDEX )
  ASCENDING
  ASSIGNING FIELD-SYMBOL(<FS_ORDER_GROUP>).

    IF <FS_ORDER_GROUP>-SIZE > 1.

      LOOP AT GROUP <FS_ORDER_GROUP> ASSIGNING FIELD-SYMBOL(<FS_ORDER>).
        CP_CHECK_OK = ABAP_FALSE.
        IF CP_ORDER1 IS INITIAL.
          CP_ORDER1 = <FS_ORDER>-AUFNR.
        ELSEIF CP_ORDER2 IS INITIAL.
          CP_ORDER2 = <FS_ORDER>-AUFNR.
        ENDIF.

        IF CP_ORDER1 IS NOT INITIAL AND
           CP_ORDER2 IS NOT INITIAL.
          EXIT.
        ENDIF. " IF CP_ORDER1 IS NOT INITIAL AND
      ENDLOOP.
      EXIT.
    ENDIF. " IF <FS_ORDER_GROUP>-SIZE > 1.
  ENDLOOP.
ENDFORM. " FRM_CHECK_ORDERS
*&---------------------------------------------------------------------*
*& Form FRM_GET_INDEX
*&---------------------------------------------------------------------*
FORM FRM_GET_INDEX CHANGING CP_FIELD TYPE STRING.

  SELECT  SINGLE MSEHI_INDEX
    FROM  MATDOCOIL_INDEX
    INTO @DATA(LV_MSEHI_INDEX)
    WHERE MSEHI = 'KTH'. "
  IF SY-SUBRC <> 0.
    MESSAGE E020(ZPDN001) WITH 'KTH'.
  ELSE.
    CP_FIELD =  'ADQNTP' &&  LV_MSEHI_INDEX.
  ENDIF.
ENDFORM. " FRM_GET_INDEX
*&---------------------------------------------------------------------*
*& Form FRM_GET_QUANTITIES_DATA
*&---------------------------------------------------------------------*
FORM FRM_GET_QUANTITIES_DATA.
  " Get quantities

  LOOP AT GT_ZPDND002MT_ALV1-TAB1 ASSIGNING FIELD-SYMBOL(<FS_TAB1>).

    PERFORM FRM_GET_QTY_ENERGY_CONS USING  <FS_TAB1>-MATNR
                                       '261-201'
                                       '262-202'
                                CHANGING  <FS_TAB1>-QTY_CONS
                                          <FS_TAB1>-ENERGY_KTH.
    IF SY-TABIX = 1.
      PERFORM FRM_CONVERT_TO_UOM_ZPDNAC USING <FS_TAB1>-BSTME
                                              'CMK'
                                              <FS_TAB1>-QTY_CONS.
    ENDIF.
    PERFORM FRM_GET_QTY_ENERGY_MOVED USING  <FS_TAB1>-MATNR
                                                '101'
                                                '102'
                                                'WE'
                                CHANGING  <FS_TAB1>-QTY_IN
                                          <FS_TAB1>-ENERGY_IN.

  ENDLOOP.

  LOOP AT GT_ZPDND002MT_ALV1-TAB2 ASSIGNING FIELD-SYMBOL(<FS_TAB2>).

    DATA: LV_ERFMG  TYPE MENGE_D,
          LV_ENERGY TYPE ZPCS23_6.
    PERFORM FRM_GET_QTY_ENERGY_MOVED USING  <FS_TAB2>-MATNR
                                                '101'
                                                '102'
                                                'WF'
                                CHANGING  LV_ERFMG "<FS_TAB2>-QTY_PROD
                                          LV_ENERGY. " <FS_TAB2>-ENERGY_OUT.
    <FS_TAB2>-QTY_PROD = LV_ERFMG.
    <FS_TAB2>-ENERGY_OUT = LV_ENERGY.
    PERFORM FRM_GET_QTY_ENERGY_MOVED USING  <FS_TAB2>-MATNR
                                            '101'
                                            '102'
                                            'WR'
                            CHANGING  LV_ERFMG "<FS_TAB2>-QTY_PROD
                                       LV_ENERGY. "<FS_TAB2>-ENERGY_OUT.
    <FS_TAB2>-QTY_PROD = <FS_TAB2>-QTY_PROD + LV_ERFMG.
    <FS_TAB2>-ENERGY_OUT =  <FS_TAB2>-ENERGY_OUT + LV_ENERGY.


    PERFORM FRM_GET_QTY_ENERGY_BOILOFF USING  <FS_TAB2>-MATNR
                                                '551'
                                                '552'
                                                '3'
                                CHANGING  <FS_TAB2>-QTY_BOILOFF
                                          <FS_TAB2>-ENERGY_BOILOFF.
  ENDLOOP.


  LOOP AT GT_ZPDND002MT_ALV1-TAB3 ASSIGNING FIELD-SYMBOL(<FS_TAB3>).
    "   Quantité Confirmé
    PERFORM FRM_GET_MAT_DOC_QTY USING  <FS_TAB3>-MATNR
                                       '261-201'
                                       '262-202'
                                CHANGING <FS_TAB3>-QTY_CONF.

  ENDLOOP.
ENDFORM. " FRM_GET_QUANTITIES_DATA
*&---------------------------------------------------------------------*
*& Form GET_CONF_QTY_ORDER
*&---------------------------------------------------------------------*
FORM GET_CONF_QTY_ORDER USING   UP_BWART_P
                                UP_BWART_N
                                UP_AUFNR
                        CHANGING CP_ERFMG TYPE MENGE_D.
  CLEAR CP_ERFMG.
  DATA LV_ERFMG TYPE F.
  SELECT SINGLE MATNR FROM @GT_ZPDNAC AS LT_PDNAC WHERE ZTYPE = '3' INTO @DATA(LV_MATNR).
  IF SY-SUBRC = 0.
    CLEAR: LV_ERFMG.
    SELECT SUM( CASE SHKZG
                      WHEN 'S'
                        THEN ERFMG
                      WHEN 'H'
                        THEN
                        ERFMG * ( -1 )
                    END ) AS ERFMG , ERFME
        FROM @GT_OUT_ORDERS AS A
            WHERE BWART = @UP_BWART_P
             AND  AUFNR = @UP_AUFNR
             AND  MATNR = @LV_MATNR
             AND  BUDAT IN @S_BUDAT
        GROUP BY ERFME
        INTO TABLE @DATA(LT_RESULT_P).
    IF SY-SUBRC <> 0.
      CLEAR: LV_ERFMG.
    ELSE.
      LOOP AT LT_RESULT_P ASSIGNING FIELD-SYMBOL(<FS_RESULT_P>) .
        CASE <FS_RESULT_P>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + ( ABS( <FS_RESULT_P>-ERFMG ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG +  ABS( <FS_RESULT_P>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG +  ABS( <FS_RESULT_P>-ERFMG ).
        ENDCASE.
      ENDLOOP.
      CP_ERFMG = CP_ERFMG + LV_ERFMG.
    ENDIF.

    CLEAR LV_ERFMG.
    SELECT SUM( CASE SHKZG
                    WHEN 'S'
                      THEN ERFMG
                    WHEN 'H'
                      THEN
                      ERFMG * ( -1 )
                  END ) AS ERFMG , ERFME
      FROM @GT_OUT_ORDERS AS A
          WHERE BWART = @UP_BWART_N
             AND  AUFNR = @UP_AUFNR
             AND  MATNR = @LV_MATNR
             AND  BUDAT IN @S_BUDAT
        GROUP BY ERFME
        INTO TABLE @DATA(LT_RESULT_N).
    IF SY-SUBRC <> 0.
      CLEAR: LV_ERFMG.
    ELSE.
      LOOP AT LT_RESULT_N ASSIGNING FIELD-SYMBOL(<FS_RESULT_N>).
        CASE <FS_RESULT_N>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + ( ABS( <FS_RESULT_N>-ERFMG  ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG +  ABS( <FS_RESULT_N>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG +  ABS( <FS_RESULT_N>-ERFMG ).
        ENDCASE.
      ENDLOOP.
      CP_ERFMG = CP_ERFMG - LV_ERFMG.
    ENDIF. " IF sy-subrc <> 0.
  ENDIF. "   IF SY-SUBRC = 0.
ENDFORM. " GET_CONF_QTY_ORDER

*&---------------------------------------------------------------------*
*& Form GET_CONF_QTY_ORDER
*&---------------------------------------------------------------------*
FORM GET_CONF_QTY_UTIL USING   UP_BWART_P
                                UP_BWART_N
                                UP_KOSTL
                        CHANGING CP_ERFMG TYPE MENGE_D.
  CLEAR CP_ERFMG.
  DATA LV_ERFMG TYPE F.
  SELECT SINGLE MATNR FROM @GT_ZPDNAC AS LT_PDNAC WHERE ZTYPE = '3' INTO @DATA(LV_MATNR).
  IF SY-SUBRC = 0.
    SELECT SUM( CASE SHKZG
                      WHEN 'S'
                        THEN ERFMG
                      WHEN 'H'
                        THEN
                        ERFMG * ( -1 )
                    END ) AS ERFMG, ERFME
        FROM @GT_MATDOC AS A
            WHERE BWART = @UP_BWART_P
             AND  KOSTL = @UP_KOSTL
             AND  MATNR = @LV_MATNR
             AND BUDAT IN @S_BUDAT
        GROUP BY ERFME
        INTO TABLE @DATA(LT_RESULT_P).
    IF SY-SUBRC <> 0.
      CLEAR: LT_RESULT_P[].
    ELSE.
      CLEAR: LV_ERFMG.
      LOOP AT LT_RESULT_P ASSIGNING FIELD-SYMBOL(<FS_RESULT_P>) .
        CASE <FS_RESULT_P>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + ( ABS( <FS_RESULT_P>-ERFMG ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG + ABS( <FS_RESULT_P>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG + ABS( <FS_RESULT_P>-ERFMG ).
        ENDCASE.
      ENDLOOP.
      CP_ERFMG = CP_ERFMG + LV_ERFMG..
    ENDIF. " IF sy-subrc <> 0

    CLEAR LV_ERFMG.
    SELECT SUM( CASE SHKZG
                    WHEN 'S'
                      THEN ERFMG
                    WHEN 'H'
                      THEN
                      ERFMG * ( -1 )
                  END ) AS ERFMG, ERFME
      FROM @GT_MATDOC AS A
      WHERE BWART = @UP_BWART_N
       AND  KOSTL = @UP_KOSTL
       AND  MATNR = @LV_MATNR
       AND  BUDAT IN @S_BUDAT
      GROUP BY ERFME
      INTO TABLE @DATA(LT_RESULT_N).
    IF SY-SUBRC <> 0.
      CLEAR: LT_RESULT_N[].
      CLEAR LV_ERFMG.
    ELSE.
      LOOP AT LT_RESULT_N ASSIGNING FIELD-SYMBOL(<FS_RESULT_N>) .
        CASE <FS_RESULT_N>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + ( ABS( <FS_RESULT_N>-ERFMG ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG + ABS( <FS_RESULT_N>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG + ABS( <FS_RESULT_N>-ERFMG ).
        ENDCASE.
        CP_ERFMG = CP_ERFMG - LV_ERFMG.
      ENDLOOP.
    ENDIF. " IF sy-subrc <> 0.
  ENDIF. "   IF SY-SUBRC = 0.
ENDFORM. " GET_CONF_QTY_ORDER

**&---------------------------------------------------------------------*
**& Form FRM_GET_QTY_ENERGY_CONS
**&---------------------------------------------------------------------*
FORM FRM_GET_QTY_ENERGY_CONS USING UP_MATNR TYPE MARA-MATNR
                               UP_BWART_P TYPE CHAR20
                               UP_BWART_N TYPE CHAR20
                        CHANGING CP_ERFMG TYPE MENGE_D
                                 CP_ENERGY TYPE ZPCS23_6.
  DATA: LT_BWART_P TYPE STANDARD TABLE OF CHAR03,
        LT_BWART_N TYPE STANDARD TABLE OF CHAR03,
        LS_BWART_P TYPE CHAR03,
        LS_BWART_N TYPE CHAR03,
        LV_ERFMG   TYPE F,
        LV_ADQNTP  TYPE OIB_01_ADQNTP,
        GT_ADQNTP  TYPE STANDARD TABLE OF OIB_01_ADQNTP.

  FIELD-SYMBOLS: <FS_ADQNTP>.
  CLEAR: CP_ERFMG, CP_ENERGY.

  SPLIT  UP_BWART_P  AT '-' INTO TABLE LT_BWART_P.
  CLEAR LV_ERFMG.
  LOOP AT LT_BWART_P INTO LS_BWART_P.
    SELECT SUM(  CASE SHKZG
                   WHEN 'S'
                     THEN ERFMG
                WHEN 'H'
                     THEN
                     ERFMG * ( -1 )
                 END
                ) AS ERFMG, ERFME
     FROM @GT_MATDOC AS A
*                 FROM MATDOC AS A
      WHERE WERKS  = @P_WERKS
        AND BUDAT IN @S_BUDAT
        AND MATNR  = @UP_MATNR
        AND BWART = @LS_BWART_P
      GROUP BY ERFME
      INTO TABLE @DATA(LT_RESULT_P).
    IF SY-SUBRC <> 0.
      CLEAR: LT_RESULT_P[].
    ELSE.
      CLEAR: LV_ERFMG.
      LOOP AT LT_RESULT_P ASSIGNING FIELD-SYMBOL(<FS_RESULT_P>) .
        CASE <FS_RESULT_P>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + (  ABS( <FS_RESULT_P>-ERFMG ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG +   ABS( <FS_RESULT_P>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG +   ABS( <FS_RESULT_P>-ERFMG ).
        ENDCASE.
      ENDLOOP.
      CP_ERFMG = CP_ERFMG + LV_ERFMG.
    ENDIF. " IF sy-subrc <> 0.

    " Energy
    CLEAR: LV_ADQNTP, GT_ADQNTP[].
    SELECT (GV_FIELD)
      FROM @GT_MATDOC AS A
*      FROM MATDOC AS A
       INNER JOIN MATDOCOIL AS B
       ON A~MBLNR = B~MBLNR
        AND A~MJAHR = B~MJAHR
        AND A~ZEILE = B~ZEILE

      WHERE A~WERKS = @P_WERKS
        AND MATNR = @UP_MATNR
        AND A~BUDAT IN @S_BUDAT
        AND A~BWART = @LS_BWART_P
      INTO TABLE @GT_ADQNTP.
    IF SY-SUBRC = 0.
      LOOP AT GT_ADQNTP ASSIGNING <FS_ADQNTP>.           "#EC CI_NESTED
        LV_ADQNTP = LV_ADQNTP + ABS( <FS_ADQNTP> ).
      ENDLOOP.
      CP_ENERGY = CP_ENERGY + LV_ADQNTP.
    ENDIF.

  ENDLOOP.

  SPLIT  UP_BWART_N  AT '-' INTO TABLE LT_BWART_N.

  LOOP AT LT_BWART_N INTO LS_BWART_N.
    CLEAR: LV_ERFMG.
    SELECT SUM(  CASE SHKZG
                   WHEN 'S'
                     THEN ERFMG
                WHEN 'H'
                     THEN
                     ERFMG * ( -1 )
                 END
                ) AS ERFMG, ERFME
      FROM @GT_MATDOC AS A
*                 FROM MATDOC AS A
       WHERE WERKS  = @P_WERKS
        AND BUDAT IN @S_BUDAT
        AND MATNR  = @UP_MATNR
        AND BWART = @LS_BWART_N
      GROUP BY ERFME
      INTO TABLE @DATA(LT_RESULT_N).
    IF SY-SUBRC <> 0.
      CLEAR: LT_RESULT_N[].
    ELSE.
      CLEAR: LV_ERFMG.
      LOOP AT LT_RESULT_N ASSIGNING FIELD-SYMBOL(<FS_RESULT_N>).
        CASE <FS_RESULT_N>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + (  ABS( <FS_RESULT_N>-ERFMG ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG +   ABS( <FS_RESULT_N>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG +   ABS( <FS_RESULT_N>-ERFMG ).
        ENDCASE.
      ENDLOOP.
      CP_ERFMG = CP_ERFMG - LV_ERFMG.
    ENDIF. " IF sy-subrc <> 0.


    " Energy
    CLEAR: LV_ADQNTP, GT_ADQNTP[].
    SELECT  (GV_FIELD)
      FROM @GT_MATDOC AS A
*      FROM MATDOC AS A
       INNER JOIN MATDOCOIL AS B
       ON A~MBLNR = B~MBLNR
        AND A~MJAHR = B~MJAHR
        AND A~ZEILE = B~ZEILE
      WHERE A~WERKS = @P_WERKS
        AND MATNR = @UP_MATNR
        AND A~BUDAT IN @S_BUDAT
        AND  A~BWART = @LS_BWART_N
      INTO TABLE @GT_ADQNTP.

    LOOP AT GT_ADQNTP ASSIGNING <FS_ADQNTP>.
      LV_ADQNTP = LV_ADQNTP + ABS( <FS_ADQNTP> ).
    ENDLOOP.

    CP_ENERGY = CP_ENERGY - LV_ADQNTP.

  ENDLOOP.
ENDFORM. " FRM_GET_QTY_ENERGY_CONS

**&---------------------------------------------------------------------*
**& Form FRM_GET_QTY_ENERGY_MOVE
**&---------------------------------------------------------------------*
FORM FRM_GET_QTY_ENERGY_MOVED USING   UP_MATNR    TYPE MARA-MATNR
                                      UP_BWART_P  TYPE CHAR20
                                      UP_BWART_N  TYPE CHAR20
                                      UP_VGART    TYPE VGART
                              CHANGING CP_ERFMG   TYPE MENGE_D
                                       CP_ENERGY TYPE ZPCS23_6.
  DATA: LT_BWART_P TYPE STANDARD TABLE OF CHAR03,
        LT_BWART_N TYPE STANDARD TABLE OF CHAR03,
        LS_BWART_P TYPE CHAR03,
        LS_BWART_N TYPE CHAR03,
        LV_ERFMG   TYPE F,
        LV_ENERGY  TYPE ZPCS23_6,
        LV_ADQNTP  TYPE OIB_01_ADQNTP,
        GT_ADQNTP  TYPE STANDARD TABLE OF OIB_01_ADQNTP.
  FIELD-SYMBOLS: <FS_ADQNTP>.

  CLEAR: CP_ERFMG, CP_ENERGY.

  SPLIT  UP_BWART_P  AT '-' INTO TABLE LT_BWART_P.
  CLEAR LV_ERFMG.
  LOOP AT LT_BWART_P INTO LS_BWART_P.
    SELECT SUM(  CASE SHKZG
                   WHEN 'S'
                     THEN ERFMG
                WHEN 'H'
                     THEN
                     ERFMG * ( -1 )
                 END
                ) AS ERFMG, ERFME
     FROM @GT_MATDOC AS A
*                 FROM MATDOC AS A
      WHERE BUDAT IN @S_BUDAT
        AND WERKS  = @P_WERKS
        AND MATNR  = @UP_MATNR
        AND BWART = @LS_BWART_P
        AND VGART = @UP_VGART
     GROUP BY ERFME
     INTO TABLE @DATA(LT_RESULT_P).
    IF SY-SUBRC <> 0.
      CLEAR: LV_ERFMG, LT_RESULT_P[].
    ELSE.
      CLEAR: LV_ERFMG.
      LOOP AT LT_RESULT_P ASSIGNING FIELD-SYMBOL(<FS_RESULT_P>).
        CASE <FS_RESULT_P>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + (  ABS( <FS_RESULT_P>-ERFMG ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG +   ABS( <FS_RESULT_P>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG +   ABS( <FS_RESULT_P>-ERFMG ).
        ENDCASE.
      ENDLOOP.
      CP_ERFMG = CP_ERFMG + LV_ERFMG.
    ENDIF.


    " Energy
    CLEAR: LV_ADQNTP, GT_ADQNTP[].
    SELECT (GV_FIELD)
      FROM @GT_MATDOC AS A
*      FROM MATDOC AS A
        INNER JOIN MATDOCOIL AS B
        ON A~MBLNR = B~MBLNR
         AND A~MJAHR = B~MJAHR
         AND A~ZEILE = B~ZEILE
       WHERE A~WERKS = @P_WERKS
         AND MATNR = @UP_MATNR
         AND A~BUDAT IN @S_BUDAT
         AND A~BWART = @LS_BWART_P
         AND VGART = @UP_VGART
      INTO TABLE @GT_ADQNTP.
    IF SY-SUBRC = 0.
      LOOP AT GT_ADQNTP ASSIGNING <FS_ADQNTP>.           "#EC CI_NESTED
        LV_ADQNTP = LV_ADQNTP + ABS( <FS_ADQNTP> ).
      ENDLOOP.
      CP_ENERGY = CP_ENERGY  + LV_ADQNTP.
    ENDIF.
  ENDLOOP.

  SPLIT  UP_BWART_N  AT '-' INTO TABLE LT_BWART_N.
  CLEAR LV_ERFMG.
  LOOP AT LT_BWART_N INTO LS_BWART_N.
    SELECT SUM(  CASE SHKZG
                 WHEN 'S'
                   THEN ERFMG
              WHEN 'H'
                   THEN
                   ERFMG * ( -1 )
               END
              ) AS ERFMG, ERFME
     FROM @GT_MATDOC AS A
*               FROM MATDOC AS A
      WHERE BUDAT IN @S_BUDAT
        AND WERKS  = @P_WERKS
        AND MATNR  = @UP_MATNR
        AND BWART = @LS_BWART_N
        AND VGART = @UP_VGART
     GROUP BY ERFME
     INTO  TABLE @DATA(LT_RESULT_N).
    IF SY-SUBRC <> 0.
      CLEAR: LV_ERFMG, LT_RESULT_N[].
    ELSE.
      CLEAR: LV_ERFMG.
      LOOP AT LT_RESULT_N ASSIGNING FIELD-SYMBOL(<FS_RESULT_N>)." WHERE ERFME = 'CM3'.
        CASE <FS_RESULT_N>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + (  ABS( <FS_RESULT_N>-ERFMG ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG +   ABS( <FS_RESULT_N>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG +   ABS( <FS_RESULT_N>-ERFMG ).
        ENDCASE.
      ENDLOOP.
      CP_ERFMG = CP_ERFMG - LV_ERFMG.
    ENDIF. " IF sy-subrc <> 0.

    " Energy
    CLEAR: LV_ADQNTP, GT_ADQNTP[].
    SELECT (GV_FIELD)
        FROM @GT_MATDOC AS A
*      FROM MATDOC AS A
        INNER JOIN MATDOCOIL AS B
        ON A~MBLNR = B~MBLNR
         AND A~MJAHR = B~MJAHR
         AND A~ZEILE = B~ZEILE
       WHERE A~WERKS = @P_WERKS
         AND MATNR = @UP_MATNR
         AND A~BUDAT IN @S_BUDAT
         AND A~BWART = @LS_BWART_N
         AND VGART = @UP_VGART
      INTO TABLE @GT_ADQNTP.
    IF SY-SUBRC = 0.
      LOOP AT GT_ADQNTP ASSIGNING <FS_ADQNTP>.           "#EC CI_NESTED
        LV_ADQNTP = LV_ADQNTP + ABS( <FS_ADQNTP> ).
      ENDLOOP.
      CP_ENERGY = CP_ENERGY  - LV_ADQNTP.
    ENDIF.
  ENDLOOP.
ENDFORM. " FRM_GET_QTY_ENERGY_MOVE


**&---------------------------------------------------------------------*
**& Form FRM_GET_QTY_ENERGY_BOILOFF
**&---------------------------------------------------------------------*
FORM FRM_GET_QTY_ENERGY_BOILOFF USING UP_MATNR    TYPE MARA-MATNR
                                      UP_BWART_P  TYPE CHAR20
                                      UP_BWART_N  TYPE CHAR20
                                      UP_GRUND    TYPE MATDOC-GRUND
                              CHANGING CP_ERFMG   TYPE MENGE_D
                                       CP_ENERGY TYPE ZPCS23_6.

  DATA: LT_BWART_P TYPE STANDARD TABLE OF CHAR03,
        LT_BWART_N TYPE STANDARD TABLE OF CHAR03,
        LS_BWART_P TYPE CHAR03,
        LS_BWART_N TYPE CHAR03,
        LV_ERFMG   TYPE F,
        LV_ENERGY  TYPE ZPCS23_6,
        LV_ADQNTP  TYPE OIB_01_ADQNTP,
        GT_ADQNTP  TYPE STANDARD TABLE OF OIB_01_ADQNTP.

  FIELD-SYMBOLS: <FS_ADQNTP>.

  CLEAR: CP_ENERGY, CP_ENERGY.

  SPLIT  UP_BWART_P  AT '-' INTO TABLE LT_BWART_P.
  CLEAR LV_ERFMG.
  LOOP AT LT_BWART_P INTO LS_BWART_P.
    SELECT SUM( CASE SHKZG
                   WHEN 'S'
                     THEN ERFMG
                   WHEN 'H'
                     THEN
                     ERFMG * ( -1 )
                 END ) AS ERFMG , ERFME
     FROM @GT_MATDOC AS A
*                 FROM MATDOC AS A
      WHERE BUDAT IN @S_BUDAT
        AND WERKS  = @P_WERKS
        AND MATNR  = @UP_MATNR
        AND BWART = @LS_BWART_P
        AND GRUND = @UP_GRUND
     GROUP BY ERFME
     INTO  TABLE @DATA(LT_RESULT_P).
    IF SY-SUBRC <> 0.
      CLEAR: LV_ERFMG.
    ELSE.
      LOOP AT LT_RESULT_P ASSIGNING FIELD-SYMBOL(<FS_RESULT_P>) .
        CASE <FS_RESULT_P>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + (  ABS( <FS_RESULT_P>-ERFMG ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG +   ABS( <FS_RESULT_P>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG +   ABS( <FS_RESULT_P>-ERFMG ).
        ENDCASE.
      ENDLOOP.
      CP_ERFMG = CP_ERFMG + LV_ERFMG.
    ENDIF.

    " Energy
    CLEAR: LV_ADQNTP, GT_ADQNTP[].
    SELECT (GV_FIELD)
        FROM @GT_MATDOC AS A
*      FROM MATDOC AS A
        INNER JOIN MATDOCOIL AS B
        ON A~MBLNR = B~MBLNR
         AND A~MJAHR = B~MJAHR
         AND A~ZEILE = B~ZEILE
       WHERE A~WERKS = @P_WERKS
         AND MATNR = @UP_MATNR
         AND A~BUDAT IN @S_BUDAT
         AND A~BWART = @LS_BWART_P
         AND GRUND = @UP_GRUND
      INTO TABLE @GT_ADQNTP.
    IF SY-SUBRC = 0.
      LOOP AT GT_ADQNTP ASSIGNING <FS_ADQNTP>.           "#EC CI_NESTED
        LV_ADQNTP = LV_ADQNTP + ABS( <FS_ADQNTP> ).
      ENDLOOP.
      CP_ENERGY = CP_ENERGY  + LV_ADQNTP.
    ENDIF.
  ENDLOOP.

  SPLIT  UP_BWART_N  AT '-' INTO TABLE LT_BWART_N.
  CLEAR LV_ERFMG.
  LOOP AT LT_BWART_N INTO LS_BWART_N.
    SELECT SUM( CASE SHKZG
                   WHEN 'S'
                     THEN ERFMG
                   WHEN 'H'
                     THEN
                     ERFMG * ( -1 )
                 END ) AS ERFMG, ERFME
     FROM @GT_MATDOC AS A
*                 FROM MATDOC AS A
      WHERE BUDAT IN @S_BUDAT
        AND WERKS  = @P_WERKS
        AND MATNR  = @UP_MATNR
        AND BWART = @LS_BWART_N
        AND GRUND = @UP_GRUND
        GROUP BY ERFME
      INTO  TABLE @DATA(LT_RESULT_N).
    IF SY-SUBRC <> 0.
      CLEAR: LV_ERFMG.
    ELSE.
      LOOP AT LT_RESULT_N ASSIGNING FIELD-SYMBOL(<FS_RESULT_N>).
        CASE <FS_RESULT_N>-ERFME.
          WHEN 'CM3'.
            LV_ERFMG = LV_ERFMG + ( ABS( <FS_RESULT_N>-ERFMG ) / 1000 ).
          WHEN 'CMK'.
            LV_ERFMG = LV_ERFMG + ABS( <FS_RESULT_N>-ERFMG ).
          WHEN OTHERS.
            LV_ERFMG = LV_ERFMG + ABS( <FS_RESULT_N>-ERFMG ).
        ENDCASE.
      ENDLOOP.
      CP_ERFMG = CP_ERFMG - LV_ERFMG.
    ENDIF. " IF sy-subrc <> 0.

    " Energy
    CLEAR: LV_ADQNTP, GT_ADQNTP[].
    SELECT (GV_FIELD)
        FROM @GT_MATDOC AS A
*      FROM MATDOC AS A
        INNER JOIN MATDOCOIL AS B
        ON A~MBLNR = B~MBLNR
         AND A~MJAHR = B~MJAHR
         AND A~ZEILE = B~ZEILE
       WHERE A~WERKS = @P_WERKS
         AND MATNR = @UP_MATNR
         AND A~BUDAT IN @S_BUDAT
         AND A~BWART = @LS_BWART_N
         AND GRUND = @UP_GRUND
      INTO TABLE @GT_ADQNTP.
    IF SY-SUBRC = 0.
      LOOP AT GT_ADQNTP ASSIGNING <FS_ADQNTP>.           "#EC CI_NESTED
        LV_ADQNTP = LV_ADQNTP + ABS( <FS_ADQNTP> ).
      ENDLOOP.
      CP_ENERGY = CP_ENERGY  - LV_ADQNTP.
    ENDIF.
  ENDLOOP.
ENDFORM. " FRM_GET_QTY_ENERGY_BOILOFF


*&---------------------------------------------------------------------*
*& Form FRM_CONVERT_TO_UOM_ZPDNAC
*&---------------------------------------------------------------------*
FORM FRM_CONVERT_TO_UOM_ZPDNAC USING UP_UNIT_FROM UP_UNIT_TO CHANGING CP_ERFMG.
  IF UP_UNIT_FROM <> UP_UNIT_TO.
    CASE UP_UNIT_FROM.
      WHEN 'CM3'.
        CP_ERFMG = CP_ERFMG / 1000.
      WHEN 'CMK'.
        CP_ERFMG = CP_ERFMG * 1000.
      WHEN OTHERS.
    ENDCASE.
  ENDIF.
ENDFORM. " FRM_CONVERT_TO_UOM_ZPDNAC
*&---------------------------------------------------------------------*
*& Form FRM_SAVE_FG
*&---------------------------------------------------------------------*
*-------------------------------------------------------------*
FORM FRM_SAVE_FG USING UP_DATE TYPE SY-DATUM
                       UP_ERFMG TYPE MATDOC-ERFMG
                       UP_KOSTL TYPE KOSTL.

  DATA: LV_QUESTION TYPE STRING,
        LV_TEXT1    TYPE STRING,
        LV_TEXT2    TYPE STRING,
        LV_ANSWER,
        LS_HEADER   TYPE BAPI2017_GM_HEAD_01,
        LS_CODE     TYPE BAPI2017_GM_CODE,
        LS_HEADRET  TYPE BAPI2017_GM_HEAD_RET,
        LT_ITEM     TYPE TABLE OF BAPI2017_GM_ITEM_CREATE WITH HEADER LINE,
        LV_MSG      TYPE STRING,
        LT_RETURN   TYPE TABLE OF BAPIRET2 WITH HEADER LINE.

  " Check date is filled
  IF UP_DATE IS INITIAL.
    MESSAGE I008(ZPDN002) DISPLAY LIKE 'E'.
    EXIT.
  ENDIF.

  " Check quantity is filled
  IF UP_ERFMG IS INITIAL.
    MESSAGE I009(ZPDN002) DISPLAY LIKE 'E'.
    EXIT.
  ENDIF.

  " Check cost center is filled
  IF UP_KOSTL IS INITIAL.
    MESSAGE I010(ZPDN002) DISPLAY LIKE 'E' WITH TEXT-013.
    EXIT.
  ENDIF.

  " Check authorization to post
  AUTHORITY-CHECK OBJECT 'ZPDN_FG'
   ID 'WERKS' FIELD P_WERKS.
  IF SY-SUBRC <> 0.
    MESSAGE I011(ZPDN002) DISPLAY LIKE 'E' WITH TEXT-013 P_WERKS.
    EXIT.
  ENDIF.

  DATA(LV_DATE) = |{ UP_DATE DATE = USER }|.
  DATA(LV_ERFMG) = UP_ERFMG.

  LV_QUESTION = TEXT-001 && | | &&  LV_ERFMG
     && | | && TEXT-009  &&  | | && UP_KOSTL
     && | | && TEXT-003 && | | && LV_DATE
     && | | &&  TEXT-004 .

  CALL FUNCTION 'POPUP_TO_CONFIRM'
    EXPORTING
      TITLEBAR              = 'Confirmation FG'
      TEXT_QUESTION         = LV_QUESTION
      TEXT_BUTTON_1         = 'Yes'
      ICON_BUTTON_1         = ''
      TEXT_BUTTON_2         = 'No'
      ICON_BUTTON_2         = ' '
      DISPLAY_CANCEL_BUTTON = ''
      START_COLUMN          = 30
      START_ROW             = 2
    IMPORTING
      ANSWER                = LV_ANSWER
    EXCEPTIONS
      TEXT_NOT_FOUND        = 1
      OTHERS                = 2.
  IF SY-SUBRC <> 0.
    CLEAR LV_ANSWER.
  ENDIF.


  CHECK LV_ANSWER = '1'.


  LS_HEADER-PSTNG_DATE = UP_DATE.
  LS_HEADER-DOC_DATE = SY-DATUM.
  LS_HEADER-HEADER_TXT = TEXT-005."'Posted from ZPDND002MT'.

  LS_CODE-GM_CODE = '03'.

  LT_ITEM-MATERIAL = GS_PDNAC_TYPE_5-MATNR.
  LT_ITEM-PLANT = P_WERKS.
  LT_ITEM-MOVE_TYPE = '201'.

  SELECT SINGLE * FROM MARA WHERE MATNR = @GS_PDNAC_TYPE_5-MATNR INTO @DATA(LS_MARA).
  IF SY-SUBRC = 0 .

    IF LS_MARA-MTART = 'ZAUT'.

      LT_ITEM-SPEC_STOCK = 'P'.

    ELSE.

      SELECT SINGLE * FROM MARC WHERE MATNR = @GS_PDNAC_TYPE_5-MATNR INTO @DATA(LS_MARC).
      IF SY-SUBRC = 0.
        IF LS_MARC-LGPRO IS INITIAL.
          MESSAGE | { TEXT-M01  } { GS_PDNAC_TYPE_5-MATNR }| TYPE 'I' DISPLAY LIKE 'E'.
        ELSE.
          LT_ITEM-STGE_LOC = LS_MARC-LGPRO.
        ENDIF.
      ENDIF.

    ENDIF.
  ENDIF.

  LT_ITEM-ENTRY_QNT = UP_ERFMG.
  LT_ITEM-ENTRY_UOM = GS_PDNAC_TYPE_5-BSTME."'CMK'.
  LT_ITEM-COSTCENTER = UP_KOSTL.
  APPEND LT_ITEM.

  CALL FUNCTION 'BAPI_GOODSMVT_CREATE'
    EXPORTING
      GOODSMVT_HEADER  = LS_HEADER
      GOODSMVT_CODE    = LS_CODE
    IMPORTING
      GOODSMVT_HEADRET = LS_HEADRET
    TABLES
      GOODSMVT_ITEM    = LT_ITEM
      RETURN           = LT_RETURN.

  LOOP AT LT_RETURN WHERE TYPE = 'A' OR TYPE = 'E'.
*    MOVE-CORRESPONDING LT_RETURN TO LT_BAPIRET2.
*    APPEND LT_BAPIRET2.
  ENDLOOP.

  IF SY-SUBRC = 0.

    CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.

    DELETE LT_RETURN WHERE TYPE NE 'A' AND TYPE NE 'E'.

    PERFORM FRM_SAVE_LOG TABLES LT_RETURN USING 'ZPDN' 'ZPDND002_DC'.

    CALL FUNCTION 'RSCRMBW_DISPLAY_BAPIRET2'
      TABLES
        IT_RETURN = LT_RETURN[].

  ELSE.

    CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
      EXPORTING
        WAIT = 'X'.

* Envoie de mail à la liste de diffusion GL3Z
* ------------------------
    PERFORM FRM_SEND_EMAIL_GL3ZEXP USING LS_HEADRET-MAT_DOC
                                         LS_HEADRET-DOC_YEAR
                                         LS_HEADER-PSTNG_DATE
                                         LS_HEADER-DOC_DATE
                                         TEXT-S02
                                         P_WERKS.

    PERFORM FRM_SAVE_LOG TABLES LT_RETURN USING 'ZPDN' 'ZPDND002_DC'.

    LV_MSG = TEXT-T23 && | | && LS_HEADRET-MAT_DOC && | | && TEXT-T24.
    MESSAGE LV_MSG TYPE 'S' .

*    LEAVE TO SCREEN 0.

  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form FRM_SAVE_LOG
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*&      --> LT_RETURN
*&---------------------------------------------------------------------*
FORM FRM_SAVE_LOG  TABLES   PT_RETURN STRUCTURE BAPIRET2
                   USING    U_OBJECT  TYPE BALOBJ_D
                            U_SUBOBJ  TYPE BALSUBOBJ.

  DATA: LS_LOG    TYPE BAL_S_LOG,
        LS_MSG    TYPE BAL_S_MSG,
        LS_HANDLE TYPE BALLOGHNDL,
        LT_HANDLE TYPE BAL_T_LOGH.

  LS_LOG-OBJECT     = U_OBJECT. "'ZPDN'.
  LS_LOG-SUBOBJECT =  U_SUBOBJ. "'ZPDND002_DC'.
  LS_LOG-ALDATE     = SY-DATUM.
  LS_LOG-ALTIME     = SY-UZEIT.
  LS_LOG-ALUSER     = SY-UNAME.
  LS_LOG-ALPROG     = SY-REPID.

  CALL FUNCTION 'BAL_LOG_CREATE'
    EXPORTING
      I_S_LOG                 = LS_LOG
    IMPORTING
      E_LOG_HANDLE            = LS_HANDLE
    EXCEPTIONS
      LOG_HEADER_INCONSISTENT = 1
      OTHERS                  = 2.
  IF SY-SUBRC <> 0.
    MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
            WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.

  LOOP AT PT_RETURN.

    LS_MSG-MSGTY = PT_RETURN-TYPE.
    LS_MSG-MSGNO = PT_RETURN-NUMBER.
    LS_MSG-MSGID = PT_RETURN-ID.
    LS_MSG-MSGV1 = PT_RETURN-MESSAGE_V1.
    LS_MSG-MSGV2 = PT_RETURN-MESSAGE_V2.
    LS_MSG-MSGV3 = PT_RETURN-MESSAGE_V3.
    LS_MSG-MSGV4 = PT_RETURN-MESSAGE_V4.

    IF LS_MSG-MSGTY CA 'EA'.
      LS_MSG-PROBCLASS = '1'.
    ELSEIF LS_MSG-MSGTY CA 'S'.
      LS_MSG-PROBCLASS = '2'.
    ELSE .
      LS_MSG-PROBCLASS = '3'.
    ENDIF.

    CALL FUNCTION 'BAL_LOG_MSG_ADD'
      EXPORTING
        I_LOG_HANDLE     = LS_HANDLE
        I_S_MSG          = LS_MSG
        "   IMPORTING
  "     E_S_MSG_HANDLE   =
  "     E_MSG_WAS_LOGGED =
  "     E_MSG_WAS_DISPLAYED       =
      EXCEPTIONS
        LOG_NOT_FOUND    = 1
        MSG_INCONSISTENT = 2
        LOG_IS_FULL      = 3
        OTHERS           = 4.
    IF SY-SUBRC <> 0.
      MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
              WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
    ENDIF.
*    APPEND ls_handle TO lt_handle. "##EC CI_APPEND_OK
    INSERT LS_HANDLE INTO TABLE LT_HANDLE.

  ENDLOOP.

  IF SY-TCODE = 'ZPDND002MT'.

    CALL FUNCTION 'BAL_DB_SAVE'
      EXPORTING
*       I_IN_UPDATE_TASK = 'X'
        I_SAVE_ALL       = 'X'
        I_T_LOG_HANDLE   = LT_HANDLE[]
      EXCEPTIONS
        LOG_NOT_FOUND    = 1
        SAVE_NOT_ALLOWED = 2
        NUMBERING_ERROR  = 3
        OTHERS           = 4.
  ELSE.

    CALL FUNCTION 'BAL_DB_SAVE'
      EXPORTING
        I_IN_UPDATE_TASK = 'X'
*       I_SAVE_ALL       = 'X'
        I_T_LOG_HANDLE   = LT_HANDLE[]
      EXCEPTIONS
        LOG_NOT_FOUND    = 1
        SAVE_NOT_ALLOWED = 2
        NUMBERING_ERROR  = 3
        OTHERS           = 4.
  ENDIF.

  IF SY-SUBRC <> 0.
    MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
            WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.

  REFRESH LT_HANDLE.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form FRM_SEND_EMAIL_GL3ZEXP
*&---------------------------------------------------------------------*
FORM FRM_SEND_EMAIL_GL3ZEXP USING U_MAT_DOC    TYPE MBLNR
                                  U_DOC_YEAR   TYPE MJAHR
                                  U_PSTNG_DATE TYPE BUDAT
                                  U_DOC_DATE   TYPE BLDAT
                                  U_STATUS     TYPE STRING
                                  U_WERKS      TYPE WERKS_D.

  DATA : LS_ZPDND002_DIFFUS LIKE LINE OF GT_ZPDND002_DIFFUS.


  DATA : LV_ID               TYPE THEAD-TDID.
  DATA : LV_LANGUAGE         TYPE THEAD-TDSPRAS.
  DATA : LV_NAME             TYPE THEAD-TDNAME.
  DATA : LV_OBJECT           TYPE THEAD-TDOBJECT.
  DATA : LS_HEADER           TYPE THEAD.
  DATA : LT_LINES            TYPE STANDARD TABLE OF TLINE.

  DATA : LV_DATATAB_X          TYPE TDTAB_X256.
  DATA : LV_DATATAB_C          TYPE TDTAB_C132.

  DATA : LO_DOCUMENT         TYPE REF TO CL_DOCUMENT_BCS.
  DATA : LV_DOCUMENT_BCS     TYPE REF TO CX_DOCUMENT_BCS.
  DATA : LO_SEND_REQUEST     TYPE REF TO CL_BCS.
  DATA : LO_SENDER           TYPE REF TO IF_SENDER_BCS.
  DATA : LV_SENT_TO_ALL      TYPE OS_BOOLEAN  .
  DATA : LT_ATT_CONTENT_HEX  TYPE SOLIX_TAB .
  DATA : LT_MESSAGE_BODY     TYPE BCSY_TEXT.
  DATA : LV_SENDER_EMAIL     TYPE ADR6-SMTP_ADDR.
  DATA : LV_RECIPIENT_EMAIL  TYPE ADR6-SMTP_ADDR.
  DATA : LO_RECIPIENT        TYPE REF TO IF_RECIPIENT_BCS  .
  DATA : LV_WITH_ERROR_SCREEN  TYPE OS_BOOLEAN .
  DATA : LV_LENGTH_MIME      TYPE NUM12.
  DATA : LV_MIME_TYPE        TYPE W3CONTTYPE.
  DATA : LT_ATTACHMENT       TYPE SOLIX_TAB.
  DATA : LV_ATTACHMENT_TYPE  TYPE SOODK-OBJTP.
  DATA : LV_ATTACHMENT_SIZE  TYPE SOOD-OBJLEN.
  DATA : LV_ATTACHMENT_SUBJECT TYPE SOOD-OBJDES.
  DATA : LT_BODY_HEX         TYPE SOLIX_TAB.
  DATA : LT_BODY_HEX2        TYPE SOLIX_TAB.
  DATA : LV_BODY_TXT         TYPE STRING.
  DATA : LV_XHTML_STRING     TYPE XSTRING.

  DATA : LT_RETURN   TYPE TABLE OF BAPIRET2 WITH HEADER LINE.


  DATA : LV_MAIL_SUBJECT     TYPE  SO_OBJ_DES.
  DATA : LV_TYPE             TYPE STRING.
  DATA : LV_EXTENSION        TYPE STRING.
  DATA : LV_DOCID_STR(12)    .
  DATA : LV_MSG              TYPE STRING.

* READ EMAIL FORM ZPDND002_DIFFUS.
*---------------------------
  SELECT * FROM ZPDND002_DIFFUS INTO TABLE GT_ZPDND002_DIFFUS WHERE WERKS = U_WERKS.
  IF SY-SUBRC <> 0.
    CLEAR GT_ZPDND002_DIFFUS[].
  ENDIF.

  IF GT_ZPDND002_DIFFUS IS INITIAL.
    RETURN.
  ENDIF.

* Send email
* ------------------------
  TRY.
      LO_SEND_REQUEST = CL_BCS=>CREATE_PERSISTENT( ).
    CATCH CX_SEND_REQ_BCS.
  ENDTRY.
  " Set the subjest of email
  LV_MAIL_SUBJECT =   |{ TEXT-S01 }|.
  REPLACE 'MM'   WITH U_PSTNG_DATE+4(2) INTO LV_MAIL_SUBJECT.
  REPLACE 'YYYY' WITH U_PSTNG_DATE+0(4) INTO LV_MAIL_SUBJECT.
  REPLACE 'STATUS' WITH U_STATUS INTO LV_MAIL_SUBJECT.
  REPLACE 'WERKS' WITH U_WERKS INTO LV_MAIL_SUBJECT.

  " Send in HTML format
  LV_ID = 'ST'.
  LV_LANGUAGE = SY-LANGU.
  LV_NAME = 'ZPDND002_EMAIL_GL3ZEXP'.
  LV_OBJECT = 'TEXT'.

  CALL FUNCTION 'READ_TEXT'
    EXPORTING
      ID                      = LV_ID
      LANGUAGE                = LV_LANGUAGE
      NAME                    = LV_NAME
      OBJECT                  = LV_OBJECT
    IMPORTING
      HEADER                  = LS_HEADER
    TABLES
      LINES                   = LT_LINES
    EXCEPTIONS
      ID                      = 1
      LANGUAGE                = 2
      NAME                    = 3
      NOT_FOUND               = 4
      OBJECT                  = 5
      REFERENCE_CHECK         = 6
      WRONG_ACCESS_TO_ARCHIVE = 7
      OTHERS                  = 8.
  IF SY-SUBRC <> 0.
    MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
            WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.

  CALL FUNCTION 'CONVERT_ITF_TO_ASCII'
    EXPORTING
      LANGUAGE          = SY-LANGU
      TABLETYPE         = 'ASC'
    IMPORTING
      X_DATATAB         = LV_DATATAB_X
      C_DATATAB         = LV_DATATAB_C
    TABLES
      ITF_LINES         = LT_LINES
    EXCEPTIONS
      INVALID_TABLETYPE = 1
      OTHERS            = 2.
  IF SY-SUBRC <> 0.
    MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
                     WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.
*
  LOOP AT LV_DATATAB_C INTO DATA(LV_LINE).
    LV_BODY_TXT = |{ LV_BODY_TXT }| && |{ LV_LINE }|.
  ENDLOOP.


  DATA : LV_WERKS TYPE STRING.
  SELECT SINGLE NAME1 FROM T001W INTO LV_WERKS WHERE WERKS = U_WERKS.
  IF SY-SUBRC <> 0. "    TC, 21/12/2023, Added code
    CLEAR LV_WERKS.
  ENDIF.
  LV_WERKS = |{ U_WERKS }| && | | && |{ LV_WERKS }| && | |.
  REPLACE 'WERKS'   WITH LV_WERKS INTO LV_BODY_TXT.


  DATA : LV_PSTNG_DATE TYPE STRING.
  LV_PSTNG_DATE = |{ U_PSTNG_DATE+4(2) }| && |/| && |{ U_PSTNG_DATE+0(4) }|.
  REPLACE 'PSTNG_DATE'   WITH LV_PSTNG_DATE INTO LV_BODY_TXT.


  DATA : LV_DOC_DATE TYPE STRING.
  LV_DOC_DATE = |{ U_DOC_DATE+6(2) }| && |/| && |{ U_DOC_DATE+4(2) }| && |/| && |{ U_DOC_DATE+0(4) }|.
  REPLACE 'DOC_DATE'   WITH LV_DOC_DATE INTO LV_BODY_TXT.


  DATA : LV_DOC_TIME TYPE STRING.
  LV_DOC_TIME = |{ SY-UZEIT+0(2) }| && |:| && |{ SY-UZEIT+2(2) }| && | |.
  REPLACE 'DOC_TIME'   WITH LV_DOC_TIME INTO LV_BODY_TXT.


  DATA : LV_SYUNAME TYPE STRING.
  DATA : LS_USER_DATA TYPE USR03.
  CALL FUNCTION 'SUSR_USER_ADDRESS_READ'
    EXPORTING
      USER_NAME              = SY-UNAME
    IMPORTING
      USER_USR03             = LS_USER_DATA
    EXCEPTIONS
      USER_ADDRESS_NOT_FOUND = 1
      OTHERS                 = 2.

  IF SY-SUBRC = 0.

    LV_SYUNAME = |{ SY-UNAME }| && | / | && |{ LS_USER_DATA-NAME1 }| && | | && |{ LS_USER_DATA-NAME2 }|.
    REPLACE 'SYUNAME'   WITH LV_SYUNAME INTO LV_BODY_TXT.

  ENDIF.


  DATA : LV_STATUS TYPE STRING.
  LV_STATUS = U_STATUS.
  REPLACE ALL OCCURRENCES OF 'STATUS' IN LV_BODY_TXT WITH LV_STATUS.


  CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
    EXPORTING
      TEXT   = LV_BODY_TXT
    IMPORTING
      BUFFER = LV_XHTML_STRING
    EXCEPTIONS
      FAILED = 1
      OTHERS = 2.
  IF SY-SUBRC <> 0.
    CLEAR LV_XHTML_STRING.
  ENDIF.

  CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
    EXPORTING
      BUFFER     = LV_XHTML_STRING
    TABLES
      BINARY_TAB = LT_BODY_HEX.
  TRY.
      LO_DOCUMENT = CL_DOCUMENT_BCS=>CREATE_DOCUMENT(
                           I_TYPE    = 'HTM'
                           I_HEX    = LT_BODY_HEX
                           I_SUBJECT = LV_MAIL_SUBJECT ) .
    CATCH CX_DOCUMENT_BCS INTO LV_DOCUMENT_BCS.
  ENDTRY.
  "Email TO...
  LOOP AT GT_ZPDND002_DIFFUS INTO LS_ZPDND002_DIFFUS WHERE WERKS = U_WERKS.
    LV_RECIPIENT_EMAIL = LS_ZPDND002_DIFFUS-MEMBER_ADR.
    TRY. "  TC, 21/12/2023, Added code
        LO_RECIPIENT = CL_CAM_ADDRESS_BCS=>CREATE_INTERNET_ADDRESS( LV_RECIPIENT_EMAIL ).
      CATCH CX_ADDRESS_BCS.
    ENDTRY.
    TRY.
        CALL METHOD LO_SEND_REQUEST->ADD_RECIPIENT
          EXPORTING
            I_RECIPIENT = LO_RECIPIENT
            I_EXPRESS   = 'X'.
      CATCH CX_SEND_REQ_BCS.
    ENDTRY.
  ENDLOOP.


  " Email FROM...
  LV_SENDER_EMAIL = 'no-reply@sonatrach.dz'.
  TRY.
      LO_SENDER = CL_CAM_ADDRESS_BCS=>CREATE_INTERNET_ADDRESS( LV_SENDER_EMAIL ).
    CATCH CX_ADDRESS_BCS.
  ENDTRY.
  TRY.
      LO_SEND_REQUEST->SET_SENDER( LO_SENDER ).
    CATCH CX_SEND_REQ_BCS.
  ENDTRY.
  TRY .
      CALL METHOD LO_SEND_REQUEST->SET_SENDER
        EXPORTING
          I_SENDER = LO_SENDER.
      LO_SEND_REQUEST->SET_SENDER( LO_SENDER ).

    CATCH CX_SEND_REQ_BCS.
  ENDTRY.


  " assign document to the send request:
  TRY.
      LO_SEND_REQUEST->SET_DOCUMENT( LO_DOCUMENT ).
    CATCH CX_SEND_REQ_BCS. " BCS: Send Request Exceptions  "  TC, 21/12/2023, Added code
  ENDTRY.
  " set send immediately flag
  TRY.
      LO_SEND_REQUEST->SET_SEND_IMMEDIATELY( 'X' ).
    CATCH CX_SEND_REQ_BCS. " BCS: Send Request Exceptions
  ENDTRY.

  " SAP Send Email CL_BCS
  MOVE SPACE TO LV_WITH_ERROR_SCREEN.
  TRY.
      LV_SENT_TO_ALL  = LO_SEND_REQUEST->SEND( LV_WITH_ERROR_SCREEN ).
    CATCH CX_SEND_REQ_BCS.
  ENDTRY.
  IF SY-TCODE = 'ZPDND002MT'.
    COMMIT WORK.
  ENDIF.

  " Save log

  IF LV_SENT_TO_ALL = 'X'.

    LT_RETURN-TYPE = 'S'.
    LT_RETURN-ID = 'ZPDN001'.
    LT_RETURN-NUMBER = '022'.

    LT_RETURN-MESSAGE_V1 = |{ U_PSTNG_DATE+4(2) }| && |/| && |{ U_PSTNG_DATE+0(4) }| .
    LT_RETURN-MESSAGE_V2 = U_STATUS.
    LT_RETURN-MESSAGE_V3 = U_MAT_DOC.
    LT_RETURN-MESSAGE_V4 = U_DOC_YEAR.
    APPEND LT_RETURN.

    PERFORM FRM_SAVE_LOG TABLES LT_RETURN USING 'ZPDN' 'ZPDND002_EMAIL'.
  ENDIF.

ENDFORM.
*&---------------------------------------------------------------------*
*& Form FRM_POST_AC_TO_COST_CENTER
*&---------------------------------------------------------------------*
FORM FRM_POST_AC_TO_COST_CENTER USING UP_DATE TYPE SY-DATUM
                                      UP_ERFMG TYPE ERFMG
                                      UP_KOSTL TYPE KOSTL.
  DATA: LV_QUESTION TYPE STRING,
        LV_TEXT1    TYPE STRING,
        LV_TEXT2    TYPE STRING,
        LV_ANSWER,
        LS_HEADER   TYPE BAPI2017_GM_HEAD_01,
        LS_CODE     TYPE BAPI2017_GM_CODE,
        LS_HEADRET  TYPE BAPI2017_GM_HEAD_RET,
        LT_ITEM     TYPE TABLE OF BAPI2017_GM_ITEM_CREATE WITH HEADER LINE,
        LV_MSG      TYPE STRING,
        LT_RETURN   TYPE TABLE OF BAPIRET2 WITH HEADER LINE.

  " Check date is filled
  IF UP_DATE IS INITIAL.
    MESSAGE I008(ZPDN002) DISPLAY LIKE 'E'.
    EXIT.
  ENDIF.
  "Check quantity is not initial
  IF UP_ERFMG IS INITIAL.
    MESSAGE I013(ZPDN002) DISPLAY LIKE 'E'.
    EXIT.
  ENDIF.

  " Check cost center is filled
  IF UP_KOSTL IS INITIAL.
    MESSAGE I010(ZPDN002) DISPLAY LIKE 'E' WITH TEXT-014."'AC'.
    EXIT.
  ENDIF.

  " Check authorization to post
  AUTHORITY-CHECK OBJECT 'ZPDN_AC'
   ID 'WERKS' FIELD P_WERKS.
  IF SY-SUBRC <> 0.
    MESSAGE I011(ZPDN002) DISPLAY LIKE 'E' WITH TEXT-014 P_WERKS." AC
    EXIT.
  ENDIF.

  DATA(LV_DATE) = |{ UP_DATE DATE = USER }|.
  DATA(LV_ERFMG) = UP_ERFMG.

  LV_QUESTION = TEXT-001 && | | &&  LV_ERFMG
       && | | && TEXT-012 && | | && UP_KOSTL
       && | | && TEXT-003 && | | && LV_DATE
       && | | &&  TEXT-004.

  CALL FUNCTION 'POPUP_TO_CONFIRM'
    EXPORTING
      TITLEBAR              = 'Confirmation FG'
      TEXT_QUESTION         = LV_QUESTION
      TEXT_BUTTON_1         = 'Yes'
      ICON_BUTTON_1         = ''
      TEXT_BUTTON_2         = 'No'
      ICON_BUTTON_2         = ' '
      DISPLAY_CANCEL_BUTTON = ''
      START_COLUMN          = 30
      START_ROW             = 2
    IMPORTING
      ANSWER                = LV_ANSWER
    EXCEPTIONS
      TEXT_NOT_FOUND        = 1
      OTHERS                = 2.
  IF SY-SUBRC <> 0.
    CLEAR LV_ANSWER.
  ENDIF.

  CHECK LV_ANSWER = '1'.

  LS_HEADER-PSTNG_DATE = UP_DATE.
  LS_HEADER-DOC_DATE = SY-DATUM.
  LS_HEADER-HEADER_TXT = TEXT-005 ."'Posted from ZPDND002MT'.

  LS_CODE-GM_CODE = '03'.

  LT_ITEM-MATERIAL = GS_PDNAC_TYPE_3-MATNR.
  LT_ITEM-PLANT = P_WERKS.
  LT_ITEM-MOVE_TYPE = '201'.

  SELECT SINGLE * FROM MARA WHERE MATNR = @GS_PDNAC_TYPE_3-MATNR INTO @DATA(LS_MARA).
  IF SY-SUBRC = 0 .

    IF LS_MARA-MTART = 'ZAUT'.

      LT_ITEM-SPEC_STOCK = 'P'.

    ELSE.

      SELECT SINGLE * FROM MARC WHERE MATNR = @GS_PDNAC_TYPE_3-MATNR INTO @DATA(LS_MARC).
      IF SY-SUBRC = 0.
        IF LS_MARC-LGPRO IS INITIAL.
          MESSAGE | { TEXT-M01  } { GS_PDNAC_TYPE_3-MATNR }| TYPE 'I' DISPLAY LIKE 'E'.
          EXIT.
        ELSE.
          LT_ITEM-STGE_LOC = LS_MARC-LGPRO.
        ENDIF.
      ENDIF.

    ENDIF.
  ENDIF.

  LT_ITEM-ENTRY_QNT = UP_ERFMG.
  LT_ITEM-ENTRY_UOM = GS_PDNAC_TYPE_3-BSTME."'CMK'.
  LT_ITEM-COSTCENTER = UP_KOSTL.
  APPEND LT_ITEM.

  CALL FUNCTION 'BAPI_GOODSMVT_CREATE'
    EXPORTING
      GOODSMVT_HEADER  = LS_HEADER
      GOODSMVT_CODE    = LS_CODE
    IMPORTING
      GOODSMVT_HEADRET = LS_HEADRET
    TABLES
      GOODSMVT_ITEM    = LT_ITEM
      RETURN           = LT_RETURN.

  LOOP AT LT_RETURN WHERE TYPE = 'A' OR TYPE = 'E'.
*    MOVE-CORRESPONDING LT_RETURN TO LT_BAPIRET2.
*    APPEND LT_BAPIRET2.
  ENDLOOP.

  IF SY-SUBRC = 0.

    CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.

    DELETE LT_RETURN WHERE TYPE NE 'A' AND TYPE NE 'E'.

    PERFORM FRM_SAVE_LOG TABLES LT_RETURN USING 'ZPDN' 'ZPDND002_DC'.

    CALL FUNCTION 'RSCRMBW_DISPLAY_BAPIRET2'
      TABLES
        IT_RETURN = LT_RETURN[].

  ELSE.

    CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
      EXPORTING
        WAIT = 'X'.


    PERFORM FRM_SAVE_LOG TABLES LT_RETURN USING 'ZPDN' 'ZPDND002_DC'.

    LV_MSG = TEXT-T23 && | | && LS_HEADRET-MAT_DOC && | | && TEXT-T24.
    MESSAGE LV_MSG TYPE 'S' .

*    LEAVE TO SCREEN 0.
  ENDIF.

ENDFORM.


FORM FRM_SEND_EMAIL_GL3ZEXP_MIGO TABLES     PT_XMKPF STRUCTURE MKPF
                                            PT_XMSEG STRUCTURE MSEG.


  IF PT_XMKPF[] IS NOT INITIAL.

    READ TABLE PT_XMKPF INTO DATA(LS_XMKPF) INDEX 1 .
    CONDENSE LS_XMKPF-BKTXT.

*    IF LS_XMKPF-BKTXT = 'Confirmer via ZPDND002'.

    SELECT *
      FROM ZPDNAC
      INTO TABLE @DATA(LT_ZPDNAC)
      WHERE ZTYPE = '5'.

    IF SY-SUBRC EQ 0.

      LOOP AT PT_XMSEG INTO DATA(LS_XMSEG) WHERE BWART = '202' AND SOBKZ = 'P'.

        READ TABLE LT_ZPDNAC INTO DATA(LS_ZPDNAC) WITH KEY MATNR = LS_XMSEG-MATNR
                                                                   ZFGCC = LS_XMSEG-KOSTL
                                                                   WERKS = LS_XMSEG-WERKS
                                                                   ZTYPE = 5.
        IF SY-SUBRC EQ 0.

          PERFORM FRM_SEND_EMAIL_GL3ZEXP USING LS_XMKPF-MBLNR
                                               LS_XMKPF-MJAHR
                                               LS_XMKPF-BUDAT
                                               LS_XMKPF-CPUDT
                                               TEXT-S03
                                               LS_XMSEG-WERKS.
          EXIT.
        ENDIF.

      ENDLOOP.

    ENDIF.
  ENDIF.
ENDFORM.


*&---------------------------------------------------------------------*
*& Form FRM_SAVE_AC
*&---------------------------------------------------------------------*
FORM FRM_SAVE_AC_NEW USING UP_DATE TYPE SY-DATUM .

  DATA:LV_QUESTION       TYPE STRING,
       LV_TEXT1          TYPE STRING,
       LV_TEXT2          TYPE STRING,
       LV_AUFNR          TYPE AUFNR,
       LT_TIMETICKETS    TYPE TABLE OF BAPI_PI_TIMETICKET1 WITH HEADER LINE,
       LT_GOODSMOVEMENTS TYPE TABLE OF BAPI2017_GM_ITEM_CREATE WITH HEADER LINE,
       LT_LINK           TYPE TABLE OF BAPI_LINK_CONF_GOODSMOV WITH HEADER LINE,
       LT_RETURN         TYPE TABLE OF BAPI_CORU_RETURN WITH HEADER LINE,
       LT_BAPIRET2       TYPE TABLE OF BAPIRET2 WITH HEADER LINE,
       LS_RETURN         TYPE BAPIRET1,
       LV_FLAG,
       LV_ANSWER,
       LV_AUTHORIZED     TYPE ABAP_BOOL VALUE ABAP_FALSE.
  DATA: GT_RANGE_CONTENT TYPE SOI_GENERIC_TABLE.
  DATA: LV_QTY TYPE F.

  " Check date is filled
  IF UP_DATE IS INITIAL.
    MESSAGE I008(ZPDN002) DISPLAY LIKE 'E'.
    EXIT.
  ENDIF.


  " Read quantities from excel
  CLEAR: GT_RANGE_CONTENT[].
  REFRESH T_RANGES.
  CLEAR T_RANGES.
  T_RANGES-NAME = GS_RANGE_TAB4-NAME.
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

  PERFORM FRM_GET_EXCEL_DATA  TABLES GT_ZPDND002MT_ALV2_C[]
                        GT_RANGE_CONTENT[]
                USING   GS_RANGE_TAB4
                        'ZPDND002MT_ALV4_C'.

  DATA LS_MAPPING TYPE TY_MAPPING_TRAINS.

  READ TABLE GT_ZPDND002MT_ALV2_C INDEX 1 ASSIGNING FIELD-SYMBOL(<FS>).
  IF SY-SUBRC = 0.
    LS_MAPPING-OLD_TRAIN = 'T1'.
    LS_MAPPING-NEW_TRAIN = <FS>-T1.
    APPEND LS_MAPPING TO GT_MAPPING.
    LS_MAPPING-OLD_TRAIN = 'T2'.
    LS_MAPPING-NEW_TRAIN = <FS>-T2.
    APPEND LS_MAPPING TO GT_MAPPING.
    LS_MAPPING-OLD_TRAIN = 'T3'.
    LS_MAPPING-NEW_TRAIN = <FS>-T3.
    APPEND LS_MAPPING TO GT_MAPPING.
    LS_MAPPING-OLD_TRAIN = 'T4'.
    LS_MAPPING-NEW_TRAIN = <FS>-T4.
    APPEND LS_MAPPING TO GT_MAPPING.
    LS_MAPPING-OLD_TRAIN = 'T5'.
    LS_MAPPING-NEW_TRAIN = <FS>-T5.
    APPEND LS_MAPPING TO GT_MAPPING.
    LS_MAPPING-OLD_TRAIN = 'T6'.
    LS_MAPPING-NEW_TRAIN = <FS>-T6.
    APPEND LS_MAPPING TO GT_MAPPING.
    LS_MAPPING-OLD_TRAIN = 'T7'.
    LS_MAPPING-NEW_TRAIN = <FS>-T7.
    APPEND LS_MAPPING TO GT_MAPPING.
    LS_MAPPING-OLD_TRAIN = 'T8'.
    LS_MAPPING-NEW_TRAIN = <FS>-T8.
    APPEND LS_MAPPING TO GT_MAPPING.
    LS_MAPPING-OLD_TRAIN = 'T9'.
    LS_MAPPING-NEW_TRAIN = <FS>-T9.
    APPEND LS_MAPPING TO GT_MAPPING.
    LS_MAPPING-OLD_TRAIN = 'UTILITES'.
    LS_MAPPING-NEW_TRAIN = <FS>-UTILITES.
    APPEND LS_MAPPING TO GT_MAPPING.
  ENDIF.

  LOOP AT GT_ORDERS INTO DATA(LS_TRAINS).
    CLEAR: LT_TIMETICKETS[],
           LT_GOODSMOVEMENTS[],
           LT_LINK[],
           LT_RETURN[],
           LT_BAPIRET2[],
           LS_RETURN.

    DATA(LV_FIELD) = LS_TRAINS-TRAIN.
    READ TABLE GT_ZPDND002MT_ALV2_C INDEX 8 ASSIGNING FIELD-SYMBOL(<FF>).
    IF SY-SUBRC = 0.

      CLEAR LS_MAPPING.
      READ TABLE GT_MAPPING INTO LS_MAPPING WITH KEY OLD_TRAIN = LV_FIELD.
      IF SY-SUBRC = 0.
        DATA(LV_FIELD_NEW) = LS_MAPPING-NEW_TRAIN(2).
      ENDIF.

      ASSIGN COMPONENT LV_FIELD OF STRUCTURE <FF> TO FIELD-SYMBOL(<FS_FIELD>).

      " Check authority of posting

      REPLACE ',' WITH '.' INTO <FS_FIELD>.
      " Check order quantity AC is not initial
      LV_QTY  = <FS_FIELD>.
      IF LV_QTY <= 0 OR <FS_FIELD> IS INITIAL.
        MESSAGE I012(ZPDN002) DISPLAY LIKE 'E' WITH LV_FIELD_NEW .
        CONTINUE.
      ENDIF.

      CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
        EXPORTING
          INPUT  = LS_TRAINS-AUFNR
        IMPORTING
          OUTPUT = LV_AUFNR.

      DATA(LV_DATE) = |{ UP_DATE DATE = USER }|.
      DATA LV_ERFMG TYPE MENGE_D.
      LV_ERFMG =  <FS_FIELD>.

      LV_QUESTION = TEXT-001 && | | &&  LV_ERFMG
      && || &&      TEXT-002 && | | && LV_AUFNR
      && || &&      TEXT-011 && | | && LV_FIELD_NEW
      && || &&      TEXT-003 && | | && LV_DATE
      && || &&      TEXT-004 .

      CALL FUNCTION 'POPUP_TO_CONFIRM'
        EXPORTING
          TITLEBAR              = 'Confirmation AC'
          TEXT_QUESTION         = LV_QUESTION
          TEXT_BUTTON_1         = 'Yes'
          ICON_BUTTON_1         = ' '
          TEXT_BUTTON_2         = 'No'
          ICON_BUTTON_2         = ' '
          DISPLAY_CANCEL_BUTTON = ''
          START_COLUMN          = 30
          START_ROW             = 2
        IMPORTING
          ANSWER                = LV_ANSWER
        EXCEPTIONS
          TEXT_NOT_FOUND        = 1
          OTHERS                = 2.
      IF SY-SUBRC <> 0.
        CLEAR LV_ANSWER.
      ENDIF.
      IF LV_ANSWER <> '1'.
        CONTINUE.
      ELSE.

        DATA LV_PHASE TYPE VORNR.
        LV_PHASE = '0102'.
        LT_TIMETICKETS-ORDERID = LS_TRAINS-AUFNR.
        LT_TIMETICKETS-PHASE = LV_PHASE.
        LT_TIMETICKETS-POSTG_DATE = UP_DATE.
        LT_TIMETICKETS-CONF_TEXT = TEXT-005 ." 'Posted from ZPDND002MT'.
        APPEND LT_TIMETICKETS.

        LT_GOODSMOVEMENTS-MATERIAL = GS_PDNAC_TYPE_3-MATNR.
        LT_GOODSMOVEMENTS-PLANT = P_WERKS.
        LT_GOODSMOVEMENTS-ENTRY_QNT =  LV_ERFMG.
        LT_GOODSMOVEMENTS-ENTRY_UOM = GS_PDNAC_TYPE_3-BSTME.
        LT_GOODSMOVEMENTS-MOVE_TYPE  = '261'.

        SELECT SINGLE * FROM MARA WHERE MATNR = @GS_PDNAC_TYPE_3-MATNR INTO @DATA(LS_MARA).
        IF SY-SUBRC = 0 .
          IF LS_MARA-MTART = 'ZAUT'.
            LT_GOODSMOVEMENTS-SPEC_STOCK = 'P'.
          ENDIF.
        ENDIF.

        SELECT SINGLE * FROM MARC WHERE MATNR = @GS_PDNAC_TYPE_3-MATNR INTO @DATA(LS_MARC).
        IF SY-SUBRC = 0 .
          LT_GOODSMOVEMENTS-STGE_LOC = LS_MARC-LGPRO.
        ENDIF.
        APPEND LT_GOODSMOVEMENTS.

        LT_LINK-INDEX_CONFIRM = 1.
        LT_LINK-INDEX_GOODSMOV = 1.
        LT_LINK-INDEX_GM_DEPEND = 0.
        APPEND LT_LINK.

        CALL FUNCTION 'BAPI_PROCORDCONF_CREATE_TT'
          IMPORTING
            RETURN             = LS_RETURN
          TABLES
            TIMETICKETS        = LT_TIMETICKETS
            GOODSMOVEMENTS     = LT_GOODSMOVEMENTS
            LINK_CONF_GOODSMOV = LT_LINK
            DETAIL_RETURN      = LT_RETURN.

        CLEAR LV_FLAG.
        LOOP AT LT_RETURN .

          MOVE-CORRESPONDING LT_RETURN TO LT_BAPIRET2.
          APPEND LT_BAPIRET2.

          IF LT_RETURN-TYPE = 'A' OR LT_RETURN-TYPE = 'E'.
            LV_FLAG = ABAP_ON.
          ENDIF.

        ENDLOOP.

        IF LV_FLAG = ABAP_ON.

          CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.

          CALL FUNCTION 'RSCRMBW_DISPLAY_BAPIRET2'
            TABLES
              IT_RETURN = LT_BAPIRET2[].
        ELSE.

          CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
            EXPORTING
              WAIT = 'X'.
          READ TABLE LT_RETURN WITH KEY NUMBER = '100' ID = 'RU'.
          IF SY-SUBRC = 0.
            MESSAGE LT_RETURN-MESSAGE TYPE 'S'.
          ENDIF.

        ENDIF.

        PERFORM FRM_SAVE_LOG TABLES LT_BAPIRET2 USING 'ZPDN' 'ZPDND002_DC'.

      ENDIF.

    ENDIF. " IF SY-SUBRC = 0.
  ENDLOOP. " LOOP AT GT_ORDERS INTO DATA(LS_TRAINS).


*  DATA LV_KOSTL TYPE KOSTL.
  DATA LV_QTY_UTILITIES TYPE CHAR40.
  DATA LV_QTY_DEC TYPE ERFMG.

  READ TABLE GT_ZPDND002MT_ALV2_C INDEX 8 INTO DATA(LS_UTILITIES).
  IF SY-SUBRC = 0.
    READ TABLE GT_MAPPING WITH KEY NEW_TRAIN = 'Utilités' INTO LS_MAPPING.
    IF SY-SUBRC = 0.
      DATA(LV_FIELD_OLD) = LS_MAPPING-OLD_TRAIN.
      ASSIGN COMPONENT LV_FIELD_OLD OF STRUCTURE LS_UTILITIES TO FIELD-SYMBOL(<FS_NEW>).

      LV_QTY_UTILITIES = <FS_NEW>.

      REPLACE ',' WITH '.' INTO LV_QTY_UTILITIES.
      LV_QTY_DEC = LV_QTY_UTILITIES.

      PERFORM FRM_POST_AC_TO_COST_CENTER USING UP_DATE
                                               LV_QTY_DEC
                                                GV_KOSTL_AC.

    ENDIF.
  ENDIF.
ENDFORM.