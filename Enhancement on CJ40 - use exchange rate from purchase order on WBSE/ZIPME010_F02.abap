*&---------------------------------------------------------------------*
*& Include          ZIPME010_F02
*&---------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*&      Form  FRM_MANAGE_EXCHANGE_RATE
*&---------------------------------------------------------------------*
FORM FRM_MANAGE_EXCHANGE_RATE USING PV_TCODE          TYPE SY-TCODE
                                    PV_SHOW           TYPE BPIN-SHOW
                                    PS_PROJ           TYPE PROJ
                                    PS_PRPS           TYPE PRPS
                                    PV_PROJ           TYPE PS_INTNR
                                    PV_POST1          TYPE PS_POST1
                                    PV_FROM_CURR      TYPE BPIN-TWAER
                                    PV_TO_CURR        TYPE BPIN-TWAER
                           CHANGING PV_EX_RATE        TYPE KURSF
                                    PV_FILLED_EX_RATE TYPE ABAP_BOOL.

  DATA: LV_CONTINUE         TYPE ABAP_BOOL.
  "By default pv_filled_ex_rate = TRUE
  PV_FILLED_EX_RATE = ABAP_TRUE.

  CHECK PV_FROM_CURR IS NOT INITIAL
    AND PV_TO_CURR   IS NOT INITIAL
    AND PV_FROM_CURR <> PV_TO_CURR.

  PERFORM FRM_CHECK_TCODE USING PV_TCODE
                       CHANGING LV_CONTINUE.

  CHECK LV_CONTINUE = ABAP_TRUE.

  DATA PROJ_ID TYPE PRPS-POSID.

  RANGES:  LT_WBS_RANGE FOR PRPS-POSID.
  CLEAR LT_WBS_RANGE.


  IF PS_PRPS-POSID IS NOT INITIAL.
    LT_WBS_RANGE-SIGN = 'I'.
    LT_WBS_RANGE-OPTION = 'EQ'.
    LT_WBS_RANGE-LOW = PS_PRPS-POSID.
    APPEND LT_WBS_RANGE.
  ENDIF.

  SELECT * FROM PRPS
        INTO TABLE @DATA(LT_WBS)
        WHERE POSID IN @LT_WBS_RANGE
          AND PSPHI = @PS_PROJ-PSPNR.
  IF SY-SUBRC <> 0.
    CLEAR LT_WBS[].
  ENDIF.
  CLEAR GT_ZIPME010_EX_RATE[].
  SELECT *
      FROM ZIPME010_EX_RATE
      INTO TABLE @GT_ZIPME010_EX_RATE
    FOR ALL ENTRIES IN @LT_WBS
     WHERE WBSE = @LT_WBS-PSPNR
      AND ACTIVE = 'X'
        AND WAERS = @PV_FROM_CURR
         AND HWAER = @PV_TO_CURR.
  IF SY-SUBRC <> 0.
    CLEAR GT_ZIPME010_EX_RATE[].
  ENDIF.

  PERFORM FRM_DISPLAY_POPUP_ALV .
ENDFORM. " FRM_MANAGE_EXCHANGE_RATE

*&---------------------------------------------------------------------*
*&      Form  FRM_DISPLAY_POPUP_SCREEN
*&---------------------------------------------------------------------*
FORM FRM_DISPLAY_POPUP_ALV.

  IF GT_ZIPME010_EX_RATE[] IS NOT INITIAL.
    MOVE-CORRESPONDING GT_ZIPME010_EX_RATE[] TO GT_POPUP_EX_RATE[].
    TRY.
        CL_SALV_TABLE=>FACTORY(
          IMPORTING
            R_SALV_TABLE = GO_ALV
          CHANGING
            T_TABLE      = GT_POPUP_EX_RATE[] ).
      CATCH CX_SALV_MSG.
    ENDTRY.

    LR_FUNCTIONS = GO_ALV->GET_FUNCTIONS( ).
    LR_FUNCTIONS->SET_ALL( 'X' ).

    IF GO_ALV IS BOUND.
      LR_COLUMNS = GO_ALV->GET_COLUMNS( ).
      LR_COLUMNS->SET_OPTIMIZE( 'X' ).
      GO_ALV->SET_SCREEN_POPUP(
        START_COLUMN = 40
        END_COLUMN  = 140
        START_LINE  = 6
        END_LINE    = 15 ).
      GO_ALV->DISPLAY( ).
    ENDIF.
  ELSE.
    MESSAGE 'Standard exchange rate will be used' TYPE 'I'.
  ENDIF.
ENDFORM. " FRM_DISPLAY_POPUP_SCREEN

*&---------------------------------------------------------------------*
*&      Form  FRM_GET_EXCHANGE_RATE
*&---------------------------------------------------------------------*
FORM FRM_GET_EXCHANGE_RATE USING PS_PROJ      TYPE PROJ
                                 PS_PRPS      TYPE PRPS
                                 PS_LV2       TYPE PRPS-PSPNR
                                 PV_FROM_CURR TYPE BPIN-TWAER
                                 PV_TO_CURR   TYPE BPIN-TWAER
                                 PV_VALUE_IN  TYPE BP_WGL
                        CHANGING PV_VALUE_OUT TYPE BP_WGL
                                 PV_SUBRC     TYPE SY-SUBRC.
  DATA: LV_EX_RATE          TYPE KURSF,
        LS_ZIPME010_EX_RATE TYPE ZIPME010_EX_RATE.
*  BREAK X-TCHABANE.
  CALL FUNCTION 'ZIPME010_GET_EX_RATE'
    IMPORTING
      PS_EX_RATE = LS_ZIPME010_EX_RATE.
  IF SY-SUBRC = 0 AND LS_ZIPME010_EX_RATE IS NOT INITIAL.
    IF LS_ZIPME010_EX_RATE-WAERS <> PV_FROM_CURR.
      MESSAGE E012(ZIPM001) WITH  LS_ZIPME010_EX_RATE-WAERS PV_FROM_CURR. " Related PO currency &1 different than Transaction currency &2
      PV_SUBRC = 4.
    ELSEIF LS_ZIPME010_EX_RATE-KURSF IS NOT INITIAL.
      PV_VALUE_OUT  = PV_VALUE_IN * LS_ZIPME010_EX_RATE-KURSF.
    ENDIF. "  IF ls_zipme010_ex_rate-waers <> pv_from_curr.
  ELSE.
    SELECT *
      FROM ZIPME010_EX_RATE
      INTO TABLE @DATA(LT_ZIPME010_EX_RATE)
     WHERE  PROJ = @PS_PROJ-PSPNR
*        AND WBSE = @PS_PRPS-PSPNR
        AND WBSE_LV2 = @PS_LV2 "ADDED
        AND WAERS = @PV_FROM_CURR
       AND HWAER = @PV_TO_CURR
       AND ACTIVE = 'X'
     ORDER BY COUNTER DESCENDING.
    IF SY-SUBRC = 0.
      READ TABLE LT_ZIPME010_EX_RATE INTO LS_ZIPME010_EX_RATE
        INDEX 1.
      IF SY-SUBRC = 0.
        IF LS_ZIPME010_EX_RATE-WAERS <> PV_FROM_CURR.
          MESSAGE E012(ZIPM001) WITH  LS_ZIPME010_EX_RATE-WAERS PV_FROM_CURR. " Related PO currency &1 different than Transaction currency &2
          PV_SUBRC = 4.
        ELSEIF LS_ZIPME010_EX_RATE-KURSF IS NOT INITIAL.
          PV_VALUE_OUT  = PV_VALUE_IN * LS_ZIPME010_EX_RATE-KURSF.
        ENDIF. "  IF ls_zipme010_ex_rate-waers <> pv_from_curr.
      ENDIF.
    ENDIF.
  ENDIF.
ENDFORM. "  FRM_GET_EXCHANGE_RATE

*&---------------------------------------------------------------------*
*&      Form  FRM_CHECK_TCODE
*&---------------------------------------------------------------------*
FORM FRM_CHECK_TCODE USING PV_TCODE    TYPE SY-TCODE
                     CHANGING PV_INCLUDED TYPE CHAR01.
  RANGES: LR_TCODE FOR BPIN-TCODE.

  PV_INCLUDED = ABAP_FALSE.
  SELECT OPTI AS OPTION
         SIGN LOW HIGH
    FROM TVARVC
    INTO CORRESPONDING FIELDS OF TABLE LR_TCODE
   WHERE NAME = 'ZIPME010_TCODE'
     AND TYPE = 'S'.
  IF SY-SUBRC = 0 AND
    ( PV_TCODE IN LR_TCODE OR  SY-CPROG = GV_UPDATE_PROG OR   SY-CPROG = GV_TEST_PROG ).
    PV_INCLUDED = ABAP_TRUE.
  ENDIF. " IF sy-subrc = 0 AND
ENDFORM. " FRM_CHECK_TCODE


FORM FRM_CONVERT_AMOUNT  USING PV_TCODE TYPE SY-TCODE
                               PS_PROJ  TYPE PROJ
                               PS_BPIN  TYPE BPIN
                               PV_PLDAT LIKE BPIN-PLDAT
                               PV_KURST LIKE BPIN-KURST
                               PV_FROM_CURRENCY LIKE BPIN-TWAER
                               PV_TO_CURRENCY LIKE BPIN-TWAER
                               PV_ORGWAER LIKE BPIN-ORGWAER
                               PV_VALUE_IN LIKE BPGE-WLGES
                      CHANGING PV_VALUE_OUT LIKE BPGE-WLGES
                               PV_SUBRC LIKE SY-SUBRC.

  DATA: LR_EXC      TYPE REF TO CX_ROOT,
        LV_CONTINUE TYPE ABAP_BOOL,
        LS_PRPS     TYPE PRPS,
        LV_PARENT   TYPE PRPS-PSPNR.

  PERFORM FRM_CHECK_TCODE USING PV_TCODE
                     CHANGING LV_CONTINUE.

  CHECK LV_CONTINUE = ABAP_TRUE.

  SELECT *
    FROM PRPS
    INTO TABLE GT_PRPS
    WHERE PSPHI = PS_PROJ-PSPNR.
  IF SY-SUBRC <> 0.
    CLEAR GT_PRPS[].
  ENDIF. " IF sy-subrc <> 0.
  SELECT * FROM PRPS
    UP TO 1 ROWS
    INTO LS_PRPS
    WHERE OBJNR = PS_BPIN-OBJNR
    ORDER BY PRIMARY KEY.
  ENDSELECT.
  IF SY-SUBRC = 0.
    DATA LV_LEVEL TYPE PRPS-STUFE.
    LV_LEVEL  = LS_PRPS-STUFE.
    LV_PARENT = LS_PRPS-PSPNR.
    IF LS_PRPS-STUFE > 2.
      DO LS_PRPS-STUFE - 1 TIMES.
        IF LV_LEVEL = 2.
          CONTINUE.
        ELSE.
          PERFORM FRM_GET_LEVEL_2_WBSE  CHANGING LS_PRPS LV_PARENT LV_LEVEL.
        ENDIF.
      ENDDO.
    ENDIF. " IF ls_prps-stufe > 2.
  ELSE. " IF sy-subrc = 0.
    CLEAR LS_PRPS.
  ENDIF. " IF sy-subrc = 0.

  IF PV_FROM_CURRENCY = PV_TO_CURRENCY OR PV_VALUE_IN IS INITIAL.
    PV_VALUE_OUT = PV_VALUE_IN.
    CLEAR PV_SUBRC.
  ELSEIF PV_FROM_CURRENCY = PV_ORGWAER.
    PERFORM FRM_GET_EXCHANGE_RATE USING PS_PROJ
                                        LS_PRPS
                                        LV_PARENT
                                        PV_FROM_CURRENCY
                                        PV_TO_CURRENCY
                                        PV_VALUE_IN
                               CHANGING PV_VALUE_OUT
                                        PV_SUBRC.

  ELSEIF PV_TO_CURRENCY = PV_ORGWAER.
    PERFORM FRM_GET_EXCHANGE_RATE USING PS_PROJ
                                        LS_PRPS
                                        LV_PARENT
                                        PV_FROM_CURRENCY
                                        PV_TO_CURRENCY
                                        PV_VALUE_IN
                               CHANGING PV_VALUE_OUT
                                        PV_SUBRC.
  ELSE.
    PERFORM FRM_GET_EXCHANGE_RATE USING PS_PROJ
                                        LS_PRPS
                                        LV_PARENT
                                        PV_FROM_CURRENCY
                                        PV_ORGWAER
                                        PV_VALUE_IN
                               CHANGING PV_VALUE_OUT
                                        PV_SUBRC.

    IF PV_SUBRC IS INITIAL.
      PERFORM FRM_GET_EXCHANGE_RATE USING PS_PROJ
                                          LS_PRPS
                                          LV_PARENT
                                          PV_TO_CURRENCY
                                          PV_ORGWAER
                                          PV_VALUE_IN
                                 CHANGING PV_VALUE_OUT
                                          PV_SUBRC.
    ENDIF.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form FRM_GET_LEVEL_2_WBSE
*&---------------------------------------------------------------------*
FORM FRM_GET_LEVEL_2_WBSE CHANGING CS_PRPS   STRUCTURE PRPS
                                   CV_PARENT TYPE      PRPS-PSPNR
                                   CV_LEVEL  TYPE PRPS-STUFE.
  SELECT SINGLE UP, POSNR
    INTO @DATA(WA_PRHI)
    FROM PRHI
    WHERE POSNR = @CV_PARENT.
  IF SY-SUBRC = 0.
    CV_PARENT = WA_PRHI-UP.
    READ TABLE GT_PRPS INTO DATA(LS_PRPS) WITH KEY PSPNR = WA_PRHI-UP.
    IF SY-SUBRC = 0.
      CV_LEVEL = LS_PRPS-STUFE.
    ENDIF.
  ENDIF.
ENDFORM." FRM_GET_LEVEL_2_WBSE
*&---------------------------------------------------------------------*
*& Form FRM_UPDATE_DB_TABLE
*&---------------------------------------------------------------------*
FORM FRM_UPDATE_DB_TABLE USING    US_EKKO      TYPE EKKO
                                  UV_LOC_CURR  TYPE WAERS
                         CHANGING CT_WBSE_PROJ TYPE ANY TABLE
                                  CV_SUBRC     TYPE SY-SUBRC.
  DATA: LT_WBSE_PROJ TYPE TABLE OF TY_WBSE_PROJ WITH HEADER LINE,
        LV_PARENT    TYPE PRPS-PSPNR,
        LV_POSID_LV2 TYPE PRPS-POSID.
  CLEAR CV_SUBRC.
  LT_WBSE_PROJ[] =  CT_WBSE_PROJ[].
  IF LT_WBSE_PROJ[] IS NOT INITIAL.
*    BREAK X-TCHABANE.
    "...Select the entries saved in the table lt_zipme010_ex_rate
    CLEAR GS_ZIPME010_EX_RATE.
    CLEAR GT_ZIPME010_EX_RATE[].
    SELECT * FROM ZIPME010_EX_RATE  INTO TABLE @GT_ZIPME010_EX_RATE
      FOR ALL ENTRIES IN @LT_WBSE_PROJ
      WHERE PROJ = @LT_WBSE_PROJ-PROJ
        AND WBSE = @LT_WBSE_PROJ-WBSE
        AND WAERS = @US_EKKO-WAERS
        AND HWAER = @LT_WBSE_PROJ-WAERS "@UV_LOC_CURR "GV_LOC_CUR
        AND ACTIVE = 'X'.
    IF SY-SUBRC = 0.
      "...Change the entry in the table
      LOOP AT LT_WBSE_PROJ INTO DATA(LS_WBSE_PROJ).
        LV_PARENT = LS_WBSE_PROJ-WBSE.
        PERFORM FRM_GET_LV2 USING LS_WBSE_PROJ CHANGING LV_PARENT LV_POSID_LV2.
        READ TABLE GT_ZIPME010_EX_RATE INTO GS_ZIPME010_EX_RATE
        WITH KEY PROJ = LS_WBSE_PROJ-PROJ WBSE = LS_WBSE_PROJ-WBSE WAERS = US_EKKO-WAERS .
        IF SY-SUBRC = 0.
          "..Unactive previous entry
          MAC_MODIFY CV_SUBRC.
          "...Create a new entry in table
          CLEAR GV_COUNTER.
          GV_COUNTER = GS_ZIPME010_EX_RATE-COUNTER + 1 .
          MAC_INSERT  LS_WBSE_PROJ US_EKKO GV_COUNTER CV_SUBRC LV_PARENT LV_POSID_LV2.
        ELSE.
          "...Create a new entry in table
          CLEAR GV_COUNTER.
          GV_COUNTER = 1.
          MAC_INSERT  LS_WBSE_PROJ  US_EKKO  GV_COUNTER  CV_SUBRC  LV_PARENT LV_POSID_LV2.
        ENDIF.
      ENDLOOP.
    ELSE.
      "...Create a new entry in table
      CLEAR: GT_ZIPME010_EX_RATE[], LS_WBSE_PROJ.
      LOOP AT LT_WBSE_PROJ INTO LS_WBSE_PROJ.
        LV_PARENT = LS_WBSE_PROJ-WBSE.
        PERFORM FRM_GET_LV2 USING LS_WBSE_PROJ CHANGING LV_PARENT LV_POSID_LV2.
        "...Create a new entry in table
        CLEAR GV_COUNTER.
        GV_COUNTER = 1.
        MAC_INSERT  LS_WBSE_PROJ US_EKKO  GV_COUNTER CV_SUBRC  LV_PARENT LV_POSID_LV2.
      ENDLOOP.
    ENDIF.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form FRM_CHECK_PO_WBSE
*&---------------------------------------------------------------------*
FORM FRM_CHECK_PO_WBSE USING US_EKKO  TYPE EKKO
                             US_EKKO_OLD  TYPE EKKO
                             PT_EKKN  TYPE  MMPR_UEKKN
                             PT_EKKN_OLD  TYPE MMPR_UEKKN.
  DATA: UT_EKKN  TYPE TABLE OF EKKN WITH HEADER LINE.
  UT_EKKN[] = PT_EKKN[].

  DATA: LV_EBELN           TYPE EBELN,
        LV_EX_RATE_CURR    TYPE WAERS,
        LV_LOC_CURR        TYPE WAERS,
        LV_CURR_CHANGED    TYPE FLAG,
        LV_EX_RATE_CHANGED TYPE FLAG,
        LV_WBSE_CHANGED    TYPE FLAG,
        LT_WBSE_PROJ       TYPE TABLE OF TY_WBSE_PROJ WITH HEADER LINE,
        LV_SUBRC           TYPE SY-SUBRC VALUE 0.

  CLEAR: LV_CURR_CHANGED, LV_EX_RATE_CHANGED, LV_WBSE_CHANGED.
  BREAK X-TCHABANE.
  IF US_EKKO-WAERS <> US_EKKO_OLD-WAERS .
    LV_CURR_CHANGED = ABAP_TRUE.
  ENDIF.

  IF US_EKKO-WKURS <> US_EKKO_OLD-WKURS.
    LV_EX_RATE_CHANGED = ABAP_TRUE.
  ENDIF.

  IF PT_EKKN[] NE PT_EKKN_OLD[].
    LV_WBSE_CHANGED = ABAP_TRUE.
  ENDIF.

  CHECK  LV_CURR_CHANGED = ABAP_TRUE OR LV_EX_RATE_CHANGED = ABAP_TRUE OR LV_WBSE_CHANGED = ABAP_TRUE.
  "...Create a range lt_wbse_range for the WBSE assigned in the PO
  IF UT_EKKN[] IS NOT INITIAL. ""Entry exists and not changed or not created yet
    SELECT A~PSPNR AS WBSE , B~PSPNR AS PROJ , A~PWPOS AS WAERS
     INTO CORRESPONDING FIELDS OF TABLE @LT_WBSE_PROJ
     FROM PRPS AS A  INNER JOIN PROJ AS B
     ON B~PSPNR = A~PSPHI
     FOR ALL ENTRIES IN @UT_EKKN
     WHERE A~PSPNR = @UT_EKKN-PS_PSP_PNR.
    IF SY-SUBRC <> 0.
      CLEAR LT_WBSE_PROJ[].
    ENDIF.
    IF LT_WBSE_PROJ[] IS NOT INITIAL.
      PERFORM FRM_UPDATE_DB_TABLE
        USING
          US_EKKO
          GV_LOC_CURR "LV_LOC_CURR
        CHANGING
          LT_WBSE_PROJ[]
          LV_SUBRC.
      IF LV_SUBRC = 0.
      ENDIF.
    ENDIF. " IF lt_wbse_range[] IS NOT INITIAL.

    " READ ASSIGNED WBSE FROM DB
  ELSE. " if ut_ekkn is NOT INITIAL.

    SELECT EKKN~PS_PSP_PNR AS WBSE , PROJ~PSPNR AS PROJ , PRPS~PWPOS AS WAERS
      INTO CORRESPONDING FIELDS OF TABLE @LT_WBSE_PROJ
      FROM EKKN
      INNER JOIN PRPS ON EKKN~PS_PSP_PNR = PRPS~PSPNR
      INNER JOIN PROJ ON PROJ~PSPNR = PRPS~PSPHI
      WHERE EKKN~EBELN = @US_EKKO-EBELN.
    IF SY-SUBRC <> 0.
      CLEAR LT_WBSE_PROJ[].
    ENDIF.

    "...Read from DB
    IF LT_WBSE_PROJ[] IS NOT INITIAL.
      PERFORM FRM_UPDATE_DB_TABLE
         USING
           US_EKKO
           GV_LOC_CURR "LV_LOC_CURR
         CHANGING
           LT_WBSE_PROJ[]
           LV_SUBRC
       .
      IF LV_SUBRC <> 0 ."AND ( LV_CURR_CHANGED = ABAP_TRUE OR LV_EX_RATE_CHANGED = ABAP_TRUE ).

      ENDIF.
    ENDIF. " IF lt_wbse_range[] IS NOT INITIAL.
  ENDIF. " IF ut_ekkn IS NOT INITIAL.
ENDFORM. " FRM_CHECK_PO_WBSE
*&---------------------------------------------------------------------*
*& Form FRM_GET_LV2
*&---------------------------------------------------------------------*
FORM FRM_GET_LV2  USING    PS_WBSE_PROJ TYPE TY_WBSE_PROJ
                  CHANGING PV_PARENT TYPE PRPS-PSPNR
                           PV_POSID_LV2 TYPE PRPS-POSID .
  DATA: LS_PRPS     TYPE PRPS.
  SELECT *
    FROM PRPS
    INTO TABLE @DATA(LT_PRPS)
    WHERE PSPHI = @PS_WBSE_PROJ-PROJ.
  IF SY-SUBRC <> 0.
    CLEAR LT_PRPS[].
  ENDIF. " IF sy-subrc <> 0.
  SELECT * FROM PRPS
      UP TO 1 ROWS
      INTO LS_PRPS
      WHERE PSPNR = PS_WBSE_PROJ-WBSE
      ORDER BY PRIMARY KEY.
  ENDSELECT.
  IF SY-SUBRC = 0.
    DATA LV_LEVEL TYPE PRPS-STUFE.
    LV_LEVEL  = LS_PRPS-STUFE.
    PV_PARENT = LS_PRPS-PSPNR.
    IF LS_PRPS-STUFE > 2.
      DO LS_PRPS-STUFE - 1 TIMES.
        IF LV_LEVEL = 2.
          CONTINUE.
        ELSE.
          SELECT SINGLE UP, POSNR
             INTO @DATA(WA_PRHI)
             FROM PRHI
             WHERE POSNR = @PV_PARENT.
          IF SY-SUBRC = 0.
            PV_PARENT = WA_PRHI-UP.
            READ TABLE LT_PRPS INTO DATA(LW_PRPS) WITH KEY PSPNR = WA_PRHI-UP.
            IF SY-SUBRC = 0.
              LV_LEVEL = LW_PRPS-STUFE.
              PV_POSID_LV2 = LW_PRPS-POSID.
            ENDIF.
          ENDIF.
        ENDIF.
      ENDDO.
    ELSEIF LS_PRPS-STUFE = 2.
      PV_POSID_LV2 = LS_PRPS-POSID.
    ENDIF. " IF ls_prps-stufe > 2.
  ELSE. " IF sy-subrc = 0.
    CLEAR LS_PRPS.
  ENDIF. " IF sy-subrc = 0.
ENDFORM. " FRM_GET_LV2