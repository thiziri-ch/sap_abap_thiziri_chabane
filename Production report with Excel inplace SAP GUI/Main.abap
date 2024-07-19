*&--------------------------------------------------------------------------------------------*
*& Report Main
*&--------------------------------------------------------------------------------------------*
*&
*& Description: Auto-consumption calculation.
*&
*&--------------------------------------------------------------------------------------------*
*& DATE        Programmer          Mod. #           Description
*&--------------------------------------------------------------------------------------------*
*& 07.05.2024  Thiziri CHABANE            
*&--------------------------------------------------------------------------------------------*
REPORT Main.
INCLUDE Main_TOP.
INCLUDE Main_SCN.
INCLUDE Main_CLS.
INCLUDE Main_PBO.
INCLUDE Main_PAI.
INCLUDE Main_F01.
INCLUDE Main_F02.

INITIALIZATION.

  PERFORM FRM_INIT_PBO_VALUES.

AT SELECTION-SCREEN ON S_BUDAT.

  PERFORM FRM_REFRESH_DATE.

START-OF-SELECTION.

  PERFORM FRM_CHECK_PLANT_AUTH USING P_WERKS CHANGING GV_PLANT_AUTHORIZED.

  " Create excel document from OAOR template
  PERFORM FRM_CREATE_BASIC_OBJECTS.

  " Insert ranges dimension in excel
  PERFORM FRM_CREATE_RANGES.

  " Prepare report data
  PERFORM FRM_GET_DATA.

END-OF-SELECTION.
IF GV_CHECK_OK = ABAP_TRUE.
  CALL SCREEN 0100.
ENDIF.