*&---------------------------------------------------------------------*
*& Include          ZIPME010_F01
*&---------------------------------------------------------------------*
DEFINE MAC_SET_COL_NAME.
  TRY.
      lr_column ?= lr_columns->get_column( columnname = &1 ).
      lr_column->set_short_text( value = &2 ).
      lr_column->set_medium_text( value = &3 ).
      lr_column->set_long_text( value = &4 ).
    CATCH cx_salv_not_found.
  ENDTRY.
END-OF-DEFINITION.

DEFINE MAC_SET_COL_TECH.
  TRY.
  lr_column ?= lr_columns->get_column( columnname = &1 ).
  lr_column->set_technical( &2 ).
  CATCH cx_salv_not_found.
  ENDTRY.
END-OF-DEFINITION.
DEFINE MAC_INSERT.
  "...Create a new entry in table
  gs_zipme010_ex_rate-mandt = sy-mandt.
  gs_zipme010_ex_rate-proj = &1-proj.
  gs_zipme010_ex_rate-WBSE_LV2 = &5.
  gs_zipme010_ex_rate-POSID_LV2 = &6.
  gs_zipme010_ex_rate-wbse = &1-wbse.
  gs_zipme010_ex_rate-waers = &2-waers.
  gs_zipme010_ex_rate-hwaer = &1-waers. "&3.
  gs_zipme010_ex_rate-counter = &3 .
  gs_zipme010_ex_rate-ebeln = &2-ebeln.
  gs_zipme010_ex_rate-kursf = &2-wkurs.
  gs_zipme010_ex_rate-active = 'X'.
  gs_zipme010_ex_rate-processed = ''.
  gs_zipme010_ex_rate-ernam = sy-uname.
  gs_zipme010_ex_rate-andat = sy-datum.

  CLEAR: gs_zipme010_ex_rate-aenam,
         gs_zipme010_ex_rate-aedat.
  INSERT zipme010_ex_rate FROM gs_zipme010_ex_rate.
  IF sy-subrc <> 0.
  &4 = 4.
  ENDIF.
END-OF-DEFINITION.

DEFINE MAC_MODIFY.
  "...Modify entry in table
  gs_zipme010_ex_rate-aenam = sy-uname.
  gs_zipme010_ex_rate-aedat = sy-datum.
  gs_zipme010_ex_rate-active = ' '.
  MODIFY zipme010_ex_rate FROM gs_zipme010_ex_rate.
  IF sy-subrc <> 0.
  &1 = 4.
  ENDIF.
END-OF-DEFINITION.