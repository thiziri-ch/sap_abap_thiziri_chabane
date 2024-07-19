*&---------------------------------------------------------------------*
*&      Module  USER_COMMAND_0100  INPUT
*&---------------------------------------------------------------------*
MODULE USER_COMMAND_0100 INPUT.
  CASE GV_OK_CODE.
      "Post quantity Flared Gas
    WHEN 'SAVE_FG'.
      PERFORM FRM_SAVE_FG USING GV_DATE_FG
                                GV_QTY_FG
                                GV_COST_CENTER.
      "Post quantity AC
    WHEN 'SAVE_AC'.
      PERFORM FRM_SAVE_AC_NEW USING GV_DATE_AC.

    WHEN 'BACK' OR 'QUIT'.

      PERFORM LEAVE_SCREEN.
    WHEN 'EXIT'.
      PERFORM LEAVE_PROGRAM.
    WHEN 'EXPORT'.
      PERFORM FRM_EXPORT_EXCEL_FILE.
    WHEN OTHERS.
  ENDCASE.
ENDMODULE.