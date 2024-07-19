*&---------------------------------------------------------------------*
*& Module STATUS_0100 OUTPUT
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
MODULE STATUS_0100 OUTPUT.

  " Set screen gui status and title
  PERFORM FRM_SET_STATUS.

  " Handle buttons hide/unhide
  PERFORM FRM_ACTIVE_BUTTON.

  " Display the container and the excel document
  PERFORM FRM_DISPLAY_DATA.

ENDMODULE.