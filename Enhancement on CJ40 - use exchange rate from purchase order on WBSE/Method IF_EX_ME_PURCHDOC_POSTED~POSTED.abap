  METHOD if_ex_me_purchdoc_posted~posted.
      PERFORM frm_check_po_wbse IN PROGRAM zipme010 IF FOUND USING  im_ekko im_ekko_old im_ekkn im_ekkn_old.
  ENDMETHOD.