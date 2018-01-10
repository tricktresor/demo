REPORT ZTRCKTRSR_DEMO_EXCEL.


TYPE-POOLS ole2.

DATA: excel         TYPE ole2_object,
      mapl          TYPE ole2_object,         " list of workbooks
      workbook      TYPE ole2_object,         " Workbook
      map           TYPE ole2_object.         " Mappe


CREATE OBJECT excel 'EXCEL.APPLICATION'.
IF sy-subrc NE 0.
  WRITE : / 'Fehler CREATE OBJECT'.
ELSE.
  SET PROPERTY OF excel   'Visible'   = 0.  "nicht sichtbar
  CALL METHOD OF excel    'Workbooks' = workbook.
  CALL METHOD OF workbook 'Add'       = map.

  PERFORM fill_cell USING 1 1 1 3 'TRICKTRESOR'.
  PERFORM fill_cell USING 2 1 1 1 'http://www.tricktresor.de'.
  PERFORM fill_cell USING 5 1 1 3 'Datum'.
  PERFORM fill_cell USING 5 2 0 5 sy-datum.

  PERFORM fill_cell USING 6 1 1 3 'Uhrzeit'.
  PERFORM fill_cell USING 6 2 0 5 sy-uzeit.

  CALL METHOD OF map 'SaveAs' EXPORTING #1 = 'd:\temp\test1.xls'.
  CALL METHOD OF workbook 'CLOSE'.
  CALL METHOD OF excel 'QUIT'.
  FREE OBJECT workbook.
  FREE OBJECT excel.

ENDIF.


*---------------------------------------------------------------------*
*       FORM FILL_CELL                                                *
*---------------------------------------------------------------------*
*  –>  I      Zeile                                                   *
*  –>  J      Spalte                                                  *
*  –>  BOLD   Fett=1, Normal=0                                        *
*  –>  COL    Farbe:                                                  *
*              1=Schwarz, 2=weiss, 3=rot, 4=grün, 5=blau, 6=gelb      *
*  –>  VAL    Wert                                                    *
*---------------------------------------------------------------------*
FORM fill_cell USING i j bold col val.

  DATA:
      h_zl TYPE ole2_object,           " cell
      h_f TYPE ole2_object.            " font

  CALL METHOD OF excel 'Cells'     = h_zl
       EXPORTING #1 = i #2 = j.
  SET PROPERTY OF h_zl 'Value'     = val .
  GET PROPERTY OF h_zl 'Font'      = h_f.
  SET PROPERTY OF h_f 'Bold'       = bold .
  SET PROPERTY OF h_f 'ColorIndex' = col. "Rot
  SET PROPERTY OF h_f 'Size'       = 16.

ENDFORM.