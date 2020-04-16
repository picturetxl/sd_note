*&---------------------------------------------------------------------*
*& Report ZTL_TEST
*&---------------------------------------------------------------------*
*& 销售订单录入模板的导入
*&---------------------------------------------------------------------*
REPORT ZTL_TEST.

TYPE-POOLS icon.
TABLES sscrfields. 

DATA functxt        TYPE smp_dyntxt.
DATA L_FILENAME     TYPE STRING.
DATA L_FULLPATH     TYPE STRING.
DATA L_ACTION       TYPE I.
DATA P_PATH         TYPE LOCALFILE VALUE'C:\'.
DATA G_EXCEL_ID     TYPE W3OBJID VALUE 'ZSDXSDDMB'.
DATA G_SHEET_NAME   TYPE CHAR20 VALUE '凭证抬头'.
DATA I_DATA         LIKE ZCRS_EXCEL OCCURS 0 WITH HEADER LINE.

PARAMETERS: p_carrid TYPE s_carr_id,
            p_cityfr TYPE s_from_cit.

SELECTION-SCREEN: FUNCTION KEY 1,
                  FUNCTION KEY 2.

INITIALIZATION.
  functxt-icon_id   = ICON_EXPORT."指定按钮的功能:输出
  functxt-icon_text = '销售订单导入模板'.
  sscrfields-functxt_01 = functxt.


AT SELECTION-SCREEN.
  CASE sscrfields-ucomm.
    WHEN 'FC01'.
       PERFORM SEARCH_HELP_PATH."下载 模板
       PERFORM DOWNLOAD_EXCEL.
  ENDCASE.


FORM SEARCH_HELP_PATH .
  DATA: L_PATH TYPE STRING .

  L_FILENAME = '凭证模板'.
  CALL METHOD CL_GUI_FRONTEND_SERVICES=>FILE_SAVE_DIALOG
    EXPORTING
      WINDOW_TITLE         = '请选择保存路径'
      DEFAULT_FILE_NAME    = L_FILENAME
      DEFAULT_EXTENSION    = 'XLS'
      INITIAL_DIRECTORY    = 'C:\'
    CHANGING
      FILENAME             = L_FILENAME
      PATH                 = L_PATH
      FULLPATH             = L_FULLPATH
      USER_ACTION          = L_ACTION
    EXCEPTIONS
      CNTL_ERROR           = 1
      ERROR_NO_GUI         = 2
      NOT_SUPPORTED_BY_GUI = 3
      OTHERS               = 4.

  IF SY-SUBRC <> 0.
    MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
               WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.
  P_PATH = L_PATH.

ENDFORM.

FORM DOWNLOAD_EXCEL .
  DATA: L_EXCEL_FNAME TYPE CHAR20,
        L_EMSG        TYPE CHAR100.
  REPLACE ALL OCCURRENCES OF  '.XLS' IN L_FILENAME WITH ' '.
  CONDENSE L_FILENAME NO-GAPS.
  L_EXCEL_FNAME = L_FILENAME.

  IF  L_ACTION = 0.
    CALL FUNCTION 'ZDOWNLOAD_TO_TEMP'
      EXPORTING
        SHEET_NAME  = G_SHEET_NAME  "Excel sheet名称
        EXCEL_ID    = G_EXCEL_ID    "Excel ID in SAP
        PATH        = P_PATH        "Path
        EXCEL_FNAME = L_EXCEL_FNAME "Excel file name
        SHOW_EXCEL  = 'X'           "是否显示EXCEL
      IMPORTING
        E_MESSAGE   = L_EMSG        "返回错误日志
      TABLES
        I_DATA      = I_DATA        "数据内容
      EXCEPTIONS
        EXCEL_ERROR = 1
        OTHERS      = 2.
    IF SY-SUBRC <> 0.  ENDIF.
  ENDIF.
ENDFORM.