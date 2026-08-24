@echo off
setlocal
if not defined RUNTIME_NODE (
  1>&2 echo unzip: RUNTIME_NODE is not set. Load the bundled workspace dependencies first.
  exit /b 2
)
"%RUNTIME_NODE%" "%~dp0unzip_compat.mjs" %*
set "_make_pptx_editable_unzip_rc=%ERRORLEVEL%"
endlocal & exit /b %_make_pptx_editable_unzip_rc%

