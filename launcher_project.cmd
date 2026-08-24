@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion

title Audion DocFlow - project launcher

set "BASE_DIR=%~dp0"
if "%BASE_DIR:~-1%"=="" set "BASE_DIR=%BASE_DIR:~0,-1%"
cd /d "%BASE_DIR%"

set "CORE_DIR=%BASE_DIR%\system_core"
set "INSTALL_DIR=%BASE_DIR%\install"
set "RUNTIME_DIR=%BASE_DIR%\._runtime"
set "MENU_FILE=%RUNTIME_DIR%\project_menu_en.txt"
set "RES_FILE=%RUNTIME_DIR%\project_menu_en_res.txt"

if not exist "%RUNTIME_DIR%" mkdir "%RUNTIME_DIR%" >nul 2>nul
del /f /q "%MENU_FILE%" "%RES_FILE%" >nul 2>nul

call :RESOLVE_PYTHON
if errorlevel 1 goto NO_PYTHON

call :RESOLVE_FZF
if errorlevel 1 (
  set "MENU_MODE=CMD fallback"
) else (
  set "MENU_MODE=FZF"
)

:MAIN
cls
echo ======================================================================
echo   Audion DocFlow - project launcher
echo ======================================================================
echo Root:      %BASE_DIR%
echo Python:    %PYTHON_CMD% %PYTHON_ARGS%
echo Menu mode: %MENU_MODE%
echo.

if defined AUDION_AUTO_EXIT (
  echo [SMOKE] AUDION_AUTO_EXIT is set. Exiting before menu interaction.
  exit /b 0
)

if defined FZF_CMD goto FZF_MENU
goto FALLBACK_MENU

:FZF_MENU
> "%MENU_FILE%" echo === Audit and styles ===               ^| _section      ^| key DOCX processors
>>"%MENU_FILE%" echo [01] DOCX style processing             ^| docx_style     ^| headings captions TOC sections
>>"%MENU_FILE%" echo [02] DOCX style fix                    ^| docx_style_fix ^| conservatively assign styles
>>"%MENU_FILE%" echo [03] DOCX audit processor              ^| docx_audit     ^| sq m number sign percent degrees RF
>>"%MENU_FILE%" echo [04] DOCX text hygiene (scan)          ^| docx_hscan     ^| spaces punctuation checks
>>"%MENU_FILE%" echo [05] DOCX text hygiene (fix)           ^| docx_hfix      ^| safe fixes inside text nodes
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo === Find and replace ===               ^| _section      ^| high-level morphological automation
>>"%MENU_FILE%" echo [06] DOCX comma lowercase              ^| docx_comma     ^| lowercase capitals after commas
>>"%MENU_FILE%" echo [07] DOCX/XLSX morph replace           ^| morph_replace  ^| lemma search and replacement inflection
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo === Control and cleanup ===             ^| _section      ^| quality gate and finalization
>>"%MENU_FILE%" echo [08] DOCX quality check                ^| docx_gate      ^| colors highlights shading comments changes
>>"%MENU_FILE%" echo [09] Strict DOCX check                 ^| docx_hardgate  ^| nonzero exit if any FAIL
>>"%MENU_FILE%" echo [10] DOCX strip comments               ^| docx_nocomm    ^| remove comments and markers
>>"%MENU_FILE%" echo [11] DOCX accept changes (simple)      ^| docx_accept    ^| keep insertions remove deletions
>>"%MENU_FILE%" echo [12] DOCX finalize black + clean       ^| docx_black     ^| force black remove highlight shading strike
>>"%MENU_FILE%" echo [13] DOCX one-click finalize + gate    ^| docx_oneclick  ^| strip accept black strike hygiene gate
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo === Compare and reconcile ===           ^| _section      ^| DOCX diff, XLSX values, table reconcile
>>"%MENU_FILE%" echo [14] DOCX near-duplicate finder        ^| docx_dupes     ^| recursive content similarity
>>"%MENU_FILE%" echo [15] DOCX diff two documents           ^| docx_pairdiff  ^| text media structure no threshold
>>"%MENU_FILE%" echo [16] XLSX values diff (2 files)        ^| xlsx_diff      ^| compare computed cell values
>>"%MENU_FILE%" echo [17] DOCX/XLSX table reconcile         ^| reconcile      ^| match mismatch only A only B
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo === WORD/EXCEL TABLES ===              ^| _section      ^| DOCX stitching unifying orientation fitting
>>"%MENU_FILE%" echo [18] Stitch after PDF export           ^| table_stitch   ^| sliced DOCX tables
>>"%MENU_FILE%" echo [19] Unify tables only                 ^| table_unify_s  ^| DOCX columns and detected headers
>>"%MENU_FILE%" echo [20] Unify tables with heading         ^| table_unify_m  ^| DOCX section headings
>>"%MENU_FILE%" echo [21] Optimize table widths             ^| table_opt_w    ^| existing DOCX tables
>>"%MENU_FILE%" echo [22] Adapt table to orientation        ^| table_adapt_o  ^| A4/A3 orientation balance margins
>>"%MENU_FILE%" echo [23] Fit tables in document            ^| table_fit_m    ^| margins and A4/A3 setup
>>"%MENU_FILE%" echo [24] Extract tables from DOCX          ^| docx_tables    ^| XLSX plus index
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo === Technical operations ===            ^| _section      ^| media merge markdown cleanup
>>"%MENU_FILE%" echo [25] DOCX/PPTX extract media           ^| docx_media     ^| images without manual zip
>>"%MENU_FILE%" echo [26] DOCX merge via Word               ^| docx_merge     ^| Word COM inserts documents in order
>>"%MENU_FILE%" echo [27] Markdown tables to XLSX           ^| md2xlsx        ^| export pipe tables into Excel
>>"%MENU_FILE%" echo [28] DOCX/XLSX reduce cell margins     ^| table_margins  ^| compact table cells
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo [A] Open input folder                  ^| open_input     ^| explorer input
>>"%MENU_FILE%" echo [B] Open output folder                 ^| open_output    ^| explorer output
>>"%MENU_FILE%" echo [C] Open logs folder                   ^| open_logs      ^| explorer logs
>>"%MENU_FILE%" echo [G] Open GUI shell                     ^| gui            ^| NiceGUI/pywebview shell
>>"%MENU_FILE%" echo [T] Tools launcher                     ^| tools          ^| utility launcher
>>"%MENU_FILE%" echo [R] Russian project launcher           ^| launcher_ru    ^| switch to Russian shell
>>"%MENU_FILE%" echo [00] Exit                              ^| exit           ^| close

"%FZF_CMD%" --prompt="audion@office-kit [PROJECT-EN] > " --pointer=">" --header="Pick tool:" --layout=reverse --border="rounded" --info=hidden --margin=1,2 < "%MENU_FILE%" > "%RES_FILE%"

set "CHOICE="
set /p CHOICE=<"%RES_FILE%"
if not defined CHOICE goto MAIN

for /f "tokens=2 delims=|" %%a in ("%CHOICE%") do set "RAW=%%a"
call :TRIM RAW

if /I "%RAW%"=="docx_gate" goto DOCX_GATE
if /I "%RAW%"=="docx_hscan" goto DOCX_HSCAN
if /I "%RAW%"=="docx_hfix" goto DOCX_HFIX
if /I "%RAW%"=="docx_comma" goto DOCX_COMMA
if /I "%RAW%"=="morph_replace" goto MORPH_REPLACE
if /I "%RAW%"=="docx_nocomm" goto DOCX_NOCOMM
if /I "%RAW%"=="docx_accept" goto DOCX_ACCEPT
if /I "%RAW%"=="docx_black" goto DOCX_BLACK
if /I "%RAW%"=="docx_dupes" goto DOCX_DUPES
if /I "%RAW%"=="xlsx_diff" goto XLSX_DIFF
if /I "%RAW%"=="reconcile" goto RECONCILE
if /I "%RAW%"=="md2xlsx" goto MD2XLSX
if /I "%RAW%"=="docx_oneclick" goto DOCX_ONECLICK
if /I "%RAW%"=="docx_hardgate" goto DOCX_HARDGATE
if /I "%RAW%"=="docx_pairdiff" goto DOCX_PAIRDIFF
if /I "%RAW%"=="docx_tables" goto DOCX_TABLES
if /I "%RAW%"=="docx_media" goto DOCX_MEDIA
if /I "%RAW%"=="docx_merge" goto DOCX_MERGE
if /I "%RAW%"=="table_margins" goto TABLE_MARGINS
if /I "%RAW%"=="table_stitch" goto TABLE_STITCH
if /I "%RAW%"=="table_unify_s" goto TABLE_UNIFY_SAFE
if /I "%RAW%"=="table_unify_m" goto TABLE_UNIFY_MERGED
if /I "%RAW%"=="table_opt_w" goto TABLE_OPT_WIDTHS
if /I "%RAW%"=="table_adapt_o" goto TABLE_ADAPT_ORIENTATION
if /I "%RAW%"=="table_fit_m" goto TABLE_FIT_MARGINS
if /I "%RAW%"=="docx_style" goto DOCX_STYLE
if /I "%RAW%"=="docx_style_fix" goto DOCX_STYLE_FIX
if /I "%RAW%"=="docx_audit" goto DOCX_AUDIT
if /I "%RAW%"=="_section" goto MAIN
if /I "%RAW%"=="open_input" goto OPEN_INPUT
if /I "%RAW%"=="open_output" goto OPEN_OUTPUT
if /I "%RAW%"=="open_logs" goto OPEN_LOGS
if /I "%RAW%"=="gui" goto GUI
if /I "%RAW%"=="tools" goto TOOLS
if /I "%RAW%"=="launcher_ru" goto LAUNCHER_RU
if /I "%RAW%"=="exit" exit /b 0
goto MAIN

:FALLBACK_MENU
echo === Audit and styles ===
echo [1] DOCX style processing
echo [2] DOCX style fix
echo [3] DOCX audit processor
echo [4] DOCX text hygiene ^(scan^)
echo [5] DOCX text hygiene ^(fix^)
echo.
echo === Find and replace ===
echo [6] DOCX comma lowercase
echo [7] DOCX/XLSX morph replace
echo.
echo === Control and cleanup ===
echo [8] DOCX quality check
echo [9] Strict DOCX check
echo [A] DOCX strip comments
echo [B] DOCX accept changes ^(simple^)
echo [C] DOCX finalize black + clean
echo [D] DOCX one-click finalize + gate
echo.
echo === Compare and reconcile ===
echo [E] DOCX near-duplicate finder
echo [F] DOCX diff two documents
echo [G] XLSX values diff ^(2 files^)
echo [H] DOCX/XLSX table reconcile
echo.
echo === WORD/EXCEL TABLES ===
echo [I] Stitch after PDF export
echo [J] Unify tables only
echo [K] Unify tables with heading
echo [L] Optimize table widths
echo [M] Adapt table to orientation
echo [N] Fit tables in document
echo [O] Extract tables from DOCX
echo.
echo === Technical operations ===
echo [P] DOCX/PPTX extract media
echo [Q] DOCX merge via Word
echo [R] Markdown tables to XLSX
echo [S] DOCX/XLSX reduce cell margins
echo.
echo === Folders and service actions ===
echo [U] Open input folder
echo [V] Open output folder
echo [W] Open logs folder
echo [X] GUI shell
echo [Y] Tools launcher
echo [Z] Russian project launcher
echo [0] Exit
echo.
choice /C 123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ0 /N /M "Select: "
if errorlevel 36 exit /b 0
if errorlevel 35 goto LAUNCHER_RU
if errorlevel 34 goto TOOLS
if errorlevel 33 goto GUI
if errorlevel 32 goto OPEN_LOGS
if errorlevel 31 goto OPEN_OUTPUT
if errorlevel 30 goto OPEN_INPUT
if errorlevel 29 goto MAIN
if errorlevel 28 goto TABLE_MARGINS
if errorlevel 27 goto MD2XLSX
if errorlevel 26 goto DOCX_MERGE
if errorlevel 25 goto DOCX_MEDIA
if errorlevel 24 goto DOCX_TABLES
if errorlevel 23 goto TABLE_FIT_MARGINS
if errorlevel 22 goto TABLE_ADAPT_ORIENTATION
if errorlevel 21 goto TABLE_OPT_WIDTHS
if errorlevel 20 goto TABLE_UNIFY_MERGED
if errorlevel 19 goto TABLE_UNIFY_SAFE
if errorlevel 18 goto TABLE_STITCH
if errorlevel 17 goto RECONCILE
if errorlevel 16 goto XLSX_DIFF
if errorlevel 15 goto DOCX_PAIRDIFF
if errorlevel 14 goto DOCX_DUPES
if errorlevel 13 goto DOCX_ONECLICK
if errorlevel 12 goto DOCX_BLACK
if errorlevel 11 goto DOCX_ACCEPT
if errorlevel 10 goto DOCX_NOCOMM
if errorlevel 9 goto DOCX_HARDGATE
if errorlevel 8 goto DOCX_GATE
if errorlevel 7 goto MORPH_REPLACE
if errorlevel 6 goto DOCX_COMMA
if errorlevel 5 goto DOCX_HFIX
if errorlevel 4 goto DOCX_HSCAN
if errorlevel 3 goto DOCX_AUDIT
if errorlevel 2 goto DOCX_STYLE_FIX
if errorlevel 1 goto DOCX_STYLE
goto MAIN

:DOCX_COMMA
set "COMMA_DRY=Y"
set "KEEP_WORDS="
echo.
set /p KEEP_WORDS=Keep words (CSV or txt/json/yaml path, empty = none): 
choice /C YN /N /M "Dry-run without writing DOCX? (Y/N): "
if errorlevel 2 set "COMMA_DRY=N"
if errorlevel 1 set "COMMA_DRY=Y"

if defined KEEP_WORDS (
  if /I "%COMMA_DRY%"=="Y" (
    call :RUNPY "%CORE_DIR%\docx_comma_lowercase.py" --input "%BASE_DIR%\input" --output "%BASE_DIR%\output\comma_lowercase" --recursive --report "%BASE_DIR%\report\comma_lowercase.json" --keep-words "%KEEP_WORDS%" --dry-run
  ) else (
    call :RUNPY "%CORE_DIR%\docx_comma_lowercase.py" --input "%BASE_DIR%\input" --output "%BASE_DIR%\output\comma_lowercase" --recursive --report "%BASE_DIR%\report\comma_lowercase.json" --keep-words "%KEEP_WORDS%"
  )
) else (
  if /I "%COMMA_DRY%"=="Y" (
    call :RUNPY "%CORE_DIR%\docx_comma_lowercase.py" --input "%BASE_DIR%\input" --output "%BASE_DIR%\output\comma_lowercase" --recursive --report "%BASE_DIR%\report\comma_lowercase.json" --dry-run
  ) else (
    call :RUNPY "%CORE_DIR%\docx_comma_lowercase.py" --input "%BASE_DIR%\input" --output "%BASE_DIR%\output\comma_lowercase" --recursive --report "%BASE_DIR%\report\comma_lowercase.json"
  )
)
echo.
echo Output:
if /I "%COMMA_DRY%"=="N" echo   output\comma_lowercase
echo   report\comma_lowercase.json
if not defined AUDION_NO_PAUSE pause
goto MAIN

:MORPH_REPLACE
set "FIND_TEXT=гаражи"
set "REPLACE_TEXT=подземные гаражи"
set "EXTS=.docx,.xlsx"
set "MORPH_MODE=replace"
set "MORPH_OUTDIR=morph_replaced"
set "MORPH_INFLECT=Y"
set "MORPH_DRY=Y"
echo.
choice /C RA /N /M "Mode: [R] replace, [A] append next to match: "
if errorlevel 2 set "MORPH_MODE=append"
if errorlevel 1 set "MORPH_MODE=replace"
if /I "%MORPH_MODE%"=="append" set "MORPH_OUTDIR=morph_appended"
set /p FIND_TEXT=Find (default гаражи): 
if not defined FIND_TEXT set "FIND_TEXT=гаражи"
set /p REPLACE_TEXT=Replacement or added phrase (default подземные гаражи): 
if not defined REPLACE_TEXT set "REPLACE_TEXT=подземные гаражи"
set /p EXTS=Extensions (default .docx,.xlsx): 
if not defined EXTS set "EXTS=.docx,.xlsx"
choice /C YN /N /M "Inflect replacement to found grammar? (Y/N): "
if errorlevel 2 set "MORPH_INFLECT=N"
if errorlevel 1 set "MORPH_INFLECT=Y"
choice /C YN /N /M "Dry-run without writing files? (Y/N): "
if errorlevel 2 set "MORPH_DRY=N"
if errorlevel 1 set "MORPH_DRY=Y"

set "MORPH_EXTRA="
if /I "%MORPH_INFLECT%"=="Y" set "MORPH_EXTRA=%MORPH_EXTRA% --inflect-replacement"
if /I "%MORPH_DRY%"=="Y" set "MORPH_EXTRA=%MORPH_EXTRA% --dry-run"

call :RUNPY "%CORE_DIR%\morph_replace.py" --input "%BASE_DIR%\input" --output "%BASE_DIR%\output\%MORPH_OUTDIR%" --find "%FIND_TEXT%" --replace "%REPLACE_TEXT%" --mode "%MORPH_MODE%" --recursive --extensions "%EXTS%" --report "%BASE_DIR%\report\morph_replace.json"%MORPH_EXTRA%
echo.
echo Output:
if /I "%MORPH_DRY%"=="N" echo   output\%MORPH_OUTDIR%
echo   report\morph_replace.json
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_GATE
call :RUNPY "%CORE_DIR%\docx_quality_gate.py" --input "%BASE_DIR%\input" --out "%BASE_DIR%\report\docx_quality_gate_report.md"
echo.
echo Output:
echo   report\docx_quality_gate_report.docx
echo   report\docx_quality_gate_report.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_HSCAN
set "DOT=0"
echo.
choice /C YN /N /M "Check missing space after '.' too? (Y/N): "
if errorlevel 2 set "DOT=0"
if errorlevel 1 set "DOT=1"

if "%DOT%"=="1" (
  call :RUNPY "%CORE_DIR%\docx_text_hygiene_scan.py" --input "%BASE_DIR%\input" --out "%BASE_DIR%\report\docx_text_hygiene_report.md" --check-dot
) else (
  call :RUNPY "%CORE_DIR%\docx_text_hygiene_scan.py" --input "%BASE_DIR%\input" --out "%BASE_DIR%\report\docx_text_hygiene_report.md"
)
echo.
echo Output:
echo   report\docx_text_hygiene_report.docx
echo   report\docx_text_hygiene_report.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_HFIX
set "DOT=0"
echo.
choice /C YN /N /M "Fix missing space after '.' too? (Y/N): "
if errorlevel 2 set "DOT=0"
if errorlevel 1 set "DOT=1"

if "%DOT%"=="1" (
  call :RUNPY "%CORE_DIR%\docx_text_hygiene_fix.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\hygiene_fixed" --fix-dot
) else (
  call :RUNPY "%CORE_DIR%\docx_text_hygiene_fix.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\hygiene_fixed"
)
echo.
echo Output folder:
echo   output\hygiene_fixed
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_NOCOMM
call :RUNPY "%CORE_DIR%\docx_strip_comments.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\no_comments"
echo.
echo Output folder:
echo   output\no_comments
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_ACCEPT
call :RUNPY "%CORE_DIR%\docx_accept_changes_simple.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\accepted_changes"
echo.
echo Output folder:
echo   output\accepted_changes
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_BLACK
call :RUNPY "%CORE_DIR%\docx_finalize_black_clean.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\final_black"
echo.
echo Output folder:
echo   output\final_black
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_DUPES
echo.
set "THR=0.99"
set /p THR=Similarity threshold (default 0.99):
if not defined THR set "THR=0.99"

set "DIFFS=0"
choice /C YN /N /M "Write diffs for matched pairs? (Y/N): "
if errorlevel 2 set "DIFFS=0"
if errorlevel 1 set "DIFFS=1"

if "%DIFFS%"=="1" (
  call :RUNPY "%CORE_DIR%\docx_near_dupes.py" --input "%BASE_DIR%\input" --threshold %THR% --out "%BASE_DIR%\report\near_duplicates.md" --diff
) else (
  call :RUNPY "%CORE_DIR%\docx_near_dupes.py" --input "%BASE_DIR%\input" --threshold %THR% --out "%BASE_DIR%\report\near_duplicates.md"
)
echo.
echo Output:
echo   report\near_duplicates.docx
echo   report\near_duplicates.md
echo   report\near_duplicates.json
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_PAIRDIFF
echo.
set "A="
set "B="
set /p A=DOCX A (path relative to input\, subfolders allowed): 
if not defined A goto MAIN
set /p B=DOCX B (path relative to input\, subfolders allowed): 
if not defined B goto MAIN

call :RUNPY "%CORE_DIR%\docx_pair_diff.py" --a "%BASE_DIR%\input\%A%" --b "%BASE_DIR%\input\%B%" --out "%BASE_DIR%\report\docx_pair_diff.md"
echo.
echo Output files:
echo   report\docx_pair_diff.docx
echo   report\docx_pair_diff.md
echo   report\docx_pair_diff.diff
echo   report\docx_pair_diff.json
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_STITCH
call :RUNPY "%CORE_DIR%\docx_table_stitcher.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\stitched" --report-dir "%BASE_DIR%\report\word_excel_tables\stitched" --recursive
echo.
echo Output files:
echo   output\word_excel_tables\stitched
echo   report\word_excel_tables\stitched
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_STITCH_HEADER
call :RUNPY "%CORE_DIR%\docx_table_stitcher_reconstruct_running_header.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\stitched_running_header" --report-dir "%BASE_DIR%\report\word_excel_tables\stitched_running_header" --recursive
echo.
echo Output files:
echo   output\word_excel_tables\stitched_running_header
echo   report\word_excel_tables\stitched_running_header
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_UNIFY_SAFE
call :RUNPY "%CORE_DIR%\docx_table_unifier.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\unified_safe" --report-dir "%BASE_DIR%\report\word_excel_tables\unified_safe" --all --recursive --mode safe --layout standard
echo.
echo Output files:
echo   output\word_excel_tables\unified_safe
echo   report\word_excel_tables\unified_safe
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_UNIFY_WIDTH
call :RUNPY "%CORE_DIR%\docx_table_unifier.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\unified_width_only" --report-dir "%BASE_DIR%\report\word_excel_tables\unified_width_only" --all --recursive --mode width-only --layout standard
echo.
echo Output files:
echo   output\word_excel_tables\unified_width_only
echo   report\word_excel_tables\unified_width_only
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_UNIFY_MERGED
call :RUNPY "%CORE_DIR%\docx_table_unifier.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\unified_merged_sections" --report-dir "%BASE_DIR%\report\word_excel_tables\unified_merged_sections" --all --recursive --mode safe --layout merged-sections
echo.
echo Output files:
echo   output\word_excel_tables\unified_merged_sections
echo   report\word_excel_tables\unified_merged_sections
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_OPT_WIDTHS
call :RUNPY "%CORE_DIR%\docx_table_width_optimizer.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\optimized_widths" --report-dir "%BASE_DIR%\report\word_excel_tables\optimized_widths" --all --recursive --mode preserve-width
echo.
echo Output files:
echo   output\word_excel_tables\optimized_widths
echo   report\word_excel_tables\optimized_widths
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_ADAPT_ORIENTATION
call :RUNPY "%CORE_DIR%\docx_table_width_optimizer.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\adapted_orientation" --report-dir "%BASE_DIR%\report\word_excel_tables\adapted_orientation" --all --recursive --mode fit-to-margins --fit-target page-setup --page-setup-margin-fallback keep --page-size A4 --page-orientation landscape
echo.
echo Output files:
echo   output\word_excel_tables\adapted_orientation
echo   report\word_excel_tables\adapted_orientation
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_FIT_MARGINS
call :RUNPY "%CORE_DIR%\docx_table_width_optimizer.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\fit_to_margins" --report-dir "%BASE_DIR%\report\word_excel_tables\fit_to_margins" --all --recursive --mode fit-to-margins
echo.
echo Output files:
echo   output\word_excel_tables\fit_to_margins
echo   report\word_excel_tables\fit_to_margins
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_TABLES
call :RUNPY "%CORE_DIR%\docx_extract_tables.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\tables" --out "%BASE_DIR%\report\docx_tables_index.md"
echo.
echo Output files:
echo   output\tables
echo   report\docx_tables_index.docx
echo   report\docx_tables_index.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_MEDIA
call :RUNPY "%CORE_DIR%\docx_extract_media.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\media" --out "%BASE_DIR%\report\office_media_index.md"
echo.
echo Output files:
echo   output\media
echo   report\office_media_index.docx
echo   report\office_media_index.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_MARGINS
call :RUNPY "%CORE_DIR%\docx_xlsx_table_cell_margins.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\table_cell_margins" --report "%BASE_DIR%\report\table_cell_margins.md" --json-out "%BASE_DIR%\report\table_cell_margins.json" --margin-cm 0.1
echo.
echo Output files:
echo   output\table_cell_margins
echo   report\table_cell_margins.docx
echo   report\table_cell_margins.md
echo   report\table_cell_margins.json
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_MERGE
call :RUNPY "%CORE_DIR%\docx_merge.py" --input "%BASE_DIR%\input" --out "%BASE_DIR%\output\merged.docx" --report "%BASE_DIR%\report\docx_merge.md"
echo.
echo Output files:
echo   output\merged.docx
echo   report\docx_merge.docx
echo   report\docx_merge.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_STYLE
call :RUNPY "%CORE_DIR%\docx_style_processor.py" --input "%BASE_DIR%\input" --out "%BASE_DIR%\report\docx_style_processing.md"
echo.
echo Output files:
echo   report\docx_style_processing.docx
echo   report\docx_style_processing.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_STYLE_FIX
call :RUNPY "%CORE_DIR%\docx_style_processor.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\style_processed" --out "%BASE_DIR%\report\docx_style_processing.md" --fix
echo.
echo Output files:
echo   output\style_processed
echo   report\docx_style_processing.docx
echo   report\docx_style_processing.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_AUDIT
set "AUDIT_FIX=0"
echo.
choice /C YN /N /M "Apply safe audit fixes? (Y/N): "
if errorlevel 2 set "AUDIT_FIX=0"
if errorlevel 1 set "AUDIT_FIX=1"

if "%AUDIT_FIX%"=="1" (
  call :RUNPY "%CORE_DIR%\docx_audit_processor.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\audit_processed" --report "%BASE_DIR%\report\docx_audit_processor.md" --json-out "%BASE_DIR%\report\docx_audit_processor.json" --fix
) else (
  call :RUNPY "%CORE_DIR%\docx_audit_processor.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\audit_processed" --report "%BASE_DIR%\report\docx_audit_processor.md" --json-out "%BASE_DIR%\report\docx_audit_processor.json"
)
echo.
echo Output files:
if "%AUDIT_FIX%"=="1" echo   output\audit_processed
echo   report\docx_audit_processor.docx
echo   report\docx_audit_processor.md
echo   report\docx_audit_processor.json
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_ONECLICK
echo.
set "OPT_COMM=Y"
set "OPT_CHG=Y"
set "OPT_DOT=N"

choice /C YN /N /M "Strip comments? (Y/N) [default Y]: "
if errorlevel 2 set "OPT_COMM=N"
if errorlevel 1 set "OPT_COMM=Y"

choice /C YN /N /M "Accept tracked changes (simple)? (Y/N) [default Y]: "
if errorlevel 2 set "OPT_CHG=N"
if errorlevel 1 set "OPT_CHG=Y"

choice /C YN /N /M "Fix missing space after '.' too? (Y/N) [default N]: "
if errorlevel 2 set "OPT_DOT=N"
if errorlevel 1 set "OPT_DOT=Y"

set "EXTRA="
if /I "%OPT_COMM%"=="N" set "EXTRA=%EXTRA% --keep-comments"
if /I "%OPT_CHG%"=="N" set "EXTRA=%EXTRA% --keep-changes"
if /I "%OPT_DOT%"=="Y" set "EXTRA=%EXTRA% --fix-dot"

call :RUNPY "%CORE_DIR%\docx_oneclick_finalize_gate.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\final" --summary "%BASE_DIR%\report\summary_pass_fail.md"%EXTRA%
echo.
echo Output folder:
echo   output\final
echo Reports:
echo   report\summary_pass_fail.docx
echo   report\summary_pass_fail.md
echo   report\gate_before.docx
echo   report\gate_before.md
echo   report\gate_after.docx
echo   report\gate_after.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_HARDGATE
call :RUNPY "%CORE_DIR%\docx_quality_gate_hard_fail.py" --input "%BASE_DIR%\input" --out "%BASE_DIR%\report\docx_gate_hard_fail.md"
echo.
if errorlevel 1 (
  echo RESULT: FAIL
) else (
  echo RESULT: PASS
)
echo Report:
echo   report\docx_gate_hard_fail.docx
echo   report\docx_gate_hard_fail.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:XLSX_DIFF
echo.
set "A="
set "B="
set /p A=File A (relative to input\, subfolders allowed): 
if not defined A goto MAIN
set /p B=File B (relative to input\, subfolders allowed): 
if not defined B goto MAIN

call :RUNPY "%CORE_DIR%\xlsx_values_diff.py" --a "%BASE_DIR%\input\%A%" --b "%BASE_DIR%\input\%B%" --out "%BASE_DIR%\report\xlsx_values_diff_report.md"
echo.
echo Output:
echo   report\xlsx_values_diff_report.docx
echo   report\xlsx_values_diff_report.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:RECONCILE
echo.
set "A="
set "B="
set "KEYS="
set "FIELDS="
set /p A=File A (xlsx/csv/docx, relative to input\): 
if not defined A goto MAIN
set /p B=File B (xlsx/csv/docx, relative to input\): 
if not defined B goto MAIN
set /p KEYS=Key columns (comma-separated):
if not defined KEYS goto MAIN
set /p FIELDS=Fields to compare (comma-separated, empty = auto):

if defined FIELDS (
  call :RUNPY "%CORE_DIR%\tabular_reconcile_4lists.py" --a "%BASE_DIR%\input\%A%" --b "%BASE_DIR%\input\%B%" --key-csv "%KEYS%" --fields-csv "%FIELDS%" --out "%BASE_DIR%\report\reconcile_4lists.md"
) else (
  call :RUNPY "%CORE_DIR%\tabular_reconcile_4lists.py" --a "%BASE_DIR%\input\%A%" --b "%BASE_DIR%\input\%B%" --key-csv "%KEYS%" --out "%BASE_DIR%\report\reconcile_4lists.md"
)

echo.
echo Output:
echo   report\reconcile_4lists.docx
echo   report\reconcile_4lists.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:MD2XLSX
echo.
set "MD="
set "OUTX="
set /p MD=Markdown file (in input\):
if not defined MD goto MAIN
set /p OUTX=Output xlsx name (e.g. report.xlsx):
if not defined OUTX set "OUTX=report.xlsx"

call :RUNPY "%CORE_DIR%\md_tables_to_xlsx.py" --md "%BASE_DIR%\input\%MD%" --out "%BASE_DIR%\output\%OUTX%"
echo.
echo Output:
echo   output\%OUTX%
if not defined AUDION_NO_PAUSE pause
goto MAIN

:OPEN_INPUT
start "" explorer "%BASE_DIR%\input"
goto MAIN

:OPEN_OUTPUT
start "" explorer "%BASE_DIR%\output"
goto MAIN

:OPEN_LOGS
start "" explorer "%BASE_DIR%\logs"
goto MAIN

:GUI
if exist "%BASE_DIR%\launcher_gui.cmd" (
  call "%BASE_DIR%\launcher_gui.cmd"
) else (
  echo.
  echo [WARN] launcher_gui.cmd was not found.
  if not defined AUDION_NO_PAUSE pause
)
goto MAIN

:TOOLS
call "%BASE_DIR%\launcher_tools.cmd"
goto MAIN

:LAUNCHER_RU
if exist "%BASE_DIR%\launcher_project_ru.cmd" (
  call "%BASE_DIR%\launcher_project_ru.cmd"
) else (
  echo.
  echo [WARN] launcher_project_ru.cmd was not found.
  if not defined AUDION_NO_PAUSE pause
)
goto MAIN

:NO_PYTHON
cls
echo [ERROR] Python runtime was not resolved.
echo.
echo Supported locations:
echo   runtime\python.exe
echo   runtime\python\python.exe
echo   py -3.12
echo   python
echo.
echo Use builder_main.cmd or install\Build_Portable_Env_Build.cmd
if not defined AUDION_NO_PAUSE pause
exit /b 1

:RUNPY
set "TARGET=%~1"
shift
if not exist "%TARGET%" (
  echo [ERROR] Script not found:
  echo %TARGET%
  goto :eof
)
"%PYTHON_CMD%" %PYTHON_ARGS% -c "import runpy, sys; from pathlib import Path; target = Path(sys.argv[1]).resolve(); sys.path.insert(0, str(target.parent)); sys.argv = [str(target), *sys.argv[2:]]; runpy.run_path(str(target), run_name='__main__')" "%TARGET%" %*
goto :eof

:RESOLVE_PYTHON
set "PYTHON_CMD="
set "PYTHON_ARGS="

if exist "%BASE_DIR%\runtime\python.exe" (
  set "PYTHON_CMD=%BASE_DIR%\runtime\python.exe"
  goto PY_OK
)

if exist "%BASE_DIR%\runtime\python\python.exe" (
  set "PYTHON_CMD=%BASE_DIR%\runtime\python\python.exe"
  goto PY_OK
)

py -3.12 -V >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_CMD=py"
  set "PYTHON_ARGS=-3.12"
  goto PY_OK
)

where python >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_CMD=python"
  goto PY_OK
)

exit /b 1

:PY_OK
exit /b 0

:RESOLVE_FZF
set "FZF_CMD="
if defined AUDION_DISABLE_FZF exit /b 1
if exist "%CORE_DIR%\fzf.exe" (
  set "FZF_CMD=%CORE_DIR%\fzf.exe"
  exit /b 0
)
where fzf >nul 2>nul
if not errorlevel 1 (
  set "FZF_CMD=fzf"
  exit /b 0
)
exit /b 1

:TRIM
for /f "tokens=* delims= " %%z in ("!%~1!") do set "%~1=%%z"
:TRIM_R
if "!%~1:~-1!"==" " set "%~1=!%~1:~0,-1!" & goto TRIM_R
goto :eof
