@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion

title Audion DocFlow - русский лаунчер

set "BASE_DIR=%~dp0"
if "%BASE_DIR:~-1%"=="" set "BASE_DIR=%BASE_DIR:~0,-1%"
cd /d "%BASE_DIR%"

set "CORE_DIR=%BASE_DIR%\system_core"
set "INSTALL_DIR=%BASE_DIR%\install"
set "RUNTIME_DIR=%BASE_DIR%\._runtime"
set "MENU_FILE=%RUNTIME_DIR%\project_menu_ru.txt"
set "RES_FILE=%RUNTIME_DIR%\project_menu_ru_res.txt"

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
echo   Audion DocFlow - русский лаунчер
echo ======================================================================
echo Корень:       %BASE_DIR%
echo Python:       %PYTHON_CMD% %PYTHON_ARGS%
echo Режим меню:   %MENU_MODE%
echo.

if defined AUDION_AUTO_EXIT (
  echo [SMOKE] Установлен AUDION_AUTO_EXIT. Выход до интерактивного меню.
  exit /b 0
)

if defined FZF_CMD goto FZF_MENU
goto FALLBACK_MENU

:FZF_MENU
> "%MENU_FILE%" echo === АУДИТ И СТИЛИ ===                 ^| _section      ^| ключевые DOCX-процедуры
>>"%MENU_FILE%" echo [01] DOCX стилевой процессинг          ^| docx_style     ^| заголовки подписи оглавление секции
>>"%MENU_FILE%" echo [02] DOCX стилевой fix                 ^| docx_style_fix ^| осторожно назначить стили
>>"%MENU_FILE%" echo [03] DOCX audit processor              ^| docx_audit     ^| кв м номера проценты градусы РФ
>>"%MENU_FILE%" echo [04] DOCX проверка текста              ^| docx_hscan     ^| пробелы и пунктуация
>>"%MENU_FILE%" echo [05] DOCX исправить текст              ^| docx_hfix      ^| не менять структуру документа
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo === Найти и заменить ===              ^| _section      ^| высший уровень морфологической автоматизации
>>"%MENU_FILE%" echo [06] DOCX регистр после запятых        ^| docx_comma     ^| понизить заглавные после запятых
>>"%MENU_FILE%" echo [07] DOCX/XLSX морфологическая замена  ^| morph_replace  ^| поиск по леммам и согласование падежа
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo === Контроль и очистка ===             ^| _section      ^| quality gate и финализация
>>"%MENU_FILE%" echo [08] Проверка качества DOCX            ^| docx_gate      ^| цвет шрифта подсветка комментарии правки
>>"%MENU_FILE%" echo [09] Строгая проверка DOCX             ^| docx_hardgate  ^| ненулевой код при любой ошибке
>>"%MENU_FILE%" echo [10] DOCX удалить комментарии          ^| docx_nocomm    ^| убрать комментарии и маркеры
>>"%MENU_FILE%" echo [11] DOCX принять правки (простой)     ^| docx_accept    ^| сохранить вставки убрать удаления
>>"%MENU_FILE%" echo [12] DOCX чёрный текст + очистка       ^| docx_black     ^| чёрный текст убрать подсветку заливку зачёркивание
>>"%MENU_FILE%" echo [13] DOCX финализация + проверка       ^| docx_oneclick  ^| комментарии правки цвет зачёркивание гигиена
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo === Сравнение и сверка ===             ^| _section      ^| DOCX diff XLSX значения табличная сверка
>>"%MENU_FILE%" echo [14] DOCX поиск почти дубликатов       ^| docx_dupes     ^| рекурсивное сравнение содержимого
>>"%MENU_FILE%" echo [15] DOCX diff двух документов         ^| docx_pairdiff  ^| текст media структура без порога
>>"%MENU_FILE%" echo [16] XLSX сравнение значений (2 файла) ^| xlsx_diff      ^| сравнить вычисленные значения ячеек
>>"%MENU_FILE%" echo [17] Сверка таблиц DOCX/XLSX           ^| reconcile      ^| совпадения отличия только A только B
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo === ТАБЛИЦЫ WORD/EXCEL ===             ^| _section      ^| DOCX сшивание объединение ориентация вписывание
>>"%MENU_FILE%" echo [18] Сшить после экспорта PDF          ^| table_stitch   ^| разрезанные DOCX-таблицы
>>"%MENU_FILE%" echo [19] Объединить только таблицы         ^| table_unify_s  ^| DOCX колонки и распознанная шапка
>>"%MENU_FILE%" echo [20] Объединить таблицы с заголовком   ^| table_unify_m  ^| DOCX заголовки и разделы
>>"%MENU_FILE%" echo [21] Оптимизировать ширину таблиц      ^| table_opt_w    ^| существующие DOCX-таблицы
>>"%MENU_FILE%" echo [22] Адаптировать таблицу к ориентации ^| table_adapt_o  ^| A4/A3 ориентация балансировка поля
>>"%MENU_FILE%" echo [23] Вписать таблицы в документе       ^| table_fit_m    ^| поля документа и A4/A3
>>"%MENU_FILE%" echo [24] Извлечь таблицы из DOCX           ^| docx_tables    ^| XLSX плюс индекс
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo === Технические операции ===           ^| _section      ^| media merge markdown очистка
>>"%MENU_FILE%" echo [25] DOCX/PPTX извлечь media           ^| docx_media     ^| изображения без ручного zip
>>"%MENU_FILE%" echo [26] DOCX склеить через Word           ^| docx_merge     ^| Word COM вставляет документы подряд
>>"%MENU_FILE%" echo [27] Markdown-таблицы в XLSX           ^| md2xlsx        ^| выгрузить pipe-таблицы в Excel
>>"%MENU_FILE%" echo [28] Уменьшить поля ячеек таблиц       ^| table_margins  ^| компактные ячейки DOCX/XLSX
>>"%MENU_FILE%" echo.
>>"%MENU_FILE%" echo [A] Открыть папку input                ^| open_input     ^| explorer input
>>"%MENU_FILE%" echo [B] Открыть папку output               ^| open_output    ^| explorer output
>>"%MENU_FILE%" echo [C] Открыть папку logs                 ^| open_logs      ^| explorer logs
>>"%MENU_FILE%" echo [G] Открыть GUI shell                  ^| gui            ^| NiceGUI/pywebview оболочка
>>"%MENU_FILE%" echo [T] Служебный лаунчер                  ^| tools          ^| tools launcher
>>"%MENU_FILE%" echo [E] Английский лаунчер проекта         ^| launcher_en    ^| перейти в английскую оболочку
>>"%MENU_FILE%" echo [00] Выход                             ^| exit           ^| close

"%FZF_CMD%" --prompt="audion@office-kit [PROJECT-RU] > " --pointer=">" --header="Выберите инструмент:" --layout=reverse --border="rounded" --info=hidden --margin=1,2 < "%MENU_FILE%" > "%RES_FILE%"

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
if /I "%RAW%"=="launcher_en" goto LAUNCHER_EN
if /I "%RAW%"=="exit" exit /b 0
goto MAIN

:FALLBACK_MENU
echo === АУДИТ И СТИЛИ ===
echo [1] DOCX стилевой процессинг
echo [2] DOCX стилевой fix
echo [3] DOCX audit processor
echo [4] DOCX проверка текста
echo [5] DOCX исправить текст
echo.
echo === Найти и заменить ===
echo [6] DOCX регистр после запятых
echo [7] DOCX/XLSX морфологическая замена
echo.
echo === Контроль и очистка ===
echo [8] Проверка качества DOCX
echo [9] Строгая проверка DOCX
echo [A] DOCX удалить комментарии
echo [B] DOCX принять правки ^(простой^)
echo [C] DOCX чёрный текст + очистка
echo [D] DOCX финализация + проверка
echo.
echo === Сравнение и сверка ===
echo [E] DOCX поиск почти дубликатов
echo [F] DOCX diff двух документов
echo [G] XLSX сравнение значений ^(2 файла^)
echo [H] Сверка таблиц DOCX/XLSX
echo.
echo === ТАБЛИЦЫ WORD/EXCEL ===
echo [I] Сшить разрезанные после экспорта PDF
echo [J] Объединить только таблицы
echo [K] Объединить таблицы с заголовком
echo [L] Оптимизировать ширину таблиц
echo [M] Адаптировать таблицу к ориентации
echo [N] Вписать таблицы в документе
echo [O] Извлечь таблицы из DOCX
echo.
echo === Технические операции ===
echo [P] DOCX/PPTX извлечь media
echo [Q] DOCX склеить через Word
echo [R] Markdown-таблицы в XLSX
echo [S] Уменьшить поля ячеек таблиц
echo.
echo === Папки и служебные действия ===
echo [U] Открыть папку input
echo [V] Открыть папку output
echo [W] Открыть папку logs
echo [X] GUI shell
echo [Y] Служебный лаунчер
echo [Z] Английский лаунчер проекта
echo [0] Выход
echo.
choice /C 123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ0 /N /M "Выбор: "
if errorlevel 36 exit /b 0
if errorlevel 35 goto LAUNCHER_EN
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
set /p KEEP_WORDS=Не менять слова (CSV или путь к txt/json/yaml, пусто = нет): 
choice /C YN /N /M "Dry-run без записи DOCX? (Y/N): "
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
echo Выход:
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
choice /C RA /N /M "Режим: [R] заменить, [A] добавить рядом: "
if errorlevel 2 set "MORPH_MODE=append"
if errorlevel 1 set "MORPH_MODE=replace"
if /I "%MORPH_MODE%"=="append" set "MORPH_OUTDIR=morph_appended"
set /p FIND_TEXT=Найти (по умолчанию гаражи): 
if not defined FIND_TEXT set "FIND_TEXT=гаражи"
set /p REPLACE_TEXT=Замена или добавляемая фраза (по умолчанию подземные гаражи): 
if not defined REPLACE_TEXT set "REPLACE_TEXT=подземные гаражи"
set /p EXTS=Расширения (по умолчанию .docx,.xlsx): 
if not defined EXTS set "EXTS=.docx,.xlsx"
choice /C YN /N /M "Согласовать падеж замены? (Y/N): "
if errorlevel 2 set "MORPH_INFLECT=N"
if errorlevel 1 set "MORPH_INFLECT=Y"
choice /C YN /N /M "Dry-run без записи файлов? (Y/N): "
if errorlevel 2 set "MORPH_DRY=N"
if errorlevel 1 set "MORPH_DRY=Y"

set "MORPH_EXTRA="
if /I "%MORPH_INFLECT%"=="Y" set "MORPH_EXTRA=%MORPH_EXTRA% --inflect-replacement"
if /I "%MORPH_DRY%"=="Y" set "MORPH_EXTRA=%MORPH_EXTRA% --dry-run"

call :RUNPY "%CORE_DIR%\morph_replace.py" --input "%BASE_DIR%\input" --output "%BASE_DIR%\output\%MORPH_OUTDIR%" --find "%FIND_TEXT%" --replace "%REPLACE_TEXT%" --mode "%MORPH_MODE%" --recursive --extensions "%EXTS%" --report "%BASE_DIR%\report\morph_replace.json"%MORPH_EXTRA%
echo.
echo Выход:
if /I "%MORPH_DRY%"=="N" echo   output\%MORPH_OUTDIR%
echo   report\morph_replace.json
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_GATE
call :RUNPY "%CORE_DIR%\docx_quality_gate.py" --input "%BASE_DIR%\input" --out "%BASE_DIR%\report\docx_quality_gate_report.md"
echo.
echo Выходной файл:
echo   report\docx_quality_gate_report.docx
echo   report\docx_quality_gate_report.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_HSCAN
set "DOT=0"
echo.
choice /C YN /N /M "Проверять и отсутствие пробела после '.' тоже? (Y/N): "
if errorlevel 2 set "DOT=0"
if errorlevel 1 set "DOT=1"

if "%DOT%"=="1" (
  call :RUNPY "%CORE_DIR%\docx_text_hygiene_scan.py" --input "%BASE_DIR%\input" --out "%BASE_DIR%\report\docx_text_hygiene_report.md" --check-dot
) else (
  call :RUNPY "%CORE_DIR%\docx_text_hygiene_scan.py" --input "%BASE_DIR%\input" --out "%BASE_DIR%\report\docx_text_hygiene_report.md"
)
echo.
echo Выходной файл:
echo   report\docx_text_hygiene_report.docx
echo   report\docx_text_hygiene_report.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_HFIX
set "DOT=0"
echo.
choice /C YN /N /M "Исправлять отсутствие пробела после '.' тоже? (Y/N): "
if errorlevel 2 set "DOT=0"
if errorlevel 1 set "DOT=1"

if "%DOT%"=="1" (
  call :RUNPY "%CORE_DIR%\docx_text_hygiene_fix.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\hygiene_fixed" --fix-dot
) else (
  call :RUNPY "%CORE_DIR%\docx_text_hygiene_fix.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\hygiene_fixed"
)
echo.
echo Выходная папка:
echo   output\hygiene_fixed
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_NOCOMM
call :RUNPY "%CORE_DIR%\docx_strip_comments.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\no_comments"
echo.
echo Выходная папка:
echo   output\no_comments
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_ACCEPT
call :RUNPY "%CORE_DIR%\docx_accept_changes_simple.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\accepted_changes"
echo.
echo Выходная папка:
echo   output\accepted_changes
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_BLACK
call :RUNPY "%CORE_DIR%\docx_finalize_black_clean.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\final_black"
echo.
echo Выходная папка:
echo   output\final_black
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_DUPES
echo.
set "THR=0.99"
set /p THR=Порог похожести (по умолчанию 0.99):
if not defined THR set "THR=0.99"

set "DIFFS=0"
choice /C YN /N /M "Сохранять diff для совпавших пар? (Y/N): "
if errorlevel 2 set "DIFFS=0"
if errorlevel 1 set "DIFFS=1"

if "%DIFFS%"=="1" (
  call :RUNPY "%CORE_DIR%\docx_near_dupes.py" --input "%BASE_DIR%\input" --threshold %THR% --out "%BASE_DIR%\report\near_duplicates.md" --diff
) else (
  call :RUNPY "%CORE_DIR%\docx_near_dupes.py" --input "%BASE_DIR%\input" --threshold %THR% --out "%BASE_DIR%\report\near_duplicates.md"
)
echo.
echo Выходные файлы:
echo   report\near_duplicates.docx
echo   report\near_duplicates.md
echo   report\near_duplicates.json
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_PAIRDIFF
echo.
set "A="
set "B="
set /p A=DOCX A (путь относительно input\, можно с подпапками): 
if not defined A goto MAIN
set /p B=DOCX B (путь относительно input\, можно с подпапками): 
if not defined B goto MAIN

call :RUNPY "%CORE_DIR%\docx_pair_diff.py" --a "%BASE_DIR%\input\%A%" --b "%BASE_DIR%\input\%B%" --out "%BASE_DIR%\report\docx_pair_diff.md"
echo.
echo Выходные файлы:
echo   report\docx_pair_diff.docx
echo   report\docx_pair_diff.md
echo   report\docx_pair_diff.diff
echo   report\docx_pair_diff.json
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_STITCH
call :RUNPY "%CORE_DIR%\docx_table_stitcher.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\stitched" --report-dir "%BASE_DIR%\report\word_excel_tables\stitched" --recursive
echo.
echo Выходные файлы:
echo   output\word_excel_tables\stitched
echo   report\word_excel_tables\stitched
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_STITCH_HEADER
call :RUNPY "%CORE_DIR%\docx_table_stitcher_reconstruct_running_header.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\stitched_running_header" --report-dir "%BASE_DIR%\report\word_excel_tables\stitched_running_header" --recursive
echo.
echo Выходные файлы:
echo   output\word_excel_tables\stitched_running_header
echo   report\word_excel_tables\stitched_running_header
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_UNIFY_SAFE
call :RUNPY "%CORE_DIR%\docx_table_unifier.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\unified_safe" --report-dir "%BASE_DIR%\report\word_excel_tables\unified_safe" --all --recursive --mode safe --layout standard
echo.
echo Выходные файлы:
echo   output\word_excel_tables\unified_safe
echo   report\word_excel_tables\unified_safe
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_UNIFY_WIDTH
call :RUNPY "%CORE_DIR%\docx_table_unifier.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\unified_width_only" --report-dir "%BASE_DIR%\report\word_excel_tables\unified_width_only" --all --recursive --mode width-only --layout standard
echo.
echo Выходные файлы:
echo   output\word_excel_tables\unified_width_only
echo   report\word_excel_tables\unified_width_only
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_UNIFY_MERGED
call :RUNPY "%CORE_DIR%\docx_table_unifier.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\unified_merged_sections" --report-dir "%BASE_DIR%\report\word_excel_tables\unified_merged_sections" --all --recursive --mode safe --layout merged-sections
echo.
echo Выходные файлы:
echo   output\word_excel_tables\unified_merged_sections
echo   report\word_excel_tables\unified_merged_sections
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_OPT_WIDTHS
call :RUNPY "%CORE_DIR%\docx_table_width_optimizer.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\optimized_widths" --report-dir "%BASE_DIR%\report\word_excel_tables\optimized_widths" --all --recursive --mode preserve-width
echo.
echo Выходные файлы:
echo   output\word_excel_tables\optimized_widths
echo   report\word_excel_tables\optimized_widths
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_ADAPT_ORIENTATION
call :RUNPY "%CORE_DIR%\docx_table_width_optimizer.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\adapted_orientation" --report-dir "%BASE_DIR%\report\word_excel_tables\adapted_orientation" --all --recursive --mode fit-to-margins --fit-target page-setup --page-setup-margin-fallback keep --page-size A4 --page-orientation landscape
echo.
echo Выходные файлы:
echo   output\word_excel_tables\adapted_orientation
echo   report\word_excel_tables\adapted_orientation
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_FIT_MARGINS
call :RUNPY "%CORE_DIR%\docx_table_width_optimizer.py" --input-dir "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\word_excel_tables\fit_to_margins" --report-dir "%BASE_DIR%\report\word_excel_tables\fit_to_margins" --all --recursive --mode fit-to-margins
echo.
echo Выходные файлы:
echo   output\word_excel_tables\fit_to_margins
echo   report\word_excel_tables\fit_to_margins
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_TABLES
call :RUNPY "%CORE_DIR%\docx_extract_tables.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\tables" --out "%BASE_DIR%\report\docx_tables_index.md"
echo.
echo Выходные файлы:
echo   output\tables
echo   report\docx_tables_index.docx
echo   report\docx_tables_index.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_MEDIA
call :RUNPY "%CORE_DIR%\docx_extract_media.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\media" --out "%BASE_DIR%\report\office_media_index.md"
echo.
echo Выходные файлы:
echo   output\media
echo   report\office_media_index.docx
echo   report\office_media_index.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:TABLE_MARGINS
call :RUNPY "%CORE_DIR%\docx_xlsx_table_cell_margins.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\table_cell_margins" --report "%BASE_DIR%\report\table_cell_margins.md" --json-out "%BASE_DIR%\report\table_cell_margins.json" --margin-cm 0,1
echo.
echo Выходные файлы:
echo   output\table_cell_margins
echo   report\table_cell_margins.docx
echo   report\table_cell_margins.md
echo   report\table_cell_margins.json
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_MERGE
call :RUNPY "%CORE_DIR%\docx_merge.py" --input "%BASE_DIR%\input" --out "%BASE_DIR%\output\merged.docx" --report "%BASE_DIR%\report\docx_merge.md"
echo.
echo Выходные файлы:
echo   output\merged.docx
echo   report\docx_merge.docx
echo   report\docx_merge.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_STYLE
call :RUNPY "%CORE_DIR%\docx_style_processor.py" --input "%BASE_DIR%\input" --out "%BASE_DIR%\report\docx_style_processing.md"
echo.
echo Выходные файлы:
echo   report\docx_style_processing.docx
echo   report\docx_style_processing.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_STYLE_FIX
call :RUNPY "%CORE_DIR%\docx_style_processor.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\style_processed" --out "%BASE_DIR%\report\docx_style_processing.md" --fix
echo.
echo Выходные файлы:
echo   output\style_processed
echo   report\docx_style_processing.docx
echo   report\docx_style_processing.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:DOCX_AUDIT
set "AUDIT_FIX=0"
echo.
choice /C YN /N /M "Применить безопасные audit-правки? (Y/N): "
if errorlevel 2 set "AUDIT_FIX=0"
if errorlevel 1 set "AUDIT_FIX=1"

if "%AUDIT_FIX%"=="1" (
  call :RUNPY "%CORE_DIR%\docx_audit_processor.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\audit_processed" --report "%BASE_DIR%\report\docx_audit_processor.md" --json-out "%BASE_DIR%\report\docx_audit_processor.json" --fix
) else (
  call :RUNPY "%CORE_DIR%\docx_audit_processor.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\audit_processed" --report "%BASE_DIR%\report\docx_audit_processor.md" --json-out "%BASE_DIR%\report\docx_audit_processor.json"
)
echo.
echo Выходные файлы:
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

choice /C YN /N /M "Удалять комментарии? (Y/N) [по умолчанию Y]: "
if errorlevel 2 set "OPT_COMM=N"
if errorlevel 1 set "OPT_COMM=Y"

choice /C YN /N /M "Принимать правки (simple)? (Y/N) [по умолчанию Y]: "
if errorlevel 2 set "OPT_CHG=N"
if errorlevel 1 set "OPT_CHG=Y"

choice /C YN /N /M "Исправлять отсутствие пробела после '.' тоже? (Y/N) [по умолчанию N]: "
if errorlevel 2 set "OPT_DOT=N"
if errorlevel 1 set "OPT_DOT=Y"

set "EXTRA="
if /I "%OPT_COMM%"=="N" set "EXTRA=%EXTRA% --keep-comments"
if /I "%OPT_CHG%"=="N" set "EXTRA=%EXTRA% --keep-changes"
if /I "%OPT_DOT%"=="Y" set "EXTRA=%EXTRA% --fix-dot"

call :RUNPY "%CORE_DIR%\docx_oneclick_finalize_gate.py" --input "%BASE_DIR%\input" --outdir "%BASE_DIR%\output\final" --summary "%BASE_DIR%\report\summary_pass_fail.md"%EXTRA%
echo.
echo Выходная папка:
echo   output\final
echo Отчёты:
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
  echo РЕЗУЛЬТАТ: FAIL
) else (
  echo РЕЗУЛЬТАТ: PASS
)
echo Отчёт:
echo   report\docx_gate_hard_fail.docx
echo   report\docx_gate_hard_fail.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:XLSX_DIFF
echo.
set "A="
set "B="
set /p A=Файл A (путь относительно input\, можно с подпапками): 
if not defined A goto MAIN
set /p B=Файл B (путь относительно input\, можно с подпапками): 
if not defined B goto MAIN

call :RUNPY "%CORE_DIR%\xlsx_values_diff.py" --a "%BASE_DIR%\input\%A%" --b "%BASE_DIR%\input\%B%" --out "%BASE_DIR%\report\xlsx_values_diff_report.md"
echo.
echo Выходной файл:
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
set /p A=Файл A (xlsx/csv/docx, путь относительно input\): 
if not defined A goto MAIN
set /p B=Файл B (xlsx/csv/docx, путь относительно input\): 
if not defined B goto MAIN
set /p KEYS=Ключевые колонки (через запятую):
if not defined KEYS goto MAIN
set /p FIELDS=Поля для сравнения (через запятую, пусто = авто):

if defined FIELDS (
  call :RUNPY "%CORE_DIR%\tabular_reconcile_4lists.py" --a "%BASE_DIR%\input\%A%" --b "%BASE_DIR%\input\%B%" --key-csv "%KEYS%" --fields-csv "%FIELDS%" --out "%BASE_DIR%\report\reconcile_4lists.md"
) else (
  call :RUNPY "%CORE_DIR%\tabular_reconcile_4lists.py" --a "%BASE_DIR%\input\%A%" --b "%BASE_DIR%\input\%B%" --key-csv "%KEYS%" --out "%BASE_DIR%\report\reconcile_4lists.md"
)

echo.
echo Выходной файл:
echo   report\reconcile_4lists.docx
echo   report\reconcile_4lists.md
if not defined AUDION_NO_PAUSE pause
goto MAIN

:MD2XLSX
echo.
set "MD="
set "OUTX="
set /p MD=Markdown-файл (в input\):
if not defined MD goto MAIN
set /p OUTX=Имя выходного xlsx (например report.xlsx):
if not defined OUTX set "OUTX=report.xlsx"

call :RUNPY "%CORE_DIR%\md_tables_to_xlsx.py" --md "%BASE_DIR%\input\%MD%" --out "%BASE_DIR%\output\%OUTX%"
echo.
echo Выходной файл:
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
  echo [WARN] launcher_gui.cmd не найден.
  if not defined AUDION_NO_PAUSE pause
)
goto MAIN

:TOOLS
call "%BASE_DIR%\launcher_tools.cmd"
goto MAIN

:LAUNCHER_EN
if exist "%BASE_DIR%\launcher_project.cmd" (
  call "%BASE_DIR%\launcher_project.cmd"
) else (
  echo.
  echo [WARN] launcher_project.cmd не найден.
  if not defined AUDION_NO_PAUSE pause
)
goto MAIN

:NO_PYTHON
cls
echo [ERROR] Не удалось определить Python runtime.
echo.
echo Поддерживаемые пути:
echo   runtime\python.exe
echo   runtime\python\python.exe
echo   py -3.12
echo   python
echo.
echo Используйте builder_main.cmd или install\Build_Portable_Env_Build.cmd
if not defined AUDION_NO_PAUSE pause
exit /b 1

:RUNPY
set "TARGET=%~1"
shift
if not exist "%TARGET%" (
  echo [ERROR] Скрипт не найден:
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
