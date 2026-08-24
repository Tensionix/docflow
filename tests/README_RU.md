# Проверки Audion DocFlow

Быстрый smoke компилирует backend, проверяет импорт NiceGUI, каноническую историю Workbench, маршрутизацию выбранных Source/Target, защиту удаления и работу одиночного DOCX через реальный backend:

```powershell
& '.\runtime\python.exe' '.\tests\smoke.py' --quick
```

Полный smoke дополнительно запускает границы ответственности текстовой гигиены:

```powershell
& '.\runtime\python.exe' '.\tests\smoke.py' --full
```

Тесты используют временные каталоги вне пользовательских `input`, `output` и `report` и удаляют их автоматически.
