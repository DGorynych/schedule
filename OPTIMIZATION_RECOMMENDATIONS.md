# Рекомендации по оптимизации schedule_client.py

## 🔴 Критические проблемы производительности

### 1. **Множественные открытия/закрытия соединений с БД**
**Проблема:** Каждый метод открывает новое соединение с SQLite, что очень медленно.

**Текущий код:**
```python
def save_pending_change(self, request):
    conn = sqlite3.connect(self.local_db)  # Открытие
    # ... операции ...
    conn.close()  # Закрытие
```

**Решение:** Использовать контекстный менеджер или пул соединений:
```python
# Вариант 1: Контекстный менеджер (проще)
def save_pending_change(self, request):
    with sqlite3.connect(self.local_db) as conn:
        cursor = conn.cursor()
        cursor.execute(...)
        conn.commit()  # Автоматически закрывается

# Вариант 2: Переиспользуемое соединение (быстрее)
class ScheduleManager:
    def __init__(self, ...):
        self._db_conn = sqlite3.connect(self.local_db, check_same_thread=False)
        self._db_conn.row_factory = sqlite3.Row
    
    def _get_connection(self):
        return self._db_conn
```

**Ожидаемый эффект:** Ускорение операций с БД в 5-10 раз.

---

### 2. **Избыточные вызовы update_counters()**
**Проблема:** `update_counters()` вызывается при каждом изменении ячейки таблицы.

**Текущий код:**
```python
def setItem(self, row, column, item):
    super().setItem(row, column, item)
    if not self.is_loading and column < self.columnCount() - 4:
        self.parent_widget.update_counters()  # Вызывается для КАЖДОЙ ячейки!
```

**Решение:** Использовать debounce или батчинг:
```python
from PyQt5.QtCore import QTimer

class NoteTableWidget(QTableWidget):
    def __init__(self, parent=None):
        # ...
        self._update_timer = QTimer()
        self._update_timer.setSingleShot(True)
        self._update_timer.timeout.connect(self._delayed_update_counters)
        self._pending_updates = False
    
    def setItem(self, row, column, item):
        super().setItem(row, column, item)
        if not self.is_loading and column < self.columnCount() - 4:
            self._pending_updates = True
            self._update_timer.start(300)  # Обновление через 300мс
    
    def _delayed_update_counters(self):
        if self._pending_updates:
            self.parent_widget.update_counters()
            self._pending_updates = False
```

**Ожидаемый эффект:** При массовой вставке (например, paste) ускорение в 10-50 раз.

---

## 🟡 Важные оптимизации

### 3. **Дублирование парсинга периода**
**Проблема:** Код парсинга периода повторяется в 10+ местах.

**Текущий код:**
```python
base_name = period
if ' (' in period:
    base_name = period.split(' (')[0]
month_name, year = base_name.rsplit(' ', 1)
year = int(year)
month_num = 1
for num, name in self.month_names.items():
    if name == month_name:
        month_num = num
        break
```

**Решение:** Создать утилиту с кэшированием:
```python
from functools import lru_cache

class PeriodParser:
    @staticmethod
    @lru_cache(maxsize=128)
    def parse_period(period, month_names):
        """Парсит период и возвращает (year, month_num, variant)"""
        base_name = period.split(' (')[0] if ' (' in period else period
        variant = period.split(' (')[1].rstrip(')') if ' (' in period else ""
        
        month_name, year = base_name.rsplit(' ', 1)
        year = int(year)
        
        month_num = next((num for num, name in month_names.items() if name == month_name), 1)
        
        return year, month_num, variant
    
    @staticmethod
    @lru_cache(maxsize=128)
    def get_days_in_month(year, month_num):
        return monthrange(year, month_num)[1]
```

**Ожидаемый эффект:** Ускорение парсинга в 2-3 раза + меньше дублирования кода.

---

### 4. **Многократные вычисления weekday()**
**Проблема:** `datetime(year, month_num, day).weekday() == 6` вычисляется многократно для одних и тех же дат.

**Текущий код:**
```python
for col, actual_day in enumerate(day_mapping):
    day_of_week = datetime(year, month_num, actual_day).weekday()  # Вычисляется каждый раз
    if day_of_week == 6:
        # ...
```

**Решение:** Вычислить один раз и кэшировать:
```python
# В load_data:
is_sunday_mapping = []
for day in range(1, days_in_month + 1):
    is_sunday = datetime(year, month_num, day).weekday() == 6
    is_sunday_mapping.append(is_sunday)

# Использовать:
is_sunday = is_sunday_mapping[day_index]
```

**Ожидаемый эффект:** Ускорение в 2-3 раза при работе с датами.

---

### 5. **Неоптимальная работа с Excel (openpyxl)**
**Проблема:** Создание стилей в цикле, множественные обращения к ячейкам.

**Текущий код:**
```python
for emp in employees:
    # ...
    cell = ws.cell(row=start_row, column=col+2)
    cell.alignment = center_alignment
    cell.border = thin_border
    cell.fill = status_styles[status]
```

**Решение:** Использовать batch операции и именованные стили:
```python
# Создать стили один раз
from openpyxl.styles import NamedStyle

center_style = NamedStyle(name="center_style")
center_style.alignment = center_alignment
center_style.border = thin_border

# Применить к диапазону
from openpyxl.utils import get_column_letter

# Применить стиль к целому столбцу
for col_idx in range(2, days_in_month + 2):
    col_letter = get_column_letter(col_idx)
    for row in ws[col_letter]:
        row[0].style = center_style
```

**Ожидаемый эффект:** Ускорение экспорта в 1.5-2 раза.

---

## 🟢 Улучшения качества кода

### 6. **Кэширование часто используемых данных**
**Проблема:** `self.db.periods` и `self.db.employees` перезагружаются при каждом обращении.

**Текущий код:**
```python
@property
def periods(self):
    if not self._periods_cache:
        self._periods_cache = self.get_periods()
    return self._periods_cache
```

**Решение:** Добавить инвалидацию кэша при изменениях:
```python
def add_period(self, period):
    # ... добавление ...
    self._periods_cache = None  # Инвалидировать кэш

def get_periods(self, force_refresh=False):
    if force_refresh or not self._periods_cache:
        self._periods_cache = self._load_periods_from_db()
    return self._periods_cache
```

---

### 7. **Оптимизация работы с таблицей**
**Проблема:** Множественные вызовы `setItem`, `setHorizontalHeaderLabels` вызывают перерисовку.

**Решение:** Использовать `setUpdatesEnabled(False)` при массовых операциях:
```python
def load_data(self, period):
    self.table.setUpdatesEnabled(False)  # Отключить перерисовку
    try:
        # ... массовые операции ...
    finally:
        self.table.setUpdatesEnabled(True)  # Включить обратно
        self.table.viewport().update()  # Одна перерисовка в конце
```

**Ожидаемый эффект:** Ускорение загрузки данных в 3-5 раз.

---

### 8. **Оптимизация поиска месяца по имени**
**Проблема:** Линейный поиск в словаре `month_names.items()`.

**Текущий код:**
```python
month_num = 1
for num, name in self.month_names.items():
    if name == month_name:
        month_num = num
        break
```

**Решение:** Создать обратный словарь:
```python
# В __init__:
self.month_names = {1: "Январь", 2: "Февраль", ...}
self.month_names_reverse = {name: num for num, name in self.month_names.items()}

# Использование:
month_num = self.month_names_reverse.get(month_name, 1)
```

**Ожидаемый эффект:** O(1) вместо O(n).

---

## 📊 Приоритеты внедрения

1. **Высокий приоритет:**
   - ✅ Debounce для `update_counters()` (#2)
   - ✅ Переиспользование соединений БД (#1)
   - ✅ Отключение обновлений при загрузке (#7)

2. **Средний приоритет:**
   - ✅ Кэширование парсинга периода (#3)
   - ✅ Кэширование вычислений weekday (#4)
   - ✅ Обратный словарь для месяцев (#8)

3. **Низкий приоритет:**
   - ✅ Оптимизация Excel экспорта (#5)
   - ✅ Улучшение кэширования (#6)

---

## 🧪 Тестирование производительности

Рекомендуется добавить профилирование:

```python
import cProfile
import pstats

def profile_export():
    profiler = cProfile.Profile()
    profiler.enable()
    app.export_to_excel()
    profiler.disable()
    
    stats = pstats.Stats(profiler)
    stats.sort_stats('cumulative')
    stats.print_stats(20)
```

---

## 📈 Ожидаемые результаты

После внедрения всех оптимизаций:
- **Загрузка данных:** ускорение в 3-5 раз
- **Сохранение данных:** ускорение в 5-10 раз
- **Экспорт в Excel:** ускорение в 1.5-2 раза
- **Отклик UI:** улучшение в 5-10 раз при массовых операциях

---

## 🔧 Дополнительные рекомендации

1. **Использовать индексы БД:**
   ```sql
   CREATE INDEX idx_periods_period ON periods(period);
   CREATE INDEX idx_schedules_period_employee ON schedules(period_id, employee_name);
   ```

2. **Асинхронные операции:**
   Для экспорта больших файлов использовать QThread, чтобы не блокировать UI.

3. **Логирование:**
   Добавить логирование медленных операций для дальнейшей оптимизации.
