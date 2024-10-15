import pandas as pd
from datetime import datetime
def openFile(path_file=None):
    excel_file = 'C://МДТ//АСОЗ ЦТ 2025 1 этап.xlsx'
    df = pd.read_excel(excel_file)
    return df, df.columns

data, columns = openFile()

# Вывод количества значений в ИЗНАЧАЛЬНОЙ таблице
num_values = data.size  # Общее количество значений
num_rows, num_cols = data.shape  # Количество строк и столбцов
print(f"Общее количество значений в таблице: {num_values}")
print(f"Количество строк: {num_rows}, Количество столбцов: {num_cols}")
print(columns)  # Для вывода названий столбцов

#Фильтрация по "Московская ЖД"
def filterDepartment(data, columns):
    # Заголовок для фильтрации
    department_column = 'ЖД'
    if department_column in columns:
        # Фильтруем данные
        filter_data = data[data[department_column] == 'Московская ЖД']
        return filter_data
    else:
        print(f"Ошибка: Столбец '{department_column}' не найден. Доступные столбцы: {list(columns)}")
        return None

# Фильтруем данные
filtered_data = filterDepartment(data, columns)

# Фильтрация по статусу заявки
def filterStatus(filtered_data, columns):
    status_column = 'Статус заявки'
    if status_column in columns:
        # Значения, которые нужно оставить
        statuses_to_keep = ['1-На согласовании', '2-На утверждении', '3-Утверждена']
        # Фильтрация по статусам
        status_filtered_data = filtered_data[filtered_data[status_column].isin(statuses_to_keep)]
        return status_filtered_data
    else:
        print(f"Ошибка: Столбец '{status_column}' не найден. Доступные столбцы: {list(columns)}")
        return None

# Проверяем и выводим информацию об отфильтрованных данных
if filtered_data is not None:
    status_filtered_data = filterStatus(filtered_data, columns)

    # Проверяем и выводим информацию об отфильтрованных данных
    if status_filtered_data is not None and not status_filtered_data.empty:
        num_values = status_filtered_data.size  # Общее количество значений
        num_rows, num_cols = status_filtered_data.shape  # Количество строк и столбцов

        print(f"Общее количество значений в отфильтрованной таблице: {num_values}")
        print(f"Количество строк: {num_rows}, Количество столбцов: {num_cols}")
        print("Отфильтрованные данные по статусам:")
        print(status_filtered_data)
    else:
        print("Нет данных по указанным фильтрам.")
else:
    print("Фильтрация по Московской ЖД не выполнена.")

# Фильтрация по дате окончания
# Класс для фильтрации по дате окончания
class DateFilter:
    def __init__(self, stage):
        self.stage = stage
        self.cutoff_date = self.set_cutoff_date()

    def set_cutoff_date(self):
        # Определяем дату окончания в зависимости от этапа
        if self.stage == 'первый этап':
            return datetime(2024, 10, 1)
        elif self.stage == 'второй этап':
            return datetime(2024, 1, 1)
        elif self.stage == 'третий этап':
            return datetime(2024, 4, 1)
        elif self.stage == 'четвертый этап':
            return datetime(2024, 7, 1)
        else:
            raise ValueError("Некорректный этап. Используйте один из: 'первый этап', 'второй этап', 'третий этап', 'четвертый этап'.")

    def filter_by_date(self, data, date_column):
        if date_column in data.columns:
            # Фильтрация данных по дате
            filtered_data = data[data[date_column] >= self.cutoff_date]
            return filtered_data
        else:
            print(f"Ошибка: Столбец '{date_column}' не найден. Доступные столбцы: {list(data.columns)}")
            return None


# Применение фильтрации по дате окончания
date_filter = DateFilter(stage='первый этап')
date_column = 'Дата окончания доступа к ИС'  # Предполагаем, что у нас есть этот столбец
filtered_by_date = date_filter.filter_by_date(status_filtered_data, date_column)

# Проверка и вывод итоговых данных
if filtered_by_date is not None and not filtered_by_date.empty:
    num_values = filtered_by_date.size  # Общее количество значений
    num_rows, num_cols = filtered_by_date.shape  # Количество строк и столбцов

    print(f"Общее количество значений в таблице после фильтрации по дате окончания: {num_values}")
    print(f"Количество строк: {num_rows}, Количество столбцов: {num_cols}")
    print("Данные после фильтрации по дате окончания:")
    print(filtered_by_date)
else:
    print("Нет данных по указанным фильтрам по дате окончания.")
