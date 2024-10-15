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

status_filtered_data = filterStatus(filtered_data, columns)

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

# Функция для очистки данных по иерархии подразделения
def cleanHierarchy(data, column_name, filial_list):
    if column_name in data.columns:
        # Очистка данных по наименованию функционального филиала
        cleaned_data = data.copy()
        cleaned_data[column_name] = cleaned_data[column_name].apply(
            lambda x: next((filial for filial in filial_list if filial in x), x))
        return cleaned_data
    else:
        print(f"Ошибка: Столбец '{column_name}' не найден.")
        return None

# Загрузка данных о функциональных филиалах
filial_df = pd.read_excel('C://МДТ//Функциональные филиалы.xlsx', sheet_name='Общая')
filial_list = filial_df['Наименование функционального филиала'].tolist()

# Очистка столбца "Иерархия подразделения"
if filtered_by_date is not None:
    filial_data = cleanHierarchy(filtered_by_date, 'Иерархия подразделения', filial_list)

    if filial_data is not None:
        # Сохранение результата в новый Excel файл
        filial_data.to_excel('C://МДТ//АСОЗ ЦТ 2025 1 этап_отфильтрованный.xlsx', index=False)
        print("\nФайл успешно сохранен.")
else:
    print("Данные после фильтрации отсутствуют.")


# Проверка и вывод итоговых данных
if filial_data is not None and not filial_data.empty:
    num_values = filial_data.size  # Общее количество значений
    num_rows, num_cols = filial_data.shape  # Количество строк и столбцов

    print(f"\nОбщее количество значений в таблице после фильтрации по дате окончания: {num_values}")
    print(f"Количество строк: {num_rows}, Количество столбцов: {num_cols}")
    print("Данные после фильтрации по дате окончания:")
    print(filial_data)
else:
    print("Нет данных по указанным фильтрам по дате окончания.")