import pandas as pd

def openFile(path_file=None):
    excel_file = 'C://МДТ//АСОЗ ЦТ 2025 1 этап.xlsx'
    df = pd.read_excel(excel_file)
    return df, df.columns

data, columns = openFile()

# Фильтрация по Московской ЖД
def filterDepartment(data, columns):
    department_column = 'ЖД'
    if department_column in columns:
        # Фильтруем данные по Московской ЖД
        filter_data = data[data[department_column] == 'Московская ЖД']
        return filter_data
    else:
        print(f"Ошибка: Столбец '{department_column}' не найден. Доступные столбцы: {list(columns)}")
        return None

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
