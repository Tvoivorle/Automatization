import pandas as pd
def openFile(path_file=None):
    excel_file = 'D://РЖД//АСОЗ ЦТ 2025 1 этап.xlsx'
    df = pd.read_excel(excel_file)
    return df, df.columns

data, columns = openFile()
# Вывод количества значений в таблице
num_values = data.size  # Общее количество значений
num_rows, num_cols = data.shape  # Количество строк и столбцов

print(f"Общее количество значений в таблице: {num_values}")
print(f"Количество строк: {num_rows}, Количество столбцов: {num_cols}")
print(columns)  # Для вывода названий столбцов

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

# Проверяем и выводим информацию об отфильтрованных данных
if filtered_data is not None:
    if not filtered_data.empty:
        num_values = filtered_data.size  # Общее количество значений
        num_rows, num_cols = filtered_data.shape  # Количество строк и столбцов

        print(f"Общее количество значений в отфильтрованной таблице: {num_values}")
        print(f"Количество строк: {num_rows}, Количество столбцов: {num_cols}")
        print("Отфильтрованные данные:")
        print(filtered_data)
    else:
        print("Нет данных по указанному фильтру.")
else:
    print("Фильтрация не выполнена.")