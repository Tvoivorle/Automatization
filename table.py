import pandas as pd

# Функция для извлечения номера услуги (формат xx.xx)
def extract_service_number(service):
    try:
        if isinstance(service, str):
            return service.split(' ')[0]  # Извлекаем номер услуги
        else:
            return None  # Если значение не строка, возвращаем None
    except Exception as e:
        print(f"Ошибка при обработке значения: {service} — {e}")
        return None

# Создание сводной таблицы
def create_pivot_table(data):
    # Добавляем колонку с номером услуги
    data['Номер услуги'] = data['Услуга'].apply(extract_service_number)

    # Удаляем строки, где номер услуги не определён
    data = data.dropna(subset=['Номер услуги'])

    # Группировка данных по 'Номер услуги', 'Услуга', 'Иерархия подразделения'
    grouped = (
        data.groupby(['Номер услуги', 'Услуга', 'Иерархия подразделения'])
        .size()
        .reset_index(name='Количество')
    )

    return grouped

# Основной код
def main():
    try:
        # Загружаем отфильтрованные данные
        filtered_data = pd.read_excel('C://МДТ//АСОЗ ЦТ 2025 1 этап_final.xlsx')

        # Создаём сводную таблицу
        pivot_table = create_pivot_table(filtered_data)

        # Сохраняем сводную таблицу на новый лист Excel
        with pd.ExcelWriter('C://МДТ//АСОЗ ЦТ 2025 1 этап_сводная таблица.xlsx') as writer:
            pivot_table.to_excel(writer, sheet_name='Сводная таблица', index=False)

        print("Сводная таблица успешно создана и сохранена.")
    except Exception as e:
        print(f"Произошла ошибка!: {e}")

if __name__ == '__main__':
    main()
