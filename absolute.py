import pandas as pd
import os
import numpy as np

def extract_columns_to_dict_multiple(file_path, key_column_index, value_column_indices, value_field_names):
    """
    Извлекает указанные столбцы из Excel файла и сохраняет их в словарь с вложенными словарями.

    :param file_path: Путь к Excel файлу
    :param key_column_index: Индекс столбца для ключей (начиная с 0)
    :param value_column_indices: Список индексов столбцов для значений (начиная с 0)
    :param value_field_names: Список имен полей для значений
    :return: Словарь с ключами и вложенными словарями значений, DataFrame
    """
    try:
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"Файл по пути '{file_path}' не найден.")

        # Чтение Excel файла с учётом заголовков
        df = pd.read_excel(file_path)
        print(f"\nЧтение файла '{file_path}':")
        print(df.head())  # Отладочный вывод первых 5 строк

        # Проверка индексов столбцов
        if key_column_index >= len(df.columns):
            raise IndexError("Индекс столбца для ключей выходит за пределы доступных столбцов.")

        for idx in value_column_indices:
            if idx >= len(df.columns):
                raise IndexError(f"Индекс столбца для значений ({idx}) выходит за пределы доступных столбцов.")

        key_column = df.columns[key_column_index]
        value_columns = [df.columns[idx] for idx in value_column_indices]

        print(f"\nИспользуемый ключевой столбец: '{key_column}'")
        print(f"Используемые столбцы значений: {value_columns}")

        # Удаление строк с пустыми значениями только в ключевом столбце
        df = df.dropna(subset=[key_column])
        print(f"\nПосле удаления строк без ключей:")
        print(df.head())  # Отладочный вывод первых 5 строк

        # Проверка уникальности ключей
        if df[key_column].duplicated().any():
            print(f"В столбце '{key_column}' обнаружены дублирующиеся ключи.")
            # Дополнительная обработка, если необходимо
            # Например, удалить дубликаты или объединить значения

        # Преобразование ключевого столбца к строковому типу для консистентности
        df[key_column] = df[key_column].astype(str)

        # Преобразование только необходимых столбцов к строковому типу (например, 'class')
        if 'class' in value_field_names:
            class_idx = value_field_names.index('class')
            class_col = value_columns[class_idx]
            df[class_col] = df[class_col].astype(str)

        # Остальные столбцы оставляем в исходных типах, чтобы сохранить NaN

        key_data = df[key_column]
        value_data = df[value_columns]

        # Создание словаря с вложенными словарями
        column_dict = {}
        for key, values in zip(key_data, value_data.itertuples(index=False)):
            column_dict[key] = {field: (value if pd.notna(value) else np.nan) for field, value in zip(value_field_names, values)}

        return column_dict, df  # Возвращаем также DataFrame для дальнейшей обработки

    except FileNotFoundError as fnf_error:
        print(fnf_error)
    except IndexError as ie:
        print(ie)
    except Exception as e:
        print(f"Произошла ошибка: {e}")

def compare_dicts(dict1, dict2):
    """
    Сравнивает два словаря и возвращает различия.

    :param dict1: Первый словарь
    :param dict2: Второй словарь
    :return: Три множества/словаря: ключи только в dict1, только в dict2, и общие ключи с их значениями из обоих словарей
    """
    # Ключи, присутствующие только в dict1
    only_in_dict1 = set(dict1.keys()) - set(dict2.keys())

    # Ключи, присутствующие только в dict2
    only_in_dict2 = set(dict2.keys()) - set(dict1.keys())

    # Все ключи, присутствующие в обоих словарях
    present_in_both = {k: (dict1[k], dict2[k]) for k in dict1.keys() & dict2.keys()}

    return only_in_dict1, only_in_dict2, present_in_both

if __name__ == "__main__":
    # Параметры для table1.xlsx
    excel_file1 = 'table1.xlsx'
    key_column_index1 = 7    # Индекс столбца для ключей (начиная с 0)
    value_column_indices1 = [11, 13, 14, 15, 16]  # Индексы столбцов для значений
    value_field_names1 = ['class', 'lactation', 'total_milk', 'milk305', 'fat']  # Имена полей

    # Параметры для table2.xlsx
    excel_file2 = 'table2.xlsx'
    key_column_index2 = 0    # Индекс столбца для ключей (начиная с 0)
    value_column_indices2 = [8, 1, 3, 2, 4]  # Индексы столбцов для значений
    value_field_names2 = ['class', 'lactation', 'total_milk', 'milk305', 'fat']  # Имена полей

    # Извлечение словарей из обоих файлов вместе с DataFrame первого файла
    dict1, df1 = extract_columns_to_dict_multiple(excel_file1, key_column_index1, value_column_indices1, value_field_names1)
    dict2, df2 = extract_columns_to_dict_multiple(excel_file2, key_column_index2, value_column_indices2, value_field_names2)

    # Проверка успешного извлечения словарей
    if dict1:
        print("\nИзвлечённый словарь из table1.xlsx (dict1):")
        # Для удобства, выводим только первые 3 элемента
        for i, (k, v) in enumerate(dict1.items()):
            print(f"{k}: {v}")
            if i >= 2:
                print("...")
                break
    else:
        print("\ndict1 не был создан.")

    if dict2:
        print("\nИзвлечённый словарь из table2.xlsx (dict2):")
        # Для удобства, выводим только первые 3 элемента
        for i, (k, v) in enumerate(dict2.items()):
            print(f"{k}: {v}")
            if i >= 2:
                print("...")
                break
    else:
        print("\ndict2 не был создан.")

    # Сравнение словарей, если оба были успешно созданы
    if dict1 and dict2:
        only_in_dict1, only_in_dict2, present_in_both = compare_dicts(dict1, dict2)

        print("\n--- Сравнение словарей ---")

        # Ключи, присутствующие только в dict1
        print("\nКлючи, присутствующие только в dict1:")
        if only_in_dict1:
            for key in only_in_dict1:
                print(f"{key}: {dict1[key]}")
        else:
            print("Нет уникальных ключей в dict1.")

        # Ключи, присутствующие только в dict2
        print("\nКлючи, присутствующие только в dict2:")
        if only_in_dict2:
            for key in only_in_dict2:
                print(f"{key}: {dict2[key]}")
        else:
            print("Нет уникальных ключей в dict2.")

        # Общие ключи с их значениями из обоих словарей
        print("\nОбщие ключи с их значениями из обоих словарей:")
        if present_in_both:
            for key, (val1, val2) in present_in_both.items():
                print(f"{key}: dict1 = {val1}, dict2 = {val2}")
        else:
            print("Нет общих ключей в обоих словарях.")

        # Сохранение общих ключей в переменную для дальнейшей работы
        common_keys = present_in_both

        # Обновление значений в dict1 на основе dict2 для общих ключей
        for key, (val1, val2) in common_keys.items():
            # Обновляем все поля из dict2 в dict1
            dict1[key].update(val2)
            print(f"Ключ {key} обновлён значениями из dict2.")

        # Применение обновлений к DataFrame table1.xlsx (df1)
        key_column = df1.columns[key_column_index1]

        for key, updated_values in dict1.items():
            # Найти индекс строки с этим ключом
            row_indices = df1.index[df1[key_column] == key].tolist()
            if row_indices:
                idx = row_indices[0]
                # Обновить соответствующие столбцы
                for field, value in updated_values.items():
                    if field in value_field_names1:
                        # Найти индекс столбца по имени поля
                        field_idx = value_field_names1.index(field)
                        # Найти соответствующий столбец из value_column_indices1
                        col_name = df1.columns[value_column_indices1[field_idx]]
                        # Обновить значение, сохраняя NaN как пустые ячейки
                        if pd.isna(value) or (isinstance(value, str) and value.lower() == 'nan'):
                            df1.at[idx, col_name] = np.nan
                        else:
                            # Если поле должно быть числовым, попытайтесь преобразовать
                            if field in ['lactation', 'total_milk', 'milk305', 'fat']:
                                try:
                                    df1.at[idx, col_name] = float(value)
                                except ValueError:
                                    # Если преобразование не удалось, сохранить как строку
                                    df1.at[idx, col_name] = value
                            else:
                                df1.at[idx, col_name] = value
            else:
                print(f"Ключ {key} не найден в DataFrame.")

        # Сохранение обновлённого DataFrame обратно в table1.xlsx
        try:
            # Создание резервной копии
            backup_file = 'table1_backup.xlsx'
            df1.to_excel(backup_file, index=False)
            print(f"\nРезервная копия сохранена в '{backup_file}'.")

            # Сохранение обновлённого DataFrame
            df1.to_excel(excel_file1, index=False, na_rep='')
            print(f"Обновлённые данные сохранены обратно в '{excel_file1}'.")
        except Exception as e:
            print(f"Не удалось сохранить обновлённые данные в '{excel_file1}': {e}")

        # Пример дальнейшей работы с обновлённым dict1
        print("\n--- Обновлённый dict1 ---")
        for i, (k, v) in enumerate(dict1.items()):
            print(f"{k}: {v}")
            if i >= 2:
                print("...")
                break

    else:
        print("\nНе удалось создать оба словаря для сравнения.")
