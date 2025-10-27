import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

# Функция для выбора файла Excel
def choose_input_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])  # Открываем окно выбора файла
    input_file_entry.delete(0, tk.END)  # Очистка предыдущего пути
    input_file_entry.insert(0, file_path)  # Вставляем путь выбранного файла

# Функция для выбора пути сохранения
def choose_output_path():
    folder_path = filedialog.askdirectory()  # Открываем окно выбора папки
    output_path_entry.delete(0, tk.END)  # Очистка предыдущего пути
    output_path_entry.insert(0, folder_path)  # Вставляем путь выбранной папки

# Функция для запуска анализа
def run_analysis():
    input_file = input_file_entry.get()  # Получаем путь к файлу
    output_path = output_path_entry.get()  # Получаем путь для сохранения

    if not input_file or not output_path:
        result_label.config(text="Пожалуйста, выберите файл и путь для сохранения.")
        return

    # Загрузка данных из Excel
    data = pd.read_excel(input_file)

    # Переименование столбцов для удобства
    data = data.rename(columns={'Штуки': 'Продажи, шт', 'Выручка, Р': 'Продажи, руб', 'ГодМесяц': 'Месяц'})

    # 1. Подготовка данных: агрегируем по товару (variant_id) и месяцу
    data['Год'] = data['Месяц'].str[:4]  # Извлекаем год
    data['Месяц_номер'] = data['Месяц'].str[5:7]  # Извлекаем месяц

    # Аггрегируем данные по товарам (variant_id)
    data_aggregated = data.groupby(['variant_id', 'Год', 'Месяц_номер'])['Продажи, шт'].sum().reset_index()

    # Создаем таблицу с месячными продажами по товарам
    pivot_data = data_aggregated.pivot_table(index='variant_id', columns='Месяц_номер', values='Продажи, шт', aggfunc='sum', fill_value=0)

    # Добавляем столбцы с общими продажами (рубли и штуки)
    total_sales = data.groupby('variant_id')['Продажи, руб'].sum()
    total_quantity = data.groupby('variant_id')['Продажи, шт'].sum()

    pivot_data['Общие_продажи_руб'] = total_sales
    pivot_data['Общие_продажи_шт'] = total_quantity

    # 2. Расчет коэффициента вариации (CV) для каждого товара
    pivot_data['CV'] = pivot_data.iloc[:, :-2].std(axis=1) / pivot_data.iloc[:, :-2].mean(axis=1) * 100

    # 3. Вычисление порогов для XYZ на основе данных
    cv_min = pivot_data['CV'].min()
    cv_max = pivot_data['CV'].max()

    # Определим пороги для XYZ
    threshold_x = cv_min + (cv_max - cv_min) * 0.3  # Порог для X: до 30% CV
    threshold_y = cv_min + (cv_max - cv_min) * 0.6  # Порог для Y: до 60% CV
    threshold_z = cv_max  # Порог для Z: выше 60%

    # 4. Группировка по XYZ на основе рассчитанных порогов
    pivot_data['XYZ'] = pd.cut(pivot_data['CV'], bins=[0, threshold_x, threshold_y, threshold_z], labels=['X', 'Y', 'Z'])

    # 5. Добавляем информацию о категории ABC
    pivot_data = pivot_data.sort_values(by='Общие_продажи_руб', ascending=False)

    # Рассчитываем долю выручки и кумулятивную долю
    pivot_data['Продажа доля'] = pivot_data['Общие_продажи_руб'] / pivot_data['Общие_продажи_руб'].sum()
    pivot_data['Кумулятивная доля'] = pivot_data['Продажа доля'].cumsum()

    # Применяем классификацию ABC
    pivot_data['ABC'] = pd.cut(pivot_data['Кумулятивная доля'], bins=[0, 0.8, 0.95, 1], labels=['A', 'B', 'C'])

    # 6. Комбинированный ABC/XYZ анализ
    pivot_data['ABC/XYZ'] = pivot_data['ABC'].astype(str) + '-' + pivot_data['XYZ'].astype(str)

    # 7. Применяем рекомендации на основе столбца ABC/XYZ
    recommendation_dict = {
        'A-X': 'KEEP',
        'A-Y': 'KEEP',
        'A-Z': 'CONTROL',
        'B-X': 'KEEP',
        'B-Y': 'CONTROL',
        'B-Z': 'OPTIMIZE',
        'C-X': 'OPTIMIZE',
        'C-Y': 'OPTIMIZE',
        'C-Z': 'DROP',
    }

    # Применяем словарь рекомендаций
    pivot_data['Рекомендации'] = pivot_data['ABC/XYZ'].map(recommendation_dict)

    # Сохраняем результаты в новый файл
    output_file = os.path.join(output_path, 'abc_xyz_analysis_results.xlsx')
    pivot_data.to_excel(output_file, index=True)

    result_label.config(text=f"Файл успешно сохранён: {output_file}")

# Создание основного окна
root = tk.Tk()
root.title("ABC/XYZ Анализ")

# 1. Кнопка для выбора входного файла
input_file_button = tk.Button(root, text="Выбрать файл данных", command=choose_input_file)
input_file_button.pack()

# Поле для отображения пути к файлу
input_file_entry = tk.Entry(root, width=50)
input_file_entry.pack()

# 2. Кнопка для выбора пути сохранения
output_path_button = tk.Button(root, text="Выбрать папку для сохранения", command=choose_output_path)
output_path_button.pack()

# Поле для отображения пути сохранения
output_path_entry = tk.Entry(root, width=50)
output_path_entry.pack()

# 3. Кнопка для запуска анализа
run_button = tk.Button(root, text="Запустить анализ", command=run_analysis)
run_button.pack()

# Метка для отображения результатов
result_label = tk.Label(root, text="")
result_label.pack()

# Запуск приложения
root.mainloop()