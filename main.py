import pandas as pd
import os
import sys
from datetime import datetime

from data_loader import DataLoader
from analyzer import SalesAnalyzer
from visualizer import DataVisualizer


class SalesAnalysisSystem:
    """Главная система для анализа данных о продажах.
       Координирует загрузку, очистку, анализ, визуализацию и экспорт."""

    def __init__(self):
        # Инициализируем компоненты системы
        self.loader = DataLoader()
        self.analyzer = None
        self.visualizer = DataVisualizer()
        self.data = None

    def run(self, file_path=None):
        """Основной метод запуска анализа. Выполняет все шаги по порядку."""
        print("=" * 50)
        print("СИСТЕМА АНАЛИЗА ПРОДАЖ")
        print("=" * 50)

        # Если путь к файлу не передан, запросим у пользователя
        if not file_path:
            file_path = self.get_file_path()

        if not file_path:
            print("Файл не выбран. Завершение работы.")
            return

        # Шаг 1: Загружаем данные из CSV
        print("\n1. ЗАГРУЗКА ДАННЫХ")
        print("-" * 30)

        raw_data = self.loader.load_csv(file_path)
        if raw_data is None:
            print("Не удалось загрузить данные")
            return

        # Шаг 2: Очищаем и подготавливаем данные
        print("\n2. ОЧИСТКА ДАННЫХ")
        print("-" * 30)

        self.data = self.loader.clean_data()
        if self.data is None or self.data.empty:
            print("Нет данных после очистки")
            return

        # Показываем краткую сводку о данных
        summary = self.loader.get_summary(self.data)
        print("\nСВОДКА ПО ДАННЫМ:")
        for key, value in summary.items():
            print(f"  {key}: {value}")

        # Шаг 3: Проводим анализ данных
        print("\n3. АНАЛИЗ ДАННЫХ")
        print("-" * 30)

        self.analyzer = SalesAnalyzer(self.data)

        print("\nАнализ по категориям:")
        category_analysis = self.analyzer.analyze_by_category()
        if not category_analysis.empty:
            print(category_analysis.head())

        print("\nАнализ по регионам:")
        region_analysis = self.analyzer.analyze_by_region()
        if not region_analysis.empty:
            print(region_analysis.head())

        print("\nАнализ продавцов:")
        rep_analysis = self.analyzer.analyze_sales_reps()
        if not rep_analysis.empty:
            print(rep_analysis.head())

        print("\nАнализ по времени:")
        time_analysis = self.analyzer.analyze_sales_over_time()
        if not time_analysis.empty:
            print(time_analysis.head())

        # Шаг 4: Создаём графики и визуализации
        print("\n4. ВИЗУАЛИЗАЦИЯ")
        print("-" * 30)

        comprehensive_report = self.analyzer.get_comprehensive_report()
        self.visualizer.create_dashboard(comprehensive_report, 'analysis_results')

        # Шаг 5: Экспортируем результаты в файлы
        print("\n5. ЭКСПОРТ РЕЗУЛЬТАТОВ")
        print("-" * 30)

        self.export_results(comprehensive_report)

        print("\n" + "=" * 50)
        print("АНАЛИЗ ЗАВЕРШЕН УСПЕШНО!")
        print("=" * 50)

    def get_file_path(self):
        """Предлагаем пользователю выбрать CSV-файл из текущей папки или ввести путь вручную."""
        # Ищем все CSV-файлы в текущей директории
        csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]

        if csv_files:
            print("\nНайденные CSV файлы:")
            for i, file in enumerate(csv_files, 1):
                print(f"  {i}. {file}")

            choice = input("\nВыберите файл (номер) или введите свой путь: ").strip()

            # Если введён номер, возвращаем соответствующий файл
            if choice.isdigit() and 1 <= int(choice) <= len(csv_files):
                return csv_files[int(choice) - 1]

        # Запрашиваем путь вручную, если файлы не найдены или пользователь ввёл свой путь
        file_path = input("\nВведите путь к CSV файлу: ").strip()

        if not file_path:
            print("Путь не указан")
            return None

        if not os.path.exists(file_path):
            print(f"Файл не найден: {file_path}")
            return None

        return file_path

    def export_results(self, analyses):
        """Сохраняем результаты анализа в Excel-файл (если установлен openpyxl), иначе в CSV."""
        # Добавляем временную метку к имени файла
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        try:
            # Пробуем сохранить в Excel
            try:
                excel_file = f'sales_analysis_{timestamp}.xlsx'
                with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                    # Сохраняем очищенные данные
                    if self.data is not None:
                        self.data.to_excel(writer, sheet_name='Очищенные_данные', index=False)

                    # Сохраняем каждый результат анализа на отдельный лист
                    for sheet_name, data in analyses.items():
                        if not data.empty:
                            sheet_name_short = sheet_name[:31]  # Excel не поддерживает длинные имена листов
                            data.to_excel(writer, sheet_name=sheet_name_short)

                    # Создаём сводный лист с ключевыми показателями
                    summary_data = []

                    if 'Анализ_по_категориям' in analyses:
                        top_category = analyses['Анализ_по_категориям'].index[0]
                        top_sales = analyses['Анализ_по_категориям'].iloc[0]['Общая_выручка']
                        summary_data.append(['Лучшая категория', top_category, f"${top_sales:,.2f}"])

                    if 'Анализ_по_регионам' in analyses:
                        top_region = analyses['Анализ_по_регионам'].index[0]
                        region_sales = analyses['Анализ_по_регионам'].iloc[0]['Общая_выручка']
                        summary_data.append(['Лучший регион', top_region, f"${region_sales:,.2f}"])

                    if 'Анализ_продавцов' in analyses:
                        top_rep = analyses['Анализ_продавцов'].index[0]
                        rep_sales = analyses['Анализ_продавцов'].iloc[0]['Общая_выручка']
                        summary_data.append(['Лучший продавец', top_rep, f"${rep_sales:,.2f}"])

                    if self.data is not None:
                        total_sales = self.data['Sales_Amount'].sum()
                        total_profit = self.data['Profit'].sum() if 'Profit' in self.data.columns else 0
                        summary_data.append(['Общая выручка', '', f"${total_sales:,.2f}"])
                        summary_data.append(['Общая прибыль', '', f"${total_profit:,.2f}"])
                        summary_data.append(['Всего транзакций', len(self.data), ''])

                    summary_df = pd.DataFrame(summary_data, columns=['Показатель', 'Значение', 'Сумма'])
                    summary_df.to_excel(writer, sheet_name='Сводка', index=False)

                print(f"Excel отчет сохранен: {excel_file}")

            except ImportError:
                # Если openpyxl не установлен, сохраняем в CSV
                print("Модуль openpyxl не установлен. Сохраняю в CSV...")
                self.export_to_csv(analyses, timestamp)

        except Exception as e:
            # Если что-то пошло не так с Excel, переключаемся на CSV
            print(f"Не удалось создать Excel файл: {e}")
            print("Сохраняю в CSV...")
            self.export_to_csv(analyses, timestamp)

    def export_to_csv(self, analyses, timestamp):
        """Сохраняем каждый датафрейм в отдельный CSV-файл."""
        # Сохраняем очищенные данные
        if self.data is not None:
            cleaned_file = f'cleaned_data_{timestamp}.csv'
            self.data.to_csv(cleaned_file, index=False, encoding='utf-8')
            print(f"Очищенные данные: {cleaned_file}")

        # Сохраняем каждый анализ
        for sheet_name, data in analyses.items():
            if not data.empty:
                csv_file = f'{sheet_name}_{timestamp}.csv'
                data.to_csv(csv_file, encoding='utf-8')
                print(f"{sheet_name}: {csv_file}")

        # Создаем текстовый сводный отчёт
        summary_file = f'summary_{timestamp}.txt'
        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write("=" * 50 + "\n")
            f.write("СВОДНЫЙ ОТЧЕТ ПО АНАЛИЗУ ПРОДАЖ\n")
            f.write("=" * 50 + "\n\n")

            if self.data is not None:
                total_sales = self.data['Sales_Amount'].sum()
                total_profit = self.data['Profit'].sum() if 'Profit' in self.data.columns else 0

                f.write(f"Общая выручка: ${total_sales:,.2f}\n")
                f.write(f"Общая прибыль: ${total_profit:,.2f}\n")
                f.write(f"Всего транзакций: {len(self.data)}\n")
                f.write(f"Дата анализа: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

        print(f"Сводный отчет: {summary_file}")

    def create_sample_data(self):
        """Генерируем тестовые данные для демонстрации работы системы."""
        print("СОЗДАНИЕ ТЕСТОВЫХ ДАННЫХ")
        print("-" * 30)

        import numpy as np

        np.random.seed(42)  # Для воспроизводимости
        n_records = 500

        # Создаём структуру тестовых данных
        data = {
            'Product_ID': [f'P{str(i).zfill(4)}' for i in range(1, n_records + 1)],
            'Sale_Date': pd.date_range('2023-01-01', periods=n_records, freq='D'),
            'Sales_Rep': np.random.choice(['Alice', 'Bob', 'Charlie', 'David', 'Eve'], n_records),
            'Region': np.random.choice(['North', 'South', 'East', 'West'], n_records),
            'Quantity_Sold': np.random.randint(1, 51, n_records),
            'Product_Category': np.random.choice(['Electronics', 'Furniture', 'Clothing', 'Food'], n_records),
            'Unit_Cost': np.random.uniform(50, 5000, n_records).round(2),
            'Customer_Type': np.random.choice(['New', 'Returning'], n_records),
            'Discount': np.random.uniform(0, 30, n_records).round(1),
            'Payment_Method': np.random.choice(['Credit Card', 'Cash', 'Bank Transfer'], n_records),
            'Sales_Channel': np.random.choice(['Online', 'Retail'], n_records),
        }

        # Рассчитываем производные поля
        data['Unit_Price'] = (data['Unit_Cost'] * np.random.uniform(1.1, 2.0, n_records)).round(2)
        data['Sales_Amount'] = (data['Unit_Price'] * data['Quantity_Sold'] *
                                (1 - data['Discount'] / 100)).round(2)
        data['Region_and_Sales_Rep'] = data['Region'] + '_' + data['Sales_Rep']

        df = pd.DataFrame(data)
        filename = 'sample_sales_data.csv'
        df.to_csv(filename, index=False, encoding='utf-8')

        print(f"   Тестовые данные созданы: {filename}")
        print(f"   Количество записей: {n_records}")
        print(f"   Общая выручка: ${df['Sales_Amount'].sum():,.2f}")

        return filename


def main():
    """Точка входа в программу. Предлагаем пользователю выбор: использовать существующий файл или создать тестовый."""
    system = SalesAnalysisSystem()

    print("\nДОСТУПНЫЕ ОПЦИИ:")
    print("  1. Использовать существующий файл")
    print("  2. Создать тестовые данные и проанализировать")

    choice = input("\nВыберите опцию (1 или 2): ").strip()

    if choice == '2':
        # Создаём тестовые данные
        file_path = system.create_sample_data()
        print(f"\nСоздан файл: {file_path}")

        # Предлагаем сразу проанализировать созданные данные
        analyze = input("Проанализировать созданные данные? (y/n): ").lower()
        if analyze == 'y':
            system.run(file_path)
    else:
        # Если запущен с аргументом командной строки, используем его как путь к файлу
        if len(sys.argv) > 1:
            file_path = sys.argv[1]
        else:
            file_path = None

        system.run(file_path)


if __name__ == "__main__":
    main()