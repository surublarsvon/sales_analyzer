import pandas as pd
import os
import sys
from datetime import datetime

from data_loader import DataLoader
from analyzer import SalesAnalyzer
from visualizer import DataVisualizer

# Проверяем наличие Plotly
try:
    from plotly_simple import SimplePlotlyVisualizer

    HAS_PLOTLY = True
    print(" Plotly доступен для интерактивной визуализации")
except ImportError:
    HAS_PLOTLY = False
    print("  Plotly не установлен. Используются только статические графики.")


class SalesAnalysisSystem:
    """Главная система для анализа данных о продажах."""

    def __init__(self):
        # Инициализируем компоненты системы
        self.loader = DataLoader()
        self.analyzer = None
        self.visualizer = DataVisualizer()

        # Добавляем Plotly визуализатор если доступен
        if HAS_PLOTLY:
            self.plotly_viz = SimplePlotlyVisualizer()

        self.data = None

    def run(self, file_path=None):
        """Основной метод запуска анализа."""
        print("=" * 50)
        print("СИСТЕМА АНАЛИЗА ПРОДАЖ")
        print("=" * 50)

        # Если путь к файлу не передан, запросим у пользователя
        if not file_path:
            file_path = self.get_file_path()

        if not file_path:
            print("Файл не выбран. Завершение работы.")
            return

        # Шаг 1: Загружаем данные
        print("\n1. ЗАГРУЗКА ДАННЫХ")
        print("-" * 30)

        raw_data = self.loader.load_csv(file_path)
        if raw_data is None:
            print("Не удалось загрузить данные")
            return

        # Шаг 2: Очищаем данные
        print("\n2. ОЧИСТКА ДАННЫХ")
        print("-" * 30)

        self.data = self.loader.clean_data()
        if self.data is None or self.data.empty:
            print("Нет данных после очистки")
            return

        # Показываем краткую сводку
        summary = self.loader.get_summary(self.data)
        print("\nСВОДКА ПО ДАННЫМ:")
        for key, value in summary.items():
            print(f"  {key}: {value}")

        # Шаг 3: Проводим анализ
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

        # Получаем все анализы для визуализации
        comprehensive_report = self.analyzer.get_comprehensive_report()

        # Создаём графики matplotlib
        print("\n4. СТАТИЧЕСКАЯ ВИЗУАЛИЗАЦИЯ (Matplotlib)")
        print("-" * 30)

        self.visualizer.create_dashboard(comprehensive_report, 'static_charts')

        # Создаём интерактивные графики Plotly (если доступно)
        if HAS_PLOTLY:
            print("\n4.1 ИНТЕРАКТИВНАЯ ВИЗУАЛИЗАЦИЯ (Plotly)")
            print("-" * 30)

            self.plotly_viz.create_simple_dashboard(
                comprehensive_report,
                self.data,
                'interactive_charts'
            )
            print("  Интерактивные графики сохранены в папке 'interactive_charts'")
            print("   Откройте 'interactive_charts/dashboard.html' в браузере")
        else:
            print("\n4.1 ИНТЕРАКТИВНАЯ ВИЗУАЛИЗАЦИЯ")
            print("-" * 30)
            print(" Plotly не установлен. Для интерактивных графиков выполните:")
            print("   pip install plotly")

        # Шаг 5: Экспортируем результаты
        print("\n5. ЭКСПОРТ РЕЗУЛЬТАТОВ")
        print("-" * 30)

        self.export_results(comprehensive_report)

        print("\n" + "=" * 50)
        print("АНАЛИЗ ЗАВЕРШЕН УСПЕШНО!")
        print("=" * 50)
        print("\nСозданные файлы:")
        print("  - Статические графики: папка 'static_charts'")
        if HAS_PLOTLY:
            print("  - Интерактивные графики: 'interactive_charts/dashboard.html'")
        print("  - Excel/CSV отчеты: в текущей папке")

    def get_file_path(self):
        """Предлагаем выбрать CSV-файл."""
        # Ищем CSV файлы в текущей директории
        csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]

        if csv_files:
            print("\nНайденные CSV файлы:")
            for i, file in enumerate(csv_files, 1):
                print(f"  {i}. {file}")

            choice = input("\nВыберите файл (номер) или введите свой путь: ").strip()

            # Если введён номер, возвращаем соответствующий файл
            if choice.isdigit() and 1 <= int(choice) <= len(csv_files):
                return csv_files[int(choice) - 1]

        # Запрашиваем путь вручную
        file_path = input("\nВведите путь к CSV файлу: ").strip()

        if not file_path:
            print("Путь не указан")
            return None

        if not os.path.exists(file_path):
            print(f"Файл не найден: {file_path}")
            return None

        return file_path

    def export_results(self, analyses):
        """Сохраняем результаты анализа в файлы."""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        try:
            # Пробуем сохранить в Excel
            excel_file = f'sales_analysis_{timestamp}.xlsx'
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                # Сохраняем очищенные данные
                if self.data is not None:
                    self.data.to_excel(writer, sheet_name='Очищенные_данные', index=False)

                # Сохраняем каждый анализ
                for sheet_name, data in analyses.items():
                    if not data.empty:
                        sheet_name_short = sheet_name[:31]
                        data.to_excel(writer, sheet_name=sheet_name_short)

                # Создаём сводный лист
                summary_data = []

                if 'Анализ_по_категориям' in analyses:
                    top_cat = analyses['Анализ_по_категориям'].index[0]
                    top_cat_sales = analyses['Анализ_по_категориям'].iloc[0]['Общая_выручка']
                    summary_data.append(['Лучшая категория', top_cat, f"${top_cat_sales:,.2f}"])

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
                    summary_data.append(['Общая выручка', '', f"${total_sales:,.2f}"])
                    summary_data.append(['Всего транзакций', len(self.data), ''])

                summary_df = pd.DataFrame(summary_data, columns=['Показатель', 'Значение', 'Сумма'])
                summary_df.to_excel(writer, sheet_name='Сводка', index=False)

            print(f" Excel отчет: {excel_file}")

        except Exception as e:
            # Если Excel не получился, сохраняем в CSV
            print(f"  Не удалось создать Excel: {e}")
            print("Сохраняю в CSV...")
            self._save_to_csv(analyses, timestamp)

    def _save_to_csv(self, analyses, timestamp):
        """Сохраняем в CSV файлы."""
        # Очищенные данные
        if self.data is not None:
            cleaned_file = f'cleaned_data_{timestamp}.csv'
            self.data.to_csv(cleaned_file, index=False, encoding='utf-8')
            print(f" Очищенные данные: {cleaned_file}")

        # Результаты анализа
        for name, data in analyses.items():
            if not data.empty:
                csv_file = f'{name}_{timestamp}.csv'
                data.to_csv(csv_file, encoding='utf-8')
                print(f"{name}: {csv_file}")


def main():
    """Точка входа в программу."""
    system = SalesAnalysisSystem()

    print("\n" + "=" * 50)
    print("СИСТЕМА АНАЛИЗА ПРОДАЖ v2.0")
    print("=" * 50)
    print(f"Интерактивные графики: {'ДОСТУПНЫ' if HAS_PLOTLY else 'НЕ ДОСТУПНЫ (pip install plotly)'}")

    print("\nДОСТУПНЫЕ ОПЦИИ:")
    print("  1. Использовать существующий файл")
    print("  2. Создать тестовые данные и проанализировать")

    choice = input("\nВыберите опцию (1 или 2): ").strip()

    if choice == '2':
        # Создаём тестовые данные
        print("\nСОЗДАНИЕ ТЕСТОВЫХ ДАННЫХ")
        print("-" * 30)

        import numpy as np
        np.random.seed(42)

        # Простой пример тестовых данных
        dates = pd.date_range('2023-01-01', periods=100, freq='D')
        data = {
            'Product_ID': [f'P{i:04d}' for i in range(1, 101)],
            'Sale_Date': dates,
            'Sales_Rep': np.random.choice(['Анна', 'Борис', 'Светлана', 'Дмитрий'], 100),
            'Region': np.random.choice(['Север', 'Юг', 'Восток', 'Запад'], 100),
            'Product_Category': np.random.choice(['Электроника', 'Одежда', 'Продукты', 'Мебель'], 100),
            'Quantity_Sold': np.random.randint(1, 50, 100),
            'Unit_Price': np.random.uniform(50, 500, 100).round(2),
        }

        data['Sales_Amount'] = (data['Quantity_Sold'] * data['Unit_Price']).round(2)
        data['Unit_Cost'] = (data['Unit_Price'] * np.random.uniform(0.6, 0.9, 100)).round(2)

        df = pd.DataFrame(data)
        filename = 'sample_data.csv'
        df.to_csv(filename, index=False, encoding='utf-8')

        print(f"  Создан файл: {filename}")
        print(f"   Записей: {len(df)}")
        print(f"   Общая выручка: ${df['Sales_Amount'].sum():,.2f}")

        # Анализируем созданные данные
        analyze = input("\nПроанализировать созданные данные? (y/n): ").lower()
        if analyze == 'y':
            system.run(filename)
    else:
        # Используем существующий файл
        if len(sys.argv) > 1:
            file_path = sys.argv[1]
        else:
            file_path = None

        system.run(file_path)


if __name__ == "__main__":
    main()
