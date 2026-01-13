import pandas as pd
import numpy as np
from datetime import datetime


class SalesAnalyzer:
    """Класс для выполнения различных анализов данных о продажах."""

    def __init__(self, data):
        self.data = data

    def analyze_sales_over_time(self):
        """Анализ динамики продаж по месяцам."""
        if 'Sale_Date' not in self.data.columns:
            return pd.DataFrame()

        df = self.data.copy()
        df['Month_Year'] = df['Sale_Date'].dt.strftime('%Y-%m')

        # Группируем по месяцам и считаем выручку, количество продаж и прибыль
        monthly_sales = df.groupby('Month_Year').agg({
            'Sales_Amount': ['sum', 'count'],
            'Profit': 'sum' if 'Profit' in df.columns else None
        }).round(2)

        # Переименовываем колонки для удобства
        monthly_sales.columns = ['Общая_выручка', 'Количество_продаж', 'Прибыль']

        return monthly_sales

    def analyze_by_category(self):
        """Анализ продаж по категориям товаров."""
        if 'Product_Category' not in self.data.columns:
            return pd.DataFrame()

        category_stats = self.data.groupby('Product_Category').agg({
            'Sales_Amount': ['sum', 'count', 'mean'],
            'Profit': 'sum' if 'Profit' in self.data.columns else None,
            'Quantity_Sold': 'sum'
        }).round(2)

        category_stats.columns = ['Общая_выручка', 'Количество_продаж',
                                 'Средний_чек', 'Прибыль', 'Количество_товаров']

        # Добавляем долю рынка для каждой категории
        total_sales = category_stats['Общая_выручка'].sum()
        category_stats['Доля_рынка_%'] = (category_stats['Общая_выручка'] / total_sales * 100).round(2)

        return category_stats.sort_values('Общая_выручка', ascending=False)

    def analyze_by_region(self):
        """Анализ продаж по регионам."""
        if 'Region' not in self.data.columns:
            return pd.DataFrame()

        region_stats = self.data.groupby('Region').agg({
            'Sales_Amount': ['sum', 'count', 'mean'],
            'Profit': 'sum' if 'Profit' in self.data.columns else None
        }).round(2)

        region_stats.columns = ['Общая_выручка', 'Количество_продаж',
                               'Средний_чек', 'Прибыль']

        return region_stats.sort_values('Общая_выручка', ascending=False)

    def analyze_sales_reps(self):
        """Анализ эффективности работы продавцов."""
        if 'Sales_Rep' not in self.data.columns:
            return pd.DataFrame()

        rep_stats = self.data.groupby('Sales_Rep').agg({
            'Sales_Amount': ['sum', 'count', 'mean'],
            'Profit': 'sum' if 'Profit' in self.data.columns else None
        }).round(2)

        rep_stats.columns = ['Общая_выручка', 'Количество_продаж',
                            'Средний_чек', 'Прибыль']

        # Добавляем показатель эффективности (средняя выручка на одну сделку)
        if 'Количество_продаж' in rep_stats.columns and rep_stats['Количество_продаж'].sum() > 0:
            rep_stats['Эффективность'] = (rep_stats['Общая_выручка'] / rep_stats['Количество_продаж']).round(2)

        return rep_stats.sort_values('Общая_выручка', ascending=False)

    def analyze_by_customer_type(self):
        """Анализ продаж по типам клиентов (новые/постоянные)."""
        if 'Customer_Type' not in self.data.columns:
            return pd.DataFrame()

        customer_stats = self.data.groupby('Customer_Type').agg({
            'Sales_Amount': ['sum', 'count', 'mean'],
            'Profit': 'sum' if 'Profit' in self.data.columns else None
        }).round(2)

        customer_stats.columns = ['Общая_выручка', 'Количество_продаж',
                                 'Средний_чек', 'Прибыль']

        return customer_stats

    def get_top_products(self, n=10):
        """Возвращает топ товаров по выручке."""
        if 'Product_ID' not in self.data.columns:
            return pd.DataFrame()

        product_stats = self.data.groupby('Product_ID').agg({
            'Sales_Amount': 'sum',
            'Quantity_Sold': 'sum',
            'Profit': 'sum' if 'Profit' in self.data.columns else None
        }).round(2)

        product_stats.columns = ['Общая_выручка', 'Количество_продаж', 'Прибыль']

        return product_stats.sort_values('Общая_выручка', ascending=False).head(n)

    def get_comprehensive_report(self):
        """Собирает все виды анализов в один словарь."""
        report = {}

        report['Временной_анализ'] = self.analyze_sales_over_time()
        report['Анализ_по_категориям'] = self.analyze_by_category()
        report['Анализ_по_регионам'] = self.analyze_by_region()
        report['Анализ_продавцов'] = self.analyze_sales_reps()
        report['Топ_продукты'] = self.get_top_products()

        return report