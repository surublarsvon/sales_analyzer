import pandas as pd
import numpy as np
from datetime import datetime
import warnings
import os

warnings.filterwarnings('ignore')


class DataLoader:
    """Класс для загрузки и предварительной обработки данных из CSV-файлов."""

    def __init__(self):
        self.data = None

    def load_csv(self, file_path):
        """Пытаемся загрузить CSV-файл, перебирая возможные кодировки."""
        try:
            print(f"Загрузка файла: {file_path}")

            # Список возможных кодировок для русскоязычных данных
            encodings = ['utf-8', 'latin-1', 'cp1251', 'windows-1251']

            for encoding in encodings:
                try:
                    self.data = pd.read_csv(file_path, encoding=encoding)
                    print(f"Файл загружен (кодировка: {encoding})")

                    # Проверяем наличие обязательных колонок
                    required_cols = ['Product_ID', 'Sale_Date', 'Sales_Rep', 'Region',
                                    'Sales_Amount', 'Quantity_Sold', 'Product_Category']

                    missing = [col for col in required_cols if col not in self.data.columns]
                    if missing:
                        print(f"⚠️ Отсутствуют колонки: {missing}")
                        return None

                    return self.data
                except UnicodeDecodeError:
                    continue

            print("Не удалось загрузить файл")
            return None

        except Exception as e:
            print(f"Ошибка: {e}")
            return None

    def clean_data(self):
        """Очистка данных: обработка пропусков, преобразование типов, добавление новых полей."""
        if self.data is None:
            print("Нет данных для очистки")
            return None

        df = self.data.copy()

        # Преобразуем строки с датами в формат datetime
        if 'Sale_Date' in df.columns:
            df['Sale_Date'] = pd.to_datetime(df['Sale_Date'], errors='coerce')

        # Удаляем строки, где отсутствуют ключевые поля
        important_cols = ['Product_ID', 'Sale_Date', 'Sales_Amount']
        df = df.dropna(subset=[col for col in important_cols if col in df.columns])

        # Заполняем пропуски в текстовых полях значением "Неизвестно"
        text_cols = ['Sales_Rep', 'Region', 'Product_Category',
                    'Customer_Type', 'Payment_Method', 'Sales_Channel']

        for col in text_cols:
            if col in df.columns:
                df[col] = df[col].fillna('Неизвестно')

        # Заполняем пропуски в числовых полях медианным значением
        num_cols = ['Quantity_Sold', 'Unit_Cost', 'Unit_Price', 'Discount']

        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
                df[col] = df[col].fillna(df[col].median() if not df[col].isnull().all() else 0)

        # Добавляем вычисляемые поля: прибыль и маржу
        if all(col in df.columns for col in ['Sales_Amount', 'Unit_Cost', 'Quantity_Sold']):
            df['Profit'] = df['Sales_Amount'] - (df['Unit_Cost'] * df['Quantity_Sold'])

        if all(col in df.columns for col in ['Sales_Amount', 'Profit']):
            df['Profit_Margin'] = (df['Profit'] / df['Sales_Amount'] * 100).round(2)

        # Добавляем временные характеристики для анализа
        if 'Sale_Date' in df.columns:
            df['Year'] = df['Sale_Date'].dt.year
            df['Month'] = df['Sale_Date'].dt.month
            df['Day'] = df['Sale_Date'].dt.day
            df['Weekday'] = df['Sale_Date'].dt.day_name()

        print(f"Данные очищены. Сохранено {len(df)} строк")
        return df

    def get_summary(self, df):
        """Формируем краткую сводку о данных для вывода на экран."""
        summary = {
            'Всего строк': len(df),
            'Всего колонок': len(df.columns),
            'Общая выручка': f"${df['Sales_Amount'].sum():,.2f}",
            'Общая прибыль': f"${df['Profit'].sum():,.2f}" if 'Profit' in df.columns else "N/A",
            'Средний чек': f"${df['Sales_Amount'].mean():.2f}",
            'Количество транзакций': len(df)
        }

        # Добавляем информацию о периоде, регионах и категориях, если они есть
        if 'Sale_Date' in df.columns:
            summary['Период'] = f"{df['Sale_Date'].min().strftime('%d.%m.%Y')} - {df['Sale_Date'].max().strftime('%d.%m.%Y')}"

        if 'Region' in df.columns:
            summary['Регионов'] = df['Region'].nunique()

        if 'Product_Category' in df.columns:
            summary['Категорий'] = df['Product_Category'].nunique()

        return summary