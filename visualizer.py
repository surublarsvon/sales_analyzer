import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import os


class DataVisualizer:
    """Класс для создания графиков и визуализаций на основе результатов анализа."""

    def __init__(self):
        # Выбираем стиль графиков (предпочитаем seaborn, если он доступен)
        available_styles = plt.style.available
        if 'seaborn-v0_8' in available_styles:
            plt.style.use('seaborn-v0_8')
        elif 'seaborn' in available_styles:
            plt.style.use('seaborn')
        else:
            plt.style.use('ggplot')  # Альтернативный стиль

        # Палитра цветов для графиков
        self.colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728',
                       '#9467bd', '#8c564b', '#e377c2', '#7f7f7f']

    def plot_sales_trend(self, time_data, save_path='sales_trend.png'):
        """Строит график тренда продаж по месяцам."""
        if time_data.empty:
            return

        fig, ax = plt.subplots(figsize=(12, 6))

        x = time_data.index
        y = time_data['Общая_выручка']

        # Линия тренда с маркерами и заливкой под ней
        ax.plot(x, y, marker='o', linewidth=2, markersize=5, color=self.colors[0])
        ax.fill_between(x, y, alpha=0.3, color=self.colors[0])

        ax.set_title('Тренд продаж по месяцам', fontsize=14, fontweight='bold')
        ax.set_xlabel('Месяц', fontsize=12)
        ax.set_ylabel('Выручка ($)', fontsize=12)
        ax.grid(True, alpha=0.3)

        plt.xticks(rotation=45)
        plt.tight_layout()

        if save_path:
            plt.savefig(save_path, dpi=100, bbox_inches='tight')

        return fig

    def plot_category_sales(self, category_data, save_path='category_sales.png'):
        """Создаёт столбчатую диаграмму выручки по категориям."""
        if category_data.empty:
            return

        fig, ax = plt.subplots(figsize=(10, 6))

        categories = category_data.index
        sales = category_data['Общая_выручка']

        bars = ax.bar(categories, sales, color=self.colors[:len(categories)])

        ax.set_title('Выручка по категориям', fontsize=14, fontweight='bold')
        ax.set_xlabel('Категория', fontsize=12)
        ax.set_ylabel('Выручка ($)', fontsize=12)
        ax.grid(True, alpha=0.3, axis='y')

        # Подписываем значения над столбцами
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2., height,
                    f'${height:,.0f}', ha='center', va='bottom', fontsize=9)

        plt.xticks(rotation=45)
        plt.tight_layout()

        if save_path:
            plt.savefig(save_path, dpi=100, bbox_inches='tight')

        return fig

    def plot_regions_pie(self, region_data, save_path='regions_pie.png'):
        """Создаёт круговую диаграмму доли регионов в продажах."""
        if region_data.empty:
            return

        fig, ax = plt.subplots(figsize=(8, 8))

        regions = region_data.index
        sales = region_data['Общая_выручка']

        # Круговая диаграмма с процентами
        wedges, texts, autotexts = ax.pie(sales, labels=regions, autopct='%1.1f%%',
                                          colors=self.colors[:len(regions)],
                                          startangle=90)

        ax.set_title('Доля регионов в продажах', fontsize=14, fontweight='bold')

        # Делаем подписи с процентами белыми и жирными
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')

        plt.tight_layout()

        if save_path:
            plt.savefig(save_path, dpi=100, bbox_inches='tight')

        return fig

    def plot_top_sellers(self, rep_data, save_path='top_sellers.png'):
        """Горизонтальная столбчатая диаграмма топ-5 продавцов."""
        if rep_data.empty:
            return

        # Берем только первых 5 продавцов
        top_5 = rep_data.head(5)

        fig, ax = plt.subplots(figsize=(10, 6))

        sellers = top_5.index
        sales = top_5['Общая_выручка']

        bars = ax.barh(sellers, sales, color=self.colors[:len(sellers)])

        ax.set_title('Топ-5 продавцов', fontsize=14, fontweight='bold')
        ax.set_xlabel('Выручка ($)', fontsize=12)
        ax.grid(True, alpha=0.3, axis='x')

        # Подписываем значения справа от столбцов
        for i, bar in enumerate(bars):
            width = bar.get_width()
            ax.text(width + width * 0.01, bar.get_y() + bar.get_height() / 2,
                    f'${width:,.0f}', va='center', fontsize=9)

        plt.tight_layout()

        if save_path:
            plt.savefig(save_path, dpi=100, bbox_inches='tight')

        return fig

    def create_dashboard(self, analyses, save_dir='plots'):
        """Создаёт все графики и сохраняет их в указанную папку."""
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        figures = []

        # 1. Тренд продаж
        if 'Временной_анализ' in analyses and not analyses['Временной_анализ'].empty:
            fig = self.plot_sales_trend(analyses['Временной_анализ'],
                                        os.path.join(save_dir, 'sales_trend.png'))
            if fig:
                figures.append(fig)
                plt.close(fig)

        # 2. Продажи по категориям
        if 'Анализ_по_категориям' in analyses and not analyses['Анализ_по_категориям'].empty:
            fig = self.plot_category_sales(analyses['Анализ_по_категориям'],
                                           os.path.join(save_dir, 'category_sales.png'))
            if fig:
                figures.append(fig)
                plt.close(fig)

        # 3. Регионы
        if 'Анализ_по_регионам' in analyses and not analyses['Анализ_по_регионам'].empty:
            fig = self.plot_regions_pie(analyses['Анализ_по_регионам'],
                                        os.path.join(save_dir, 'regions_pie.png'))
            if fig:
                figures.append(fig)
                plt.close(fig)

        # 4. Продавцы
        if 'Анализ_продавцов' in analyses and not analyses['Анализ_продавцов'].empty:
            fig = self.plot_top_sellers(analyses['Анализ_продавцов'],
                                        os.path.join(save_dir, 'top_sellers.png'))
            if fig:
                figures.append(fig)
                plt.close(fig)

        print(f"✅ Графики сохранены в папке: {save_dir}")

        return figures

    def show_all(self):
        """Отображает все созданные графики (если нужно посмотреть в интерактивном режиме)."""
        plt.show()