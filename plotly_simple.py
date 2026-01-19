
import plotly.graph_objects as go
import os


class SimplePlotlyVisualizer:
    """Пвизуализация данных с Plotly"""

    def create_simple_dashboard(self, analyses, data, save_dir='interactive_charts'):
        """Создаёт простую панель с Plotly графиками"""
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        figs = {}

        # 1. Линейный график тренда продаж
        if 'Временной_анализ' in analyses and not analyses['Временной_анализ'].empty:
            fig = self._create_trend_chart(analyses['Временной_анализ'])
            if fig:
                fig.write_html(f"{save_dir}/trend.html")
                figs['trend'] = fig

        # 2. Столбчатая диаграмма категорий
        if 'Анализ_по_категориям' in analyses and not analyses['Анализ_по_категориям'].empty:
            fig = self._create_bar_chart(analyses['Анализ_по_категориям'], 'Категории')
            if fig:
                fig.write_html(f"{save_dir}/categories.html")
                figs['categories'] = fig

        # 3. Круговая диаграмма регионов
        if 'Анализ_по_регионам' in analyses and not analyses['Анализ_по_регионам'].empty:
            fig = self._create_pie_chart(analyses['Анализ_по_регионам'], 'Регионы')
            if fig:
                fig.write_html(f"{save_dir}/regions.html")
                figs['regions'] = fig

        # 4. Горизонтальная диаграмма продавцов
        if 'Анализ_продавцов' in analyses and not analyses['Анализ_продавцов'].empty:
            fig = self._create_horizontal_bar(analyses['Анализ_продавцов'])
            if fig:
                fig.write_html(f"{save_dir}/reps.html")
                figs['reps'] = fig

        # Создаём сводную панель только если есть графики
        if figs:
            self._create_summary_dashboard(figs, save_dir)
            print(f" Plotly графики сохранены в папке: {save_dir}")
        else:
            print("  Нет данных для создания Plotly графиков")

        return figs

    def _create_trend_chart(self, time_data):
        """Линейный график тренда"""
        try:
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=time_data.index,
                y=time_data['Общая_выручка'],
                mode='lines+markers',
                name='Выручка',
                line=dict(color='blue', width=2)
            ))

            fig.update_layout(
                title='Тренд продаж',
                xaxis_title='Месяц',
                yaxis_title='Выручка ($)',
                template='plotly_white'
            )

            return fig
        except Exception as e:
            print(f"  Ошибка при создании графика тренда: {e}")
            return None

    def _create_bar_chart(self, data, title):
        """Столбчатая диаграмма"""
        try:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=data.index,
                y=data['Общая_выручка'],
                name='Выручка',
                marker_color='lightblue'
            ))

            fig.update_layout(
                title=f'Выручка по {title.lower()}',
                xaxis_title=title,
                yaxis_title='Выручка ($)',
                template='plotly_white'
            )

            return fig
        except Exception as e:
            print(f" Ошибка при создании столбчатой диаграммы: {e}")
            return None

    def _create_pie_chart(self, data, title):
        """Круговая диаграмма"""
        try:
            fig = go.Figure()
            fig.add_trace(go.Pie(
                labels=data.index,
                values=data['Общая_выручка'],
                name='Выручка'
            ))

            fig.update_layout(
                title=f'Доля {title.lower()} в продажах',
                template='plotly_white'
            )

            return fig
        except Exception as e:
            print(f" Ошибка при создании круговой диаграммы: {e}")
            return None

    def _create_horizontal_bar(self, data):
        """Горизонтальная столбчатая диаграмма"""
        try:
            # Берем топ-10
            top_data = data.head(10).sort_values('Общая_выручка')

            fig = go.Figure()
            fig.add_trace(go.Bar(
                y=top_data.index,
                x=top_data['Общая_выручка'],
                orientation='h',
                marker_color='green'
            ))

            fig.update_layout(
                title='Топ-10 продавцов',
                xaxis_title='Выручка ($)',
                yaxis_title='Продавец',
                template='plotly_white',
                height=500
            )

            return fig
        except Exception as e:
            print(f"  Ошибка при создании горизонтальной диаграммы: {e}")
            return None

    def _create_summary_dashboard(self, figs, save_dir):
        """Создаёт простую HTML панель"""
        # Словарь соответствия названий и отображаемых имен
        display_names = {
            'trend': 'Тренд продаж',
            'categories': 'Выручка по категориям',
            'regions': 'Продажи по регионам',
            'reps': 'Топ продавцов'
        }

        html = """<!DOCTYPE html>
<html>
<head>
    <title>Интерактивная панель анализа продаж</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                  color: white; padding: 20px; border-radius: 10px; margin-bottom: 20px; 
                  text-align: center; }
        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(600px, 1fr)); 
                gap: 20px; }
        .chart-container { background: white; padding: 15px; border-radius: 10px; 
                          box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .chart-title { color: #2c3e50; margin-top: 0; border-bottom: 2px solid #3498db; 
                      padding-bottom: 10px; }
        iframe { width: 100%; height: 500px; border: none; }
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 Интерактивная панель анализа продаж</h1>
        <p>Используйте интерактивные графики для глубокого анализа данных</p>
    </div>

    <div class="grid">
"""

        # Добавляем iframe для каждого графика
        for name, fig in figs.items():
            if name in display_names:
                html += f"""
        <div class="chart-container">
            <h2 class="chart-title">{display_names[name]}</h2>
            <iframe src="{name}.html"></iframe>
        </div>
"""

        html += """    </div>

    <div style="text-align: center; margin-top: 30px; color: #666;">
        <p>Создано с помощью Plotly | Система анализа продаж</p>
    </div>
</body>
</html>"""

        # Сохраняем HTML файл
        with open(f'{save_dir}/dashboard.html', 'w', encoding='utf-8') as f:
            f.write(html)