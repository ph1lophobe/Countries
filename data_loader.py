import pandas as pd

def get_city_data():
    # Загружаем твой CSV (пропускаем технические строки, если нужно)
    df = pd.read_csv('Country.xlsx - Лист1.csv', skiprows=4)
    # Здесь можно добавить логику очистки данных специально под твои колонки
    return df

def calculate_savings_path(monthly_save, growth_rate, years):
    k = 1 + growth_rate / 100
    current_yearly_save = monthly_save * 12
    total = 0
    path = []
    for y in range(1, years + 1):
        total += current_yearly_save
        path.append({"Year": y, "Saved_This_Year": current_yearly_save, "Total": total})
        current_yearly_save *= k
    return pd.DataFrame(path)