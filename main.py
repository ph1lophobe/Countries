from fastapi import FastAPI
import pandas as pd
from data_loader import get_city_data, calculate_savings_path

app = FastAPI(title="Relocation 2031 API")

@app.get("/cities")
def list_cities():
    df = get_city_data()
    # Возвращаем список городов для выпадающего списка
    return {"cities": df['Country'].dropna().tolist()}

@app.get("/calculate")
def calculate(monthly_save: float, growth: float, years: int):
    df_path = calculate_savings_path(monthly_save, growth, years)
    return df_path.to_dict(orient="records")