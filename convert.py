import pandas as pd
import argparse
import json
import os

def filter_excel(file_path, month, year,config):
    # Загрузка файла Excel
    xls = pd.ExcelFile(file_path)

    # Получение названия второго листа и его данных
    sheet_name = xls.sheet_names[1]
    df = pd.read_excel(file_path, sheet_name=sheet_name,header=1)

    # Добавление столбца "Currency"
    currency = sheet_name.split('-')[-1]
    df['Currency'] = currency

    # Фильтрация по дате
    start_date = f"{year}-{month}-01"
    end_date = f"{year}-{month}-{pd.Period(start_date).days_in_month}"
    df['Date'] = pd.to_datetime(df['Date'])
    df = df[df['Date'].between(start_date, end_date)]
    df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')


    params = config[currency]
    for col,vals in params.get("exceptions").items():
        for val in vals:
            if df[col].dtype == "object":
                df = df[(df[col] != val) & ~(df[col].str.startswith(val)) ]
            else:
                df = df[(df[col] != val)]
            if val is None:
                df = df.dropna(subset=[col])
    df['Description'] = df['Description'] + '\n' + df['Additional Information']

    # Выборка определенных полей
    df = df[['Date','Currency', 'Paid Out', 'Description']]
    df = df.astype(str)

    return df,currency


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('month',type=str)
    args = parser.parse_args()
    if args.month:
        month,year = args.month.split(".")
        print(month,year)
    with open('excel_budget/config.json','r') as f:
        config = json.load(f)

    data_path = config.get("data_path")

    files = os.listdir(os.path.join(data_path,f"{month}.{year}"))
    files = [file for file in files 
             if file.startswith("account_statement_") and file.endswith(".xlsx")]
    
    for file in files:
        file_path = os.path.join(data_path,f"{month}.{year}",file)
        filtered_data,currency = filter_excel(file_path, month, year,config)
        print(currency)
        if not os.path.exists(os.path.join(data_path,f"{month}.{year}","result")):
            os.mkdir(os.path.join(data_path,f"{month}.{year}","result"))
        filtered_data.to_excel(os.path.join(data_path,f"{month}.{year}","result",f"{currency}.xlsx"),
                               engine='openpyxl',
                               index=False)
    
    files = os.listdir(os.path.join(data_path,f"{month}.{year}","result"))
    files = [
                file for file in files
                if not (file.startswith("union") or file.startswith("~$"))
                    and file.endswith(".xlsx")
            ]
    # Считывание каждого файла в DataFrame и сохранение в список
    dataframes = [pd.read_excel(os.path.join(data_path,f"{month}.{year}","result",file)) 
                  for file in files]

    # Объединение всех DataFrame из списка в один DataFrame
    combined_df = pd.concat(dataframes, ignore_index=True)
    combined_df['Date'] = pd.to_datetime(combined_df['Date'])
    combined_df['Date'] = combined_df['Date'].dt.strftime('%d.%m.%Y')
    if os.path.exists(os.path.join(data_path,f"{month}.{year}","result","union.xlsx")):
        os.remove(os.path.join(data_path,f"{month}.{year}","result","union.xlsx"))
    combined_df.to_excel(os.path.join(data_path,f"{month}.{year}","result","union.xlsx"),
                               engine='openpyxl',
                               index=False)
if __name__ == "__main__":
    main()