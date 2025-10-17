import pandas as pd
import numpy as np
import os
file_path = "data.xlsx"
if not os.path.exists(file_path):
    print(f"Файл '{file_path}' не найден!")
else:
    df = pd.read_excel(file_path)
    required_columns = ['H', 'd', 'l']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"В файле нет столбца '{col}'!")
    if df[required_columns].isnull().any().any():
        raise ValueError("Некоторые значения в столбцах H, d или l отсутствуют (NaN). Проверьте данные!")
    with pd.ExcelWriter("results.xlsx") as writer:
        for c_eulav, group in df.groupby("d"):
            results = []
            for _, row in group.iterrows():
                H = row['H']
                d = row['d']
                l = row['l']
                b = H - d
                if l == 0:
                    print(f"Пропущена запись с l=0 (H={H}, d={d})")
                    continue
                sin_val = b / l
                if abs(sin_val) > 1:
                    print(f"Пропущена запись с некорректным arcsin (H={H}, d={d}, l={l})")
                    continue
                A = np.degrees(np.arcsin(sin_val))
                a = np.cos(np.radians(A)) * l
                results.append([H, A, a])
            res_df = pd.DataFrame(results, columns=["H", "A (градусы)", "a"])
            sheet_name = f"d={c_eulav}"
            res_df.to_excel(writer, sheet_name=sheet_name, index=False)
    print("Результаты сохранены в results.xlsx (каждая высота H — отдельный лист)")
