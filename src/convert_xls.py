import os
import pandas as pd

def main():
    xls_files = os.listdir("../resources/xls")
    for file in xls_files:
        df = pd.read_excel(f"../resources/xls/{file}", skiprows=1)
        df.to_excel(f"../resources/faturas/{file.replace("xls", "xlsx")}", index=False, engine='openpyxl')

if __name__ == "__main__":
    main()