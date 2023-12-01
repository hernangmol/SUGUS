import pandas as pd
import mysql.connector
from tkinter import messagebox

def main():
    # Extracción datos de excel ########################
    df = pd.read_excel("fuente.xlsx")
    data = df.iloc[3]
    #print(data.iloc[11])
    #print(pd.isnull(data.iloc[11]))
    print(data.iloc[10])

  

if __name__ == '__main__':
    main()