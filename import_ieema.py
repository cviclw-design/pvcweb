import pandas as pd
import mysql.connector
from sqlalchemy import create_engine

MYSQL_CONFIG = {
    'host': 'localhost', 'user': 'root', 'password': '', 'database': 'pvcdb'
}

# Adjust columns to match your ieema.xlsx
df = pd.read_excel('ieema.xlsx')  # Put ieema.xlsx in E:\PVC-WebApp\
df.columns = df.columns.str.strip().str.lower()  # Clean headers

# Map to table (adjust as needed)
df = df.rename(columns={
    'base date': 'base_date', 'wpi': 'wpi', 'copper': 'copper', 
    'crgo': 'crgo', 'labour': 'labour', 'freight': 'freight',
    'other materials': 'other_materials'
})

engine = create_engine(f"mysql+mysqlconnector://root:@localhost/pvcdb")
df.to_sql('ieema_data', con=engine, if_exists='append', index=False, dtype={
    'wpi': 'DECIMAL(10,4)', 'copper': 'DECIMAL(10,4)', 'crgo': 'DECIMAL(10,4)',
    'labour': 'DECIMAL(10,4)', 'freight': 'DECIMAL(10,4)', 'other_materials': 'DECIMAL(10,4)'
})

print(f"✅ Imported {len(df)} IEEMA rows to pvcdb.ieema_data")
print(df.head())