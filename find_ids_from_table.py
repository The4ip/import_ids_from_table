import pymysql
import pandas as pd
import msoffcrypto
import io
import tkinter as tk
import subprocess
from tqdm import tqdm
from time import sleep


config = {
    'user': 'localhost',
    'password': 'password',
    'host': 'localhost',
    'database': 'db',
    'cursorclass': pymysql.cursors.DictCursor,
    'charset': 'utf8mb4'
}

#optional deletin last char from cell 
def remove_trailing_zero(id_request):
    if isinstance(id_request, int):
        id_str = str(id_request)
        if id_str[-1] == '0':
            return int(id_str[:-1]) if len(id_str) > 1 else 0
    return id_request


def get_request(id_request):
    try:
        cnx = pymysql.connect(**config)
        cur = cnx.cursor()
        query = "SELECT somethin FROM `request` WHERE `id`=%s;" #your query is here
        cur.execute(query, (id_request,))
        fullname = cur.fetchall()
        cur.close()
        cnx.close()
        return fullname
    except pymysql.MySQLError as e:
        print(f"Error connectin database: {e}")#error code
        return None

def read_column(filepath, column_name, output_file,password):
   #if file hass possword we try to use pwd 
    try:
        with open(filepath, 'rb') as f:
            encrypted = msoffcrypto.OfficeFile(f)
            encrypted.load_key(password=password)
            
            decrypted = io.BytesIO()
            encrypted.decrypt(decrypted)
            decrypted.seek(0)
        # Read Excel in DataFrame
        df = pd.read_excel(decrypted)
        
        # Chekin existing co,umn 
        if column_name in df.columns:
            # Convert column values to rows, remove dots and convert to integers
            def convert(value):
                if pd.notna(value):
                    try:
                        return int(str(value).replace('.', ''))
                    except ValueError:
                        return None
                return None
            
            df[column_name] = df[column_name].apply(convert)
            # Getting a list of unique IDs
            id_list = df[column_name].dropna().astype(int).tolist()
            
            with open(output_file, 'w') as file:
                # ЗWe request information for each id and write it to a file
                for id_request in tqdm(id_list, desc="Processing of applications", unit="application"): #progress bar
                    id_request = remove_trailing_zero(id_request)
                    #print(id_request)
                    result = get_request(id_request)
                    if result:
                        # Extracting the somethin value from the result
                        somethin = result[0]['somethin'] if result else 'No data available'
                        file.write(f"{somethin}\n")  
                        index = df[df[column_name] == id_request].index
                        if not index.empty:
                            df.at[index[0], outpt_column] = somethin
                        #print(f"{somethin}")  # For debugging
                    else:
                        file.write("No data\n")
                        print("No data")  # For debugging
                
            print(f"The results are saved in'{output_file}'")
        else:
            print(f"Error: Column '{column_name}' was not found.")
    except Exception as e:
        print(f"Error:: {e}")


if __name__ == "__main__":
    filepath = '//Path/to/your/file.xlsb'
    column_name = "Column name" 
    outpt_column= "outpt_column"
    output_file = 'results.txt'
    password = "password"#if u wanna u create input dialog to enter the pwd
    read_column(filepath, column_name, output_file,password)
    subprocess.Popen('notepad.exe ' + output_file)




