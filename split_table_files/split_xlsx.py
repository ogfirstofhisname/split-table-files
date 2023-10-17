# split_csv.py
import pandas as pd
import os, sys

# get n from user input. convert to integer. n is max number of rows per file, not including header.
success = False
input_s = input('How many rows per file? Type a number and press Enter:')
while not success:
    try:
        n = int(input_s)
        success = True
    except ValueError:
        print('Error - not a valid number. Please type a number and press Enter:')
        input_s = input()

# # path to the dir of this python script
# path = os.path.dirname(os.path.realpath(__file__))

# determine if application is a script file or frozen exe
if getattr(sys, 'frozen', False):
    path = os.path.dirname(sys.executable)
elif __file__:
    path = os.path.dirname(__file__)
print(f'path: {path}')



# list all xlsx files in the dir, only if they have n+1 rows or more. Open each one with pandas to count the rows
files = [f for f in os.listdir(path) if f.endswith('.xlsx') and pd.read_excel(f'{path}/{f}').shape[0] > n]

# for each file, read it in and split it into multiple files, each with up to n rows and a header. Read and save as xlsx
for file in files:
    print(f'file: {file}')
    df = pd.read_excel(f'{path}/{file}')
    for i in range(0, len(df), n):
        df[i:i+n].to_excel(f'{path}/{file}_{i}.xlsx', index=False, header=True)
        print(f'{path}/{file}_{i}.xlsx, length = {len(df[i:i+n])}')
    # df = pd.read_csv(f'{path}/{file}')
    # for i in range(0, len(df), n):
    #     df[i:i+n].to_csv(f'{path}/{file}_{i}.csv', index=False)
    #     print(f'{path}/{file}_{i}.csv, length = {len(df[i:i+n])}')