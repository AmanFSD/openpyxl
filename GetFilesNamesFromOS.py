import os
import pandas as pd


print(os.listdir())
print(os.getcwd())
folder_path = input('Folder Path? >')
os.chdir(rf'{folder_path}')

print(os.getcwd())

files = os.listdir()

for file in files:
    print(file)