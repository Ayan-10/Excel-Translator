import pandas as pd
from deep_translator import GoogleTranslator
import concurrent.futures
import time

# Function to translate a specific text
def translate_text(text):
    if pd.isna(text):
        return ""
    
    if isinstance(text, (str, bytes, bytearray)):
        try:
            translation = GoogleTranslator(source='auto', target='en').translate(text)
            return translation
        except Exception as e:
            print('Please wait..')
    
    return text  

# Function to translate a specific row in parallel
def translate_row(row):
    index, data = row
    return (index, data.apply(translate_text) if isinstance(data, pd.Series) else translate_text(data))

# parallel processing
def translate_excel_parallel(input_file_path, output_file_path):
    start_time = time.time()

    df = pd.read_excel(input_file_path, header=None)

    index, translated_header = translate_row((0, df.iloc[0]))
    df.iloc[0] = translated_header

    with concurrent.futures.ThreadPoolExecutor() as executor:
        translated_rows = list(executor.map(translate_row, df.iloc[1:].iterrows()))

    for index, translated_row in translated_rows:
        df.loc[index] = translated_row

    df.to_excel(output_file_path, index=False, header=False)

    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Translation and save completed in {elapsed_time:.2f} seconds.")

input_excel_file_path = 'Order Export.xls'
output_excel_file_path = 'Translated_Order_Export.xls'

print("Translation started...Please stand by....")
translate_excel_parallel(input_excel_file_path, output_excel_file_path)
