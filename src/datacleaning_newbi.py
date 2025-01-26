import pandas as pd
import re



def read_excel_store_data(filename):
    
    df_raw = pd.read_excel(filename, 
                             sheet_name="Sheet", header = None)
    
    hours_col = pd.Series([
        int(i) if (i is not None and pd.notna(i)) else 0
        for i in df_raw.loc[0, :]
    ])
    
    return df_raw, hours_col





def extract_before_after_store_totals(df_raw, split = "before"):
    
    # Drop rows that are completely empty
    df_raw.dropna(how='all', inplace=True)
    
    #Find the row index that contains Store Totals
    store_totals_idx = df_raw.index[df_raw[0].astype(str)\
                                    .str.contains("Store Totals", na=False)]
    store_totals_row = store_totals_idx[0]
    
    if split == "before":
        #Split into “before” and “after” DataFrames
        df = df_raw.loc[:store_totals_row - 1]
    else:
        df  = df_raw.loc[store_totals_row + 1:]
        
    return df


#Extract the actual date (assuming format "Date: 1/5/2025")
def extract_date(text):
    match = re.search(r"Date:\s*(.*)", str(text))
    return match.group(1).strip() if match else None


def hour12_to_hour24(h12, am_pm):
    """Convert 12-hour integer + AM/PM => 24-hour integer."""
    if am_pm.upper() == "AM":
        return 0 if h12 == 12 else h12
    else:  # PM
        return 12 if h12 == 12 else h12 + 12


#Identify AM/PM rows
def get_am_pm(val):
    val_str = str(val).strip().upper()
    return val_str if val_str in ["AM", "PM"] else None



def col_name(h):
    """
    Given an hour like 0..23,
    return a string "HH.00-HH+1.00".
    Example: 0 -> "00.00-01.00", 12 -> "12.00-13.00"
    """
    h_end = (h + 1) % 24
    return f"{h:01d}.00-{h_end:01d}.00"



#Main Data Cleaning for Store Data
def clean_store_data(df_raw, hours_col, split = "before"):
    """ Takes a 'raw' subset DataFrame and performs 
        date-extraction, forward-fill, AM/PM extraction, etc.
        Returns a cleaned DataFrame.
    """
    
    df = extract_before_after_store_totals(df_raw, split)
    
    #Drop rows that are completely empty
    df = df.dropna(how='all').copy()
    
    #Identify date rows
    df['is_date_row'] = df[0].astype(str).str.startswith("Date:")
    
    #Extract Date
    df['ExtractedDate'] = df.apply(
        lambda row: extract_date(row[0]) if row['is_date_row'] else None, axis=1
    )
    
    
    #Forward-fill the date down
    df['ExtractedDate'] = df['ExtractedDate'].ffill()
    
    #Identify AM/PM rows
    df['TimeOfDay'] = df[0].apply(get_am_pm)
    
    #Keep only AM/PM rows (the actual data rows)
    df = df[df['TimeOfDay'].notnull()].copy()
    
    #Convert ExtractedDate to datetime (optional)
    df['ExtractedDate'] = pd.to_datetime(df['ExtractedDate'], errors='coerce')
    
    #Create a new column in the format: MonthName Day Year
    #e.g., "January 25 2025"
    df['FormattedDate'] = df['ExtractedDate'].dt.strftime('%B %d %Y')

    #Create column for the day of the week
    df['DayOfWeek'] = df['ExtractedDate'].dt.day_name()
    
    #Rename numeric columns (assuming columns 1,2,3,... have numeric data)
    df.rename(columns=hours_col , inplace=True)
    
    #Drop columns you don’t need (like the original text col)
    df.drop(columns=[0, 'is_date_row'], errors='ignore', inplace=True)
    
    #Converting into right format
    df = df.melt(
        id_vars=['ExtractedDate', 'TimeOfDay', 'FormattedDate', 'DayOfWeek'],
        var_name='Hour12',
        value_name='Value'
    )
    
    # Convert Hour12 to integer
    df["Hour12"] = df["Hour12"].astype(int)
    
    # Create a 24-hour column
    df["Hour24"] = df.apply(
        lambda row: hour12_to_hour24(row["Hour12"], row["TimeOfDay"]), 
        axis=1 )
    

    df = df.pivot(
        index=["FormattedDate", "DayOfWeek"], 
        columns="Hour24", 
        values="Value" )
    
    df = df.copy()
    df.columns = [col_name(h) for h in df.columns]
    df.reset_index(inplace=True)
    
    return df




if __name__ == '__main__':
    
    filename = './Data/20250125/Basket_Analysis_Report_Store Name.xlsx'
    df_raw, hours_col = \
        read_excel_store_data(filename)
    # Now apply this function to df_before and df_after
    cleaned_before = clean_store_data(df_raw, hours_col, split = "before")
    cleaned_after  = clean_store_data(df_raw, hours_col, split = "after")

    df = cleaned_after.copy()
    included_col = ["FormattedDate", "DayOfWeek",
                    "7.00-8.00", "8.00-9.00", "9.00-10.00", 
                    "10.00-11.00", "11.00-12.00",
                    "12.00-13.00", "13.00-14.00", "14.00-15.00", "15.00-16.00",
                    "16.00-17.00", "17.00-18.00", "18.00-19.00", 
                    "19.00-20.00", "20.00-21.00"]
    
    if len(df) > 0:
        unclean_col = df.columns.tolist()
        new_col = []
        c_count = 0
        for c in unclean_col:
            c_new = str(c).lstrip(" ").rstrip(" ").lstrip("\xa0")
            new_col.append(c_new)
            c_count += 1
        df.columns = new_col
        
        # Check if the required columns exist"
        total_columns = []
        for col in included_col:
            if col not in df.columns:
                total_columns.append(col)
    
        total_columns_string = ", ".join(total_columns)
        if len(total_columns) > 1:
            
            string_error = f"Error1: The Excel file must contain {total_columns_string} columns."
            print(string_error)
        
        elif len(total_columns) == 1:
            string_error = f"Error2: The Excel file must contain {total_columns_string} column."
            print(string_error)
                
        
    else:
        string_error = f"Error:3 The Excel file must contain {', '.join(included_col)} column."
        print(string_error)
        
