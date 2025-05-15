import pandas as pd

# 🔁 Change this to the actual path of your Excel file
file_path = r"D:\Data_Analysis\Projects\spotify_data\DataSet\spotify_history.xlsx"

# 🔁 Change this to your actual column name or index (e.g., "Name" or 0)
column_name = 'track_name'


def read_column_values(file_path, column_name):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, engine='openpyxl')

        # Check if column exists
        if column_name not in df.columns:
            print(f"Column '{column_name}' not found in the file.")
            return []

        # Get the column values as a list, drop NaN values
        values = df[column_name].dropna().tolist()
        return {"column_name": "track_name", "no_of_rows": len(values)}
    except Exception as e:
        print(f"Error reading file: {e}")
        return []


# ✅ Example usage
if __name__ == "__main__":
    values = read_column_values(file_path, column_name)
    print("Column values:")
    print(values)
