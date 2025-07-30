import pandas as pd
from datetime import timedelta

# Change this to the actual path of your Excel file
file_path = r"D:\Data_Analysis\Projects\spotify_data\DataSet\spotify_history.xlsx"

# Change this to your actual column name or index (e.g., "Name" or 0)
column_name = 'track_name'


def read_column_values(file_path, column_name):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, engine='openpyxl')

        # Check if column exists
        if column_name not in df.columns:
            print(f"Column '{column_name}' not found in the file.")
            return {}

        # Get the column values as a list, drop NaN values
        values = df[column_name].dropna().tolist()
        return {"column_name": column_name, "no_of_rows": len(values)}

    except Exception as e:
        print(f"Error reading file: {e}")
        return {}


def generate_top_10_trending_songs(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, engine='openpyxl')

        # Convert 'date' column to datetime format (handle errors)
        df['date'] = pd.to_datetime(df['date'], dayfirst=True, errors='coerce')

        # Get the most recent date in the dataset
        latest_date = df['date'].max()

        # Calculate the date 3 months (90 days) before the latest date
        three_months_ago = latest_date - timedelta(days=90)

        # Filter data within this 3-month window
        recent_df = df[df['date'] >= three_months_ago]

        if recent_df.empty:
            print("No data found in the last 3 months of the dataset.")
            return

        # Compute top 10 songs by 'ms_played'
        top_tracks = (recent_df.groupby('track_name')['ms_played']
                      .sum()
                      .sort_values(ascending=False)
                      .head(10)
                      .reset_index())

        # Add artist information
        top_tracks = top_tracks.merge(df[['track_name', 'artist_name']].drop_duplicates(),
                                      on='track_name', how='left')

        # Save to Excel
        output_file = "ResultData/top_10_trending_songs.xlsx"
        top_tracks.to_excel(output_file, index=False)

        print(f"Top 10 trending songs (based on most recent 3 months) saved to '{output_file}'")

    except Exception as e:
        print(f"Error generating top 10 trending songs: {e}")


# Example usage
if __name__ == "__main__":
    # Read column values
    values = read_column_values(file_path, column_name)
    print("Column values:")
    print(values)

    # Generate top 10 trending songs
    generate_top_10_trending_songs(file_path)
