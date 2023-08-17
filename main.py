import pandas as pd

# Load the Excel file into a DataFrame
excel_file_path = './Base_data_R1.xlsx'

# Define required columns
required_columns = ['Prod Type- L3', 'Parent SKU #1', 'Volume Completed', 'Container size']


try:
    df = pd.read_excel(excel_file_path)

    # Check if all required columns exist in the DataFrame
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")

    # Filter and group data by buckets and SKU
    grouped = df.groupby(['Prod Type- L3', 'Parent SKU #1'])

    # Initialize lists to store calculated values
    buckets = []
    skus = []
    batch_counts = []
    min_batch_sizes = []
    avg_batch_sizes = []
    max_batch_sizes = []
    avg_container_sizes = []

    # Iterate through each group and calculate required statistics
    for (bucket, sku), group_df in grouped:
        batch_counts.append(len(group_df))
        min_batch_sizes.append(group_df['Volume Completed'].min())
        avg_batch_sizes.append(group_df['Volume Completed'].mean())
        max_batch_sizes.append(group_df['Volume Completed'].max())
        avg_container_sizes.append(group_df['Container size'].mean())
        buckets.append(bucket)
        skus.append(sku)

    # Create a new DataFrame with calculated values
    result_df = pd.DataFrame({
        'Bucket': buckets,
        'SKU': skus,
        'Batch Count': batch_counts,
        'Min Batch Size': min_batch_sizes,
        'Avg Batch Size': avg_batch_sizes,
        'Max Batch Size': max_batch_sizes,
        'Avg Container Size': avg_container_sizes
    })

    # Save the result DataFrame to a new Excel file
    result_excel_path = 'result.xlsx'
    result_df.to_excel(result_excel_path, index=False)

    print("Table creation complete. Result saved to", result_excel_path)
    
except pd.errors.FileFormatError:
    print("Error reading the Excel file. Please make sure the file is in the correct format.")
except ValueError as ve:
    print(ve)