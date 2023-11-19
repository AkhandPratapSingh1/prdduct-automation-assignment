# Code by - Akhand Pratap singh

import pandas as pd
import matplotlib.pyplot as plt
import openpyxl

def analyze_and_visualize_data(input_file, output_file):
    input_data = pd.read_csv(input_file)
    high_rating_data = input_data[input_data['Feedback Rating'] >= 4]

    # Calculate summary statistics
    mean_rating = high_rating_data['Feedback Rating'].mean()
    median_rating = high_rating_data['Feedback Rating'].median()
    std_dev_rating = high_rating_data['Feedback Rating'].std()

    print("Mean: {}, Median: {}, Std Dev: {}".format(mean_rating, median_rating, std_dev_rating))

    # Create a bar chart
    high_rating_data['Feedback Rating'].hist(bins=5, edgecolor='black')
    plt.title('Distribution of Feedback Ratings')
    plt.xlabel('Feedback Rating')
    plt.ylabel('Frequency')
    plt.savefig('BarChartOfData.png')
    plt.close()

    # Save high rating data to a CSV file
    high_rating_data.to_csv(output_file, index=False)

    # Create a DataFrame for summary statistics
    summary_data = pd.DataFrame({
        'Statistic': ['Mean', 'Median', 'Std Dev'],
        'Value': [mean_rating, median_rating, std_dev_rating]
    })

    # Save summary data to the same CSV file in a new sheet
    with open(output_file, 'a') as f:
        f.write('\n\n')  # Add empty lines for separation
        summary_data.to_csv(f, index=False, header=True)

    # Save summary statistics and chart to an Excel file with separate sheets
    with pd.ExcelWriter('output_data.xlsx', engine='openpyxl') as writer:
        high_rating_data.to_excel(writer, sheet_name='HighRatingData', index=False)

        # Create a new sheet named 'Summary' for summary statistics
        summary_sheet = writer.book.create_sheet('Summary')

        # Write summary statistics to the 'Summary' sheet
        summary_sheet['A1'] = 'Summary Statistics'
        summary_sheet['A3'] = 'Mean:'
        summary_sheet['B3'] = mean_rating

        summary_sheet['A4'] = 'Median:'
        summary_sheet['B4'] = median_rating

        summary_sheet['A5'] = 'Std Dev:'
        summary_sheet['B5'] = std_dev_rating

        # Save the chart to the Excel file
        img = openpyxl.drawing.image.Image('BarChartOfData.png')
        summary_sheet.add_image(img, 'D8')

    print("Data and summary statistics saved to {}".format(output_file))


# Specify input and output file paths
input_file_path = 'inputProductCsv.csv'
output_file_path = 'output_data.csv'

# Execute the analysis and visualization function
analyze_and_visualize_data(input_file_path, output_file_path)
