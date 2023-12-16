import pandas as pd
import matplotlib.pyplot as plt

# Read Excel file into DataFrame
received_data = pd.read_excel(r"C:\Users\User\Desktop\python program\Group_Assignment\1.3.2 no edit.xlsx")

df = pd.DataFrame(received_data)

# Create a DataFrame using .iloc to select from excel rows 9 to 138 and columns 1 to 3
selected_rows_and_columns = pd.DataFrame(received_data.iloc[7:137, 0:3])

# Set column names
selected_rows_and_columns.columns = ['Year', 'Monthly', 'Total']

# Loop to fill NaN values in 'Year' column with 2013
for i in range(8, 19):
    selected_rows_and_columns.at[i, 'Year'] = 2013

for i in range(20, 31):
    selected_rows_and_columns.at[i, 'Year'] = 2014

for i in range(32, 43):
    selected_rows_and_columns.at[i, 'Year'] = 2015

for i in range(44, 55):
    selected_rows_and_columns.at[i, 'Year'] = 2016

for i in range(56, 67):
    selected_rows_and_columns.at[i, 'Year'] = 2017

for i in range(68, 79):
    selected_rows_and_columns.at[i, 'Year'] = 2018

for i in range(80, 91):
    selected_rows_and_columns.at[i, 'Year'] = 2019

for i in range(92, 103):
    selected_rows_and_columns.at[i, 'Year'] = 2020

for i in range(104, 115):
    selected_rows_and_columns.at[i, 'Year'] = 2021

for i in range(116, 127):
    selected_rows_and_columns.at[i, 'Year'] = 2022

for i in range(128, 137):
    selected_rows_and_columns.at[i, 'Year'] = 2023

# Extract the first column
selected_total_net_claims_government = pd.DataFrame(received_data.iloc[7:137, 3])
selected_total_net_claims_government.columns = ['Total Amount Of Net Claim government']

# Extract the second column
selected_total_claims_private_sector = pd.DataFrame(received_data.iloc[7:137, 6])
selected_total_claims_private_sector.columns = ['Total Amount Of Claim Private Sector']

# Extract the third column
selected_total_net_foreign_assets = pd.DataFrame(received_data.iloc[7:137, 9])
selected_total_net_foreign_assets.columns = ['Total Amount Of Net Foreign Assets']

# Extract the fourth column
selected_total_other_influences = pd.DataFrame(received_data.iloc[7:137, 12])
selected_total_other_influences.columns = ['Total Amount Of Other Influences']

# Concatenate the DataFrames horizontally
resulting_dataframe = pd.concat([selected_rows_and_columns[['Year', 'Monthly']], selected_total_net_claims_government, selected_total_claims_private_sector, selected_total_net_foreign_assets, selected_total_other_influences], axis=1)

# Group by 'Year' and sum the totals for each sector
sum_by_year_and_sector = resulting_dataframe.groupby(['Year']).agg({
    'Total Amount Of Net Claim government': 'sum',
    'Total Amount Of Claim Private Sector': 'sum',
    'Total Amount Of Net Foreign Assets': 'sum',
    'Total Amount Of Other Influences': 'sum'
}).reset_index()

def show_piechart_3sector():
    # Sum the total amounts for each sector over the entire period
    total_net_claim_government = sum_by_year_and_sector['Total Amount Of Net Claim government'].sum()
    total_claim_private_sector = sum_by_year_and_sector['Total Amount Of Claim Private Sector'].sum()
    total_net_foreign_assets = sum_by_year_and_sector['Total Amount Of Net Foreign Assets'].sum()

    # Labels for the sectors (excluding 'Other Influences')
    labels = ['Net Claim Government', 'Claim Private Sector', 'Net Foreign Assets']

    # Values for each sector (excluding 'Other Influences')
    values = [total_net_claim_government, total_claim_private_sector, total_net_foreign_assets]

    # Plotting the pie chart
    plt.figure(figsize=(8, 8))
    plt.pie(values, labels=labels, autopct='%1.1f%%', startangle=140, colors=['#1f77b4', '#ff7f0e', '#2ca02c'])

    # Adding title
    plt.title('Distribution of Total Amounts for 3 Sectors Over 10 Years and 10 Months')

    # Display the pie chart
    plt.show()

def create_Line_Chart2():
    # Plotting the line chart
    plt.figure(figsize=(12, 6))

    # Plot each sector
    plt.plot(sum_by_year_and_sector['Year'], sum_by_year_and_sector['Total Amount Of Net Claim government'], marker='o', label='Net Claim Government')
    plt.plot(sum_by_year_and_sector['Year'], sum_by_year_and_sector['Total Amount Of Claim Private Sector'], marker='o', label='Claim Private Sector')
    plt.plot(sum_by_year_and_sector['Year'], sum_by_year_and_sector['Total Amount Of Net Foreign Assets'], marker='o', label='Net Foreign Assets')
    plt.plot(sum_by_year_and_sector['Year'], sum_by_year_and_sector['Total Amount Of Other Influences'], marker='o', label='Other Influences')

    # Adding labels and title
    plt.title('Total Amount Of Each Year For 3 Sector And Other Influences Over the Years')
    plt.xlabel('Year')
    plt.ylabel('Total Amount (RM Million)')

    # Adding legend
    plt.legend()

    # Display the line chart
    plt.grid(True)
    plt.show()

def show_table_data_3Sector_OtherInfluences(resulting_dataframe):
    print("\nTotal Amount Of 3 Sector And Other Influences from 2013 Until October 2023")
    print(resulting_dataframe)
    print("...............................Total Amount Of Each Year For 3 Sector And Other Influences...............................")
    print(sum_by_year_and_sector)
    show_Line_Chart = input("\nDo you want to show the line chart of Total Amount Of Each Year For 3 Sector And Other Influences? Please insert 'Yes': ")
    if show_Line_Chart.lower() == 'yes':
        create_Line_Chart2()
    else:
        print("You don't want to show. Thank you")

# Function to display the table data
def show_table_data_Factors_Affecting_M3(data):
    print("\nTable Data Of Total Amount Of Factors Affecting M3 From 2013 Until October 2023:")
    print(data)

# Function to create and display the line chart
def create_line_chart(data):
    # Handle non-finite values in 'Year' column
    data['Year'] = pd.to_numeric(data['Year'], errors='coerce').astype('Int64')

    # Plotting the data

    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf', '#00ff00']

    plt.figure(figsize=(12, 6))

    # Extract unique years
    unique_years = data['Year'].unique()

    # Loop through each unique year and plot a line
    for year, color in zip(unique_years, colors):
        year_data = data[data['Year'] == year]
        # Plot all available data for the year 2023, even if incomplete
        if year == 2023:
            plt.plot(year_data['Monthly'], year_data['Total'], marker='o', label=f'Year {year}', color=color)
        else:
            # For other years, plot normally
            plt.plot(year_data['Monthly'], year_data['Total'], marker='o', label=f'Year {year}', color=color)

    # Adjust the x-axis ticks and labels
    month_labels = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    unique_months = range(1, 13)
    plt.xticks(unique_months, month_labels)

    # Adding titles and legend
    plt.title('Total Amount Of Factors Affecting M3 From 2013 Until October 2023')
    plt.xlabel('Month')
    plt.ylabel('Total (RM million)')
    plt.legend(title='Year', bbox_to_anchor=(1.05, 1), loc='upper left')  # Move legend outside
    plt.grid(True)

    # Show the chart
    plt.show()

# Main menu
while True:
    print("\nMain Menu:")
    print("1. Show the table data of Total Amount Of Factors Affecting M3 From 2013 Until October 2023")
    print("2. Show Line Chart of Total Amount Of Factors Affecting M3 From 2013 Until October 2023")
    print("3. Show the table and Line Chart of the Total Amount Of 3 Sector And Other Influences from 2013 Until October 2023")
    print("4. Show the pie Chart for 3 Sector And Other Influences During 10 years and 10 months")
    print("5. Exit")

    choice = input("Enter your choice (1, 2, 3, 4 OR 5): ")

    if choice == '1':
        show_table_data_Factors_Affecting_M3(selected_rows_and_columns)
    elif choice == '2':
        create_line_chart(selected_rows_and_columns)
    elif choice == '3':
        show_table_data_3Sector_OtherInfluences(resulting_dataframe)
    elif choice == '4':
        show_piechart_3sector()
    elif choice == '5':
        print("Exiting the program. Goodbye!")
        break
    else:
        print("Invalid choice. Please enter 1, 2, 3, 4 OR 5.")
