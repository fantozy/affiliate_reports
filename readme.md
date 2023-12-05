# Python Developer Test

## Introduction

This Python application is designed to demonstrate proficiency in Python and pandas, along with the ability to work with data, clean it, transform it, and perform aggregations. The goal is to create a simple Python application that fulfills the specified tasks.

## Tasks

### 1. Load & Clean Data

- Handle duplicates, typos / inconsistencies, and missing values in the provided data tables: orders, currency rates, and affiliate rates.

### 2. Convert “Order Amount” into EUR

- Apply the provided currency rates to convert the "Order Amount" into Euros.

### 3. Calculate Fees for Each Order

- Calculate fees for each order using the appropriate affiliate rates. The fee formulas are as follows:
  - Processing Fee = Order Amount (EUR) * Processing Rate (applied to all transactions).
  - Refund Fee = Refund Fee (applied to orders with Order Status = 'Refunded').
  - Chargeback Fee = Chargeback Fee (applied to orders with Order Status = 'Chargeback').

### 4. Generate Weekly Aggregation for Each Affiliate

- For each affiliate, generate a weekly aggregation and save an Excel report (`{affiliate_name}.xlsx`) with the following columns:
  - Week (format: 01-10-2023 - 07-10-2023, week start day: Sunday)
  - Number of Orders
  - Total Order Amount (EUR)
  - Total Processing Fee
  - Total Refund Fee
  - Total Chargeback Fee

### Running the Application

To run the application, follow these steps:

1. Clone the repository: `git clone https://github.com/fantozy/affiliate_reports.git`
2. Navigate to the project directory: `cd affiliate_reports`
3. Install the required dependencies: `pip install -r requirements.txt`
4. Run the application: `python main.py`

Ensure that the necessary data files are present in the correct locations and are formatted as expected.

## Notes

- Each affiliate rate setup is applied from the start date up to:
  - The start date (not including) of the chronologically next setup for the same affiliate.
  - Or if there is no next setup, up to and including today.

- We should be able to run your application on similar data and obtain correct results.

## Author

Nika Beglarishvli

