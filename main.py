import pandas as pd
from datetime import timedelta

def load_and_clean_data():
    """
    Load and clean the data from Excel files.
    
    Returns:
    - test_affiliate: DataFrame containing affiliate rates data.
    - test_currency: DataFrame containing currency rates data.
    - test_orders: DataFrame containing orders data.
    """
    # Load data from Excel into DataFrame
    test_affiliate = pd.read_excel('test-affiliate-rates.xlsx')
    test_currency = pd.read_excel('test-currency-rates.xlsx')
    test_orders = pd.read_excel('test-orders.xlsx')

    # Remove duplicate rows in the orders DataFrame
    test_orders.drop_duplicates(inplace=True)

    # Replace 'None' and empty values with 'Unknown'
    test_orders['Affiliate ID'] = test_orders['Affiliate ID'].replace({'none': 'Unknown', pd.NaT: 'Unknown'})

    return test_affiliate, test_currency, test_orders

def convert_to_eur(test_currency: pd.DataFrame, test_orders: pd.DataFrame) -> pd.DataFrame: 
    """
    Convert 'Order Amount' to EUR based on currency rates and update 'Currency' column.
    
    Args:
    - test_currency: DataFrame containing currency rates data.
    - test_orders: DataFrame containing orders data.
    """
    merged_data = pd.merge(test_orders, test_currency, how='left', left_on='Order Date', right_on='date')

    # Convert necessary columns to float to avoid data type issues
    merged_data['Order Amount'] = merged_data['Order Amount'].astype(float)
    merged_data['USD'] = merged_data['USD'].astype(float)
    merged_data['GBP'] = merged_data['GBP'].astype(float)

    # Iterate through rows and convert 'Order Amount' to EUR based on the currency rate
    for index, row in merged_data.iterrows():
        if row['Currency'] == 'USD':
            merged_data.at[index, 'Order Amount'] = row['Order Amount'] * row['USD']
        elif row['Currency'] == 'GBP':
            merged_data.at[index, 'Order Amount'] = row['Order Amount'] * row['GBP']
    
    merged_data['Currency'] = merged_data['Currency'].replace({'GBP': 'EUR', 'USD': 'EUR'})    
    merged_data.drop(['date', 'USD', 'GBP'], axis=1, inplace=True)
    return merged_data

def calculate_fees(test_affiliate:pd.DataFrame, test_orders:pd.DataFrame) -> pd.DataFrame:
    """
    Calculating fees based on test_affiliate 'Chargeback Fee', 'Refund Fee', 'Proccesing Rate'

    Args:
    - test_affiliate: DataFrame containing Affiliate data.
    - test_orders: DataFrame containig orders data.

    """
    # groupby Affiliate ID
    aggregated_fees = (
        test_affiliate.groupby("Affiliate ID")
        .agg({"Processing Rate": "last", "Chargeback Fee": "last", "Refund Fee": "last"})
        .reset_index()
    )
    
    # Merge test_orders and aggregated_fees based on Affiliate ID
    merged_data = pd.merge(test_orders, aggregated_fees, on="Affiliate ID", how="left")
    merged_data['Processing Fee'] = merged_data['Order Amount'] * merged_data['Processing Rate']

    refunded_orders = merged_data[merged_data['Order Status'] == 'Refunded']
    merged_data.loc[refunded_orders.index, 'Refund Fee'] = refunded_orders['Refund Fee']
    
    chargeback_orders = merged_data[merged_data['Order Status'] == 'Chargeback']
    merged_data.loc[chargeback_orders.index, 'Chargeback Fee'] = chargeback_orders['Chargeback Fee']
    return merged_data

    
def aggregate_and_save_reports(proceseed_data:pd.DataFrame, test_affiliate:pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate and save reports for each affiliate based on processed data.

    Args:
    - processed_data: DataFrame containing processed order data.
    - test_affiliate: DataFrame containing affiliate rate data.
    """
    proceseed_data['Order Date'] = pd.to_datetime(proceseed_data['Order Date'])
    
    def get_week_start_date(date):
        """
        Get the start date of the week for a given date.

        Args:
        - date: Date for which the week start date is calculated.

        Returns:
        - Week start date.
        """
        return date - timedelta(days=date.weekday())

    proceseed_data['Week Start Date'] = proceseed_data['Order Date'].apply(get_week_start_date)
    
    # Aggregate data on 'Affiliate ID' and 'Week Start Date'
    weekly_aggregation = (
        proceseed_data.groupby(['Affiliate ID', 'Week Start Date'])
        .agg({
            'Order Number': 'count',
            'Order Amount': 'sum',
            'Processing Fee': 'sum',
            'Refund Fee': 'sum',
            'Chargeback Fee': 'sum'
        })
        .reset_index()
    )
    # Rename columns for clarity
    weekly_aggregation.columns = [
        'Affiliate ID',
        'Week Start Date',
        'Number of Orders',
        'Total Order Amount (EUR)',
        'Total Processing Fee',
        'Total Refund Fee',
        'Total Chargeback Fee'
    ]

    affiliate_ids = weekly_aggregation['Affiliate ID'].unique()
    
    # Iterate through unique affiliate IDs
    for affiliate in affiliate_ids:
        affiliate_data = weekly_aggregation[weekly_aggregation['Affiliate ID'] == affiliate]
        
        # Check if affiliate data exists in test_affiliate DataFrame
        if not test_affiliate[test_affiliate['Affiliate ID'] == affiliate].empty:
            affiliate_name = test_affiliate[test_affiliate['Affiliate ID'] == affiliate]['Affiliate Name'].iloc[0]
            filename = f"{affiliate_name}.xlsx"
            
            # Save data to Excel file
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                affiliate_data.to_excel(writer, sheet_name='Weekly Aggregation', index=False)
        else:
            print(f"No data found for Affiliate ID {affiliate}.")
        


def main():
    test_affiliate, test_currency, test_orders = load_and_clean_data()
    merged_data = convert_to_eur(test_currency, test_orders)
    proceseed_data = calculate_fees(test_affiliate, merged_data)
    aggregate_and_save_reports(proceseed_data, test_affiliate)

if __name__ == '__main__':
    main()
