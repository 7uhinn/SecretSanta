import pandas as pd
import random

def filter_entries(df, filter_column, filter_value):

    # Filter entries based on the specified column and value
    df = df[df[filter_column] == filter_value].reset_index(drop=True)

    return df

def create_secret_santa_pairs(df):
    # Shuffle the DataFrame to randomize the order
    df = df.sample(frac=1).reset_index(drop=True)

    # Create pairs
    pairs = []
    receivers = df[['Name', 'Email']].to_dict(orient='records')

    for giver in df[['Name', 'Email']].to_dict(orient='records'):
        available_receivers = [receiver for receiver in receivers if receiver['Name'] != giver['Name']]
        if not available_receivers:
            raise ValueError("Error: Unable to create pairs. Please check your input data.")

        receiver = random.choice(available_receivers)
        pairs.append({'Giver': giver['Name'], 'Giver_Email': giver['Email'], 'Receiver': receiver['Name'], 'Receiver_Email': receiver['Email']})
        receivers.remove(receiver)

    return pairs

def write_to_excel(pairs, output_file):
    # Create a DataFrame from the pairs
    df_result = pd.DataFrame(pairs, columns=['Giver', 'Giver_Email', 'Receiver'])

    # Write the DataFrame to an Excel file
    df_result.to_excel(output_file, index=False)
    print(f"Secret Santa pairs written to {output_file}")

def main():
    # Assign a default value to excel_file_path
    excel_file_path = "Secret Santa - GCOC(1-42).xlsx"
    
    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file_path)

    # Filter entries in the Excel file
    filtered_df = filter_entries(df, df.columns[7],"Yes")
    filtered_df = filter_entries(filtered_df, df.columns[6], "Yes, I would love to participate in the Secret Santa gift exchange!")
    
    # Create Secret Santa pairs
    secret_santa_pairs = create_secret_santa_pairs(filtered_df)
    
    # Write the pairs to the output Excel file
    write_to_excel(secret_santa_pairs, "Secret Santa Pairs.xlsx")

if __name__ == "__main__":
    main()
