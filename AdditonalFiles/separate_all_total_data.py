import pandas as pd


def seperate_total_data():
    # # separate total seen rx data
    seen_df = pd.read_excel('./Data/SeenRx/Seen_Rx_Data.xlsx')
    seen_df['FFTR'] = seen_df['FFTR'].astype('str')
    mask_seen = (seen_df['FFTR'].str.len() < 4)
    seen_df = seen_df.loc[mask_seen]
    seen_df.to_csv('./Data/SeenRx/rsm_seen_total.csv', index=False)

    # # separate total doctor call data
    call_df = pd.read_excel('./Data/Call/doctor_call_data.xlsx')
    call_df['FFTR'] = call_df['FFTR'].astype('str')
    mask_call = (call_df['FFTR'].str.len() < 4)
    call_df = call_df.loc[mask_call]
    call_df.to_csv('./Data/Call/rsm_call_total.csv', index=False)
