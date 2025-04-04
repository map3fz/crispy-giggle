import pandas as pd
import numpy as np

def main():
    ''' This is a function meant to intake a colony report excel file path
    and return us a similar excel file where it describes what needs to occur
    within the colony for the week'''
    df = pd.read_excel(r"C:\Users\stapu\Downloads\SoftMouse.NET-Litter List-JohnLukens2025-04-02 0941.xlsx")
    df['Comment'] = df['Comment'].astype(str)
    # df = pd.read_excel(input("Enter the excel file path: "))
    transfer_and_tamoxifen(df)


def transfer_and_tamoxifen(df):


    transfer_df = df[df['Comment'].str.contains('__send to__', case=False, na=False) | df['Comment'].str.contains('__put on tam__', case=False, na=False)]

    print(transfer_df)

if __name__ == "__main__":
    main()


