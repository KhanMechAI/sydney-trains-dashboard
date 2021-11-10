import numpy as np
import pandas as pd


def _cm_to_inch(length):
    return np.divide(length, 2.54)


def resolve_list(config_list):
    return [config_list[x] for x in range(len(config_list))]


def resolve_dictionary(config_dict):
    return {k: v for k, v in config_dict.items()}


def reset_index(df, index_start: int = 1):
    df.reset_index(inplace=True, drop=True)

    # make index start from the new index start point, default index is 0, new default is 1
    df.index = df.index + index_start
    return df


def drop_empty_rows(df) -> pd.DataFrame:
    """
    Pandas doesnt recognise empty string as an empty value. So change all empty strings to nan, then drop and
    replace all nans with empty strings.
    """
    df.replace(["", " "], np.nan, inplace=True)
    df.dropna(inplace=True, how="all")
    df.replace(np.nan, "", inplace=True)

    # need to reset index after dropped rows.
    df = reset_index(df)

    return df