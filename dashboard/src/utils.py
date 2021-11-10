import numpy as np
import pandas as pd


class DynamicResolver:
    yaml_list_type = "YamlList"
    yaml_dict_type = "YamlDict"

    @staticmethod
    def _get_type_string(obj):
        return obj.__class__.__name__

    @staticmethod
    def _resolve_dictionary(dyn_list):
        return {k: v for k, v in dyn_list.items()}

    @staticmethod
    def _resolve_list(dyn_dict):
        return [dyn_dict[x] for x in range(len(dyn_dict))]

    @classmethod
    def _resolve_dynamic_yaml_object(cls, d_yml_ob):
        obj_type: str = cls._get_type_string(d_yml_ob)

        if obj_type == cls.yaml_list_type:
            partial_resovled_list = cls._resolve_list(d_yml_ob)
            fully_resolved_list = []
            for element in partial_resovled_list:
                fully_resolved_list.append(cls._resolve_dynamic_yaml_object(element))

            return fully_resolved_list
        elif obj_type == cls.yaml_dict_type:
            partial_resovled_dict = cls._resolve_dictionary(d_yml_ob)
            fully_resolved_dict = {}
            for k, v in partial_resovled_dict.items():
                fully_resolved_dict[k] = cls._resolve_dynamic_yaml_object(v)
            return fully_resolved_dict
        else:
            return d_yml_ob

    @classmethod
    def resolve(cls, dynamic_yaml_object):
        return cls._resolve_dynamic_yaml_object(dynamic_yaml_object)


def _cm_to_inch(length):
    return np.divide(length, 2.54)


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
