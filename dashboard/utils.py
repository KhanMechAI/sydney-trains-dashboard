import numpy as np


def _cm_to_inch(length):
    return np.divide(length, 2.54)


def resolve_list(config_list):
    return [config_list[x] for x in range(len(config_list))]


def resolve_dictionary(config_dict):
    return {k: v for k, v in config_dict.items()}