import dynamic_yaml
import pandas as pd

yml = '''
names:
  n1: "name1"
  n2: "name2"
columns:
  - "{names.n1}"
  - "{names.n2}"
'''


if __name__ == '__main__':
    config = dynamic_yaml.load(yml)

    cols = [config.columns[x] for x in range(len(config.columns))]

    df = pd.DataFrame(columns=cols)

    print(df.columns)