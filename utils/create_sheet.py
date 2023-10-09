import os
import random

import pandas as pd

row_num = range(1, 21)

dir_path = '01_raw/'


class ExcelFile(object):

    def __init__(self):
        self.names = ['Student' + str(i) for i in row_num]

    def _to_dataframe(self, seed):
        random.seed(seed)

        math = [random.randint(30, 100) for _ in row_num]
        english = [random.randint(30, 100) for _ in row_num]

        return pd.DataFrame({
            'math': math,
            'english': english}, index=self.names)

    def create(self, seed):
        file_name = '期末テスト特典表_{}.xlsx'.format(seed)

        df = self._to_dataframe(seed)

        if not os.path.isdir(dir_path):
            os.makedirs(dir_path)

        df.to_excel(dir_path + file_name, encoding='utf-8')


if __name__ == "__main__":
    excel_file = ExcelFile()

    excel_file.create(seed=0)
    excel_file.create(seed=1)
