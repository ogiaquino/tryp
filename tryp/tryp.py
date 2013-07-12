import os
from pandas.io.parsers import read_csv

from excel import to_excel
from dataset import Dataset
from parser import parse


class Tryp(object):
    def __init__(self, tryp_file, csv_file, output_file):
        self.report = parse(tryp_file)
        self.df = self.data_frame(csv_file)
        self.rows = self.report['rows']
        self.columns = self.report['columns']
        self.values = self.report['values']
        self.labels = self.report['labels']
        self.rows_totals = self.report['rows_totals']
        self.columns_totals = self.report['columns_totals']

        self.excel = {}
        self.excel['filename'] = output_file
        self.excel['sheetname'] = 'Sheet1'

        self.extmodule = self.is_extmodule_exist(tryp_file)
        self.crosstab = Dataset(self).crosstab

    def is_extmodule_exist(self, tryp_file):
        tryp_path = os.path.abspath(tryp_file)
        tryp_path, tryp_file = os.path.split(tryp_file)
        tryp_filename = os.path.splitext(tryp_file)[0]
        extmodule = os.path.join(tryp_path, tryp_filename + '.py')
        if os.path.exists(extmodule):
            return (tryp_filename, extmodule)

    def to_excel(self):
        return to_excel(self)

    def data_frame(self, csv_file):
        df = read_csv(csv_file)
        return df


def main():
        import argparse
        parser = argparse.ArgumentParser(description='Generate Excel File.')
        parser.add_argument('-f')
        parser.add_argument('-d')
        parser.add_argument('-o')
        args = parser.parse_args()
        tryp_file = args.f
        output_file = args.o
        csv_file = args.d
        Tryp(tryp_file, csv_file, output_file).to_excel()

if __name__ == '__main__':
    main()
