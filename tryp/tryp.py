import os
from pandas.io.parsers import read_csv
from parser import parse
from crosstab import Crosstab


class MetaCrosstab(object):
    def __init__(self, tryp_file, csv_file, output_file):
        self.report = parse(tryp_file, 'crosstab')
        self.df = self.data_frame(csv_file)
        self.index = self.report['index']
        self.columns = self.report['columns']
        self.values = self.report['values']
        self.labels = self.report['labels']
        self.index_totals = self.report['index_totals']
        self.columns_totals = self.report['columns_totals']
        self.excel = {}
        self.excel['filename'] = output_file
        self.excel['sheetname'] = os.path.splitext(output_file)[0]
        self.extmodule = self.is_extmodule_exist(tryp_file)

    def is_extmodule_exist(self, tryp_file):
        tryp_path = os.path.abspath(tryp_file)
        tryp_path, tryp_file = os.path.split(tryp_file)
        tryp_filename = os.path.splitext(tryp_file)[0]
        extmodule = os.path.join(tryp_path, tryp_filename + '.py')
        if os.path.exists(extmodule):
            return (tryp_filename, extmodule)

    def data_frame(self, csv_file):
        df = read_csv(csv_file)
        return df


class Tryp(object):
    def __init__(self, tryp_file, csv_file, output_file):
        meta = MetaCrosstab(tryp_file, csv_file, output_file)
        self.crosstab = Crosstab(meta)


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
        Tryp(tryp_file, csv_file, output_file).crosstab.to_excel()

if __name__ == '__main__':
    main()
