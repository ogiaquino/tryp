import os
from pandas.io.parsers import read_csv
from parser import parse
from crosstab import Crosstab


class CrosstabMetaData(object):
    def __init__(self, tryp_file, csv_file, output_file):
        self.report = parse(tryp_file, 'crosstab')
        self.source_dataframe = self.data_frame(csv_file)
        self.xaxis = self.report['xaxis']
        self.yaxis = self.report['yaxis']
        self.zaxis = self.report['zaxis']
        self.visible_yaxis_summary = self.report['visible_yaxis_summary']
        self.visible_xaxis_summary = self.report['visible_xaxis_summary']
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
        ctmeta = CrosstabMetaData(tryp_file, csv_file, output_file)
        self.crosstab = Crosstab(ctmeta)


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
