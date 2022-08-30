import xlwings as xw
import pandas as pd
from os.path import join
import re
import os


class MasterWB:
    def __init__(self):
        # self.wb_model = xw.Book.caller()
        self.app = xw.App(visible=False)
        self.wb_model = self.app.books.open(r'C:\Users\BTHAMAN\Documents\NPDES\ECHO_All_DO_NOT_EDIT_DATA.xlsm')
        self.working_dir = None
        self.set_working_dir()
        self.outfall_file_path = join(self.working_dir, 'DischargesByBasin')
        self.sht_active = self.wb_model.sheets.active
        self.sht_data = self.wb_model.sheets('Data')
        self.sht_outfalls = self.wb_model.sheets('Outfalls')

    def set_working_dir(self):
        thepath = self.wb_model.fullname
        dir = os.path.dirname(thepath)
        os.chdir(dir)
        self.working_dir = dir

    def enable_autofilter(self):
        rng_to_filter = self.wb_model.sheets('Hierarchy').range('A1')
        # rng_to_filter.select()  # don't need to select to apply filter
        rng_to_filter.api.AutoFilter(1)  # only applies filter; does not remove it

    def disable_autofilter(self):
        sht = self.wb_model.sheets('Hierarchy')
        sht.api.AutoFilterMode = False

    def close(self):
        self.wb_model.close()
        self.app.kill()

    def save(self, path=None):
        if path:
            self.wb_model.save(path=path)
        else:
            self.wb_model.save()

    def get_df(self, outfile):
        wb = self.app.books.open(outfile)
        sht = wb.sheets('Sheet1')
        rng = sht.range('A1').expand()
        df = rng.options(convert=pd.DataFrame, index=True).value
        df.drop('Total', axis=1, inplace=True)
        for c in df:
            if '_1' in c:
                df.drop(c, axis=1, inplace=True)
        wb.close()
        return df

    def get_outfall_files(self):
        try:
            pattern = r'^SA_Guad.*\.xlsx$'
            re_pattern = re.compile(pattern, re.IGNORECASE)
            out_files = []
            for root, dirs, files in os.walk(self.outfall_file_path):
                for basename in files:
                    if re_pattern.match(basename):
                        pth = join(root, basename)
                        out_files.append(pth)
            return out_files
        except Exception as e:
            print('error in get_outfall_files: ' + str(e))
            return []

    def write_outfall_ids(self, df):
        self.sht_outfalls.range('A2').expand().clear_contents()
        self.sht_outfalls.range('A2').options(transpose=True).value = df.columns.T.values


if __name__ == "__main__":
    wb = MasterWB()
    out_files = wb.get_outfall_files()
    df_list = []
    for f in out_files:
        print(f)
        df_list.append(wb.get_df(outfile=f))
    df_all = pd.concat(df_list, axis=1).sort_index()
    # duplicate column names exist, and need to be combined
    df_all = df_all.groupby(level=0, axis=1).sum()
    # write df to data sheet
    wb.sht_data.range('A1').expand().clear_contents()
    wb.sht_data.range('A1').options(index=True).value = df_all
    print('\nFiles processed\n')
    # write out max of each column
    dfmax = df_all[df_all.columns].max()
    dfmin = df_all[df_all.columns].min()
    dfavg = df_all[df_all.columns].mean()
    dfstats = pd.concat([dfmax, dfmin, dfavg], axis=1).sort_index()
    dfmax.columns = ['OUTFALL_ID', 'MAX_VALUE', 'MIN_VALUE', 'MEAN']
    print(dfstats)
    wb.sht_outfalls.range('A1').options(index=True).value = dfstats
    wb.save()
    wb.close()
