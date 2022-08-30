import xlwings as xw
import pandas as pd
import numpy as np
import os


class MasterWB:
    def __init__(self):
        # self.wb_model = xw.Book.caller()
        self.app = xw.App(visible=False)
        self.wb_model = self.app.books.open(r'C:\Users\BThaman\OneDrive - HDR, Inc\Documents\PycharmProjects\TWDB_ASR\ECHO_All_DATA.xlsm')
        self.working_dir = None
        self.set_working_dir()
        self.sht_active = self.wb_model.sheets.active
        self.sht_data = self.wb_model.sheets('Data')
        self.sht_describe = self.wb_model.sheets('Describe')
        self.sht_annual = self.wb_model.sheets('Annual')
        self.sht_ann_min = self.wb_model.sheets('Annual_minimums')
        self.sht_ann_avg = self.wb_model.sheets('Annual_averages_acftperyear')

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


if __name__ == "__main__":
    wb = MasterWB()
    rng = wb.sht_data.range('A1').expand()
    df = rng.options(convert=pd.DataFrame, index=True).value
    # calc stats excluding zeros
    df = df.replace(0, np.nan)
    # example of filtering on specific months and years
    dfm1 = df[(df.index.month >= 4) & (df.index.month <= 6) & (df.index.year < 2020)]
    print(dfm1)
    # group by year and sum
    df = df[(df.index.year < 2020)]
    dfg = df.groupby(pd.Grouper(freq='Y'))
    dfg_mean = dfg.mean()
    dfg_cnt = dfg.count()
    print(dfg_cnt)
    wb.sht_annual.range('A1').expand().clear_contents()
    wb.sht_annual.range('A1').options(index=True).value = dfg_mean
    # get annual minimums
    dfg_sum = dfg_mean.replace(0, np.nan)
    dfmin = dfg_sum[dfg_sum.columns].min()
    wb.sht_ann_min.range('A1').expand().clear_contents()
    wb.sht_ann_min.range('A1').options(index=True).value = dfmin
    # get annual averages in units of ac-ft/year (mgd * 1120.14)
    dfavg = dfg_sum[dfg_sum.columns].mean() * 1120.14
    wb.sht_ann_avg.range('A1').expand().clear_contents()
    wb.sht_ann_avg.range('A1').options(index=True).value = dfavg
    # describe basic stats
    df_desc = df.describe(percentiles=[0.1, 0.25, 0.5, 0.75, 0.9])
    wb.sht_describe.range('A1').expand().clear_contents()
    wb.sht_describe.range('A1').options(index=True).value = df_desc
    wb.save()
    wb.close()
