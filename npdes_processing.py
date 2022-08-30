import xlwings as xw
import pandas as pd
import numpy as np
from scipy import stats
from os.path import join
import os
import logger_handler
import re
# import plot_flow as pf
import csv


def get_number(v):
    if isinstance(v, int) or isinstance(v, float):
        return v
    elif isinstance(v, str):
        pattern = r'-?\d*\.?\d+'
        matches = re.findall(pattern, v)
        if len(matches) > 0:
            if '.' in matches[0]:
                return float(matches[0])
            else:
                return int(matches[0])
        else:
            return None
    else:
        return v


def remove_outliers_stdv(df, num_stdev=3.0):
    if df[df.columns[0]].min() != df[df.columns[0]].max():
        z_scores = stats.zscore(df)
        abs_z_scores = np.abs(z_scores)
        filtered_entries = (abs_z_scores < num_stdev).all(axis=1)
        df = df[filtered_entries]
    return df


def remove_outliers_iq(df, whisker_ratio=1.5):
    q1 = df[df.columns[0]].quantile(0.25)
    q3 = df[df.columns[0]].quantile(0.75)
    iqr = q3 - q1
    if iqr == 0:
        stdev = df.std()[0]
        mean = df.mean()[0]
        limit_low = mean - 2.5 * stdev
        limit_high = mean + 2.5 * stdev
    else:
        limit_low = q1 - iqr * whisker_ratio
        limit_high = q3 + iqr * whisker_ratio
    print('Limit low = {}, Limit high = {}'.format(limit_low, limit_high))
    filter_include = (df[df.columns[0]] >= limit_low) & (df[df.columns[0]] <= limit_high)
    return df[filter_include]


class NewExcelWB:
    def __init__(self):
        self.wb = xw.Book()

    def df_to_excel(self, df, path):
        self.wb.sheets.active.range('A1').options(index=True).value = df
        self.enable_autofilter()
        self.columnformat()
        self.save(path=path)

    def enable_autofilter(self):
        # xw.Range('A1').api.AutoFilter(1)
        self.wb.sheets.active.range('A1').api.AutoFilter(1)

    def columnformat(self):
        sht = self.wb.sheets.active
        sht.api.Columns('A:LL').ColumnWidth = 20
        sht.api.Rows('1:1').Font.Bold = True

    def close(self):
        self.wb.close()

    def save(self, path=None):
        if path:
            self.wb.save(path=path)
        else:
            self.wb.save()


class DischargeDataProcessor:
    def __init__(self):
        self.app = None
        self.wb = None
        self.sht = None
        self.basins_list = None
        self.num_months = 0
        self.non_null_ratio_target = 1
        self.df_list = []
        self.data_folder = r'C:\Users\BThaman\documents\npdes\downloads'
        self.df_input = pd.read_csv('discharge_points.csv')
        self.logger_debug = logger_handler.setup_logger(name='processing_debug', log_file=join(os.getcwd(),'processing_debug.log'))
        self.logger_exception = logger_handler.setup_logger(name='processing_exception', log_file=join(os.getcwd(),'processing_error.log'))

    def set_basins_list(self, basin_list):
        self.basins_list = basin_list

    def get_basins_list(self):
        return self.basins_list

    def set_params(self, num_months, non_null_ratio_target):
        self.num_months = num_months
        self.non_null_ratio_target = non_null_ratio_target

    def npdes_dfs(self, out_path):
        def add_df(found_cell, pt_id):
            rng = self.sht.range(found_cell.Row + 1, 2).expand('right')
            for c in range(1, rng.columns.count + 1):
                if rng(1, c).value == 'Flow, in conduit or thru treatment plant':
                    if self.sht.range(rng(1, c).row + 2, rng(1, c).column).value == 'DAILY AV':
                        # know that the right column exists; create a df and keep only that column
                        rng_df = self.sht.range('A' + str(rng(1, c).row + 2)).expand(mode='table')
                        df = rng_df.options(convert=pd.DataFrame).value
                        df = df.iloc[:, [c - 1]]
                        df.columns = [id + '_' + str(pt_id)]
                        df.index = pd.to_datetime(df.index)
                        for col in df:
                            df[col] = df[col].apply(lambda x: get_number(x))
                        df = df.dropna()
                        # df = remove_outliers_stdv(df, num_stdev=2.5)
                        # df = remove_outliers_iq(df)
                        non_null_cnt = df.count()[0]
                        # row_cnt = len(df.index)
                        non_null_ratio = non_null_cnt / self.num_months
                        if non_null_ratio >= self.non_null_ratio_target:
                            # df = df.replace(0, np.nan)
                            print(df)
                            print('{} of {} have numeric values'.format(non_null_cnt, self.num_months))
                            self.df_list.append(df)
                        else:
                            print('\nSkipping ' + id + '_' + str(pt_id) + '\n')
                        continue
        try:
            self.app = xw.App(visible=False)
            self.wb = self.app.books.open(out_path)
            self.sht = self.wb.sheets('NPDES Monitoring Data Download')
            id = self.sht.range(4, 1).value.split(':')[1].strip()
            # ###########################################################################
            # get one or more dfs from the NPDES output file
            # ###########################################################################
            re_pattern = re.compile(r'.*Limit Set: (\d+) - 1 - A.*', re.IGNORECASE)
            found_all = False
            find_cell = self.sht.api.Cells.Find(What="Outfall - Monitoring Location - Limit Set:",
                                                After=self.sht.api.Cells(1, 1),
                                                LookAt=xw.constants.LookAt.xlPart,
                                                LookIn=xw.constants.FindLookIn.xlFormulas,
                                                SearchOrder=xw.constants.SearchOrder.xlByRows,
                                                SearchDirection=xw.constants.SearchDirection.xlPrevious,
                                                MatchCase=False)
            if find_cell:
                if re_pattern.findall(find_cell.value):
                    add_df(found_cell=find_cell, pt_id=re_pattern.findall(find_cell.value)[0])
                find_after = find_cell
                while not found_all:
                    find_next = self.sht.api.Cells.FindNext(After=find_after)
                    if find_next:
                        if find_next.Row == find_cell.Row and find_next.Column == find_cell.Column:
                            found_all = True
                        else:
                            find_after = find_next
                            if re_pattern.findall(find_next.value):
                                add_df(found_cell=find_next, pt_id=re_pattern.findall(find_next.value)[0])
            else:
                print('\nNO DATA in ' + id + '..........................\n')
        except Exception as e:
            print('Error: ' + str(e))
        finally:
            self.wb.close()
            self.app.kill()

    def process_files(self):
        try:
            for basin in self.basins_list:
                df_basin_input = self.df_input[self.df_input['BASIN_NAME'] == basin]
                for i, row in df_basin_input.iterrows():
                    f = join(self.data_folder, row['MONITORING_DATA_DOWNLOADED'])
                    self.npdes_dfs(f)
            print('.'*75)
            df_all = pd.concat(self.df_list, axis=1).sort_index()
            df_interp = df_all.interpolate(limit_area='inside')
            df_all['Total'] = df_all.sum(axis=1)
            df_interp['Total'] = df_interp.sum(axis=1)
            new_wb = NewExcelWB()
            path_raw = self.basins_list[0] + '_Raw.xlsx'
            path_interp = self.basins_list[0] + '_Interp.xlsx'
            new_wb.df_to_excel(df=df_all, path=path_raw)
            new_wb.df_to_excel(df=df_interp, path=path_interp)
            new_wb.close()
            # pf.plot_discharge(xl_raw=path_raw, xl_interp=path_interp, basin=self.basins_list[0])
            print(df_all.head(50))
        except Exception as e:
            self.logger_exception.exception('process_files')

    def calc_averages(self, county):
        self.df_list = []
        w = csv.writer(open('discharge_pt_avg_unconsolidated.csv', 'a+', newline=''))
        df_county_input = self.df_input[self.df_input['COUNTY'] == county]
        for i, row in df_county_input.iterrows():
            f = join(self.data_folder, row['MONITORING_DATA_DOWNLOADED'])
            self.npdes_dfs(f)
        # df_all = pd.concat(self.df_list, axis=1).sort_index()
        # df_interp = df_all.interpolate(limit_area='inside')
        # df_interp = df_interp[(df_interp.index.year < 2020)]
        # df_interp.to_csv('test.csv')
        print('*'*75)
        for df in self.df_list:
            df = df.resample('M')
            df_interp = df.interpolate(limit_area='inside')
            print(df_interp)
            df_interp = df_interp[(df_interp.index.year < 2020)]
            id = df_interp.columns[0]
            total = df_interp[df_interp.columns[0]].sum()
            avg_annual = total / 5
            avg_monthly = total / 60
            w.writerow([id[:9], id, county, str(total), str(avg_annual), str(avg_monthly)])

    def aggregate_averages(self):
        df = pd.read_csv('discharge_pt_avg_unconsolidated.csv')
        # print(df)
        df_grp = df.groupby(['NPDESID'], as_index=False).agg({'ANNUAL_AVG': sum,
                                                              'MONTHLY_AVG': sum,
                                                              'COUNTY': 'first',
                                                              'TOTAL_FLOW': 'mean',
                                                              'NPDESID_SEQ': 'max'})
        df_grp.to_csv('discharge_pt_avg_grouped.csv', index=False)

    def investigate_large_discharges(self):
        self.df_list = []
        df = pd.read_excel('_WWTP_download_tool_2020.04.10.xlsm', sheet_name='TCEQ_Outfalls')
        df_large = df[df['SUSPECT_BAD_DATA_FLAG']]
        for i, row in df_large.iterrows():
            f = join(self.data_folder, row['MONITORING_DATA_DOWNLOADED'])
            self.npdes_dfs(f)
        for df in self.df_list:
            # df = df.resample('M')
            themin, themax, themean, themedian = df[df.columns[0]].min(), df[df.columns[0]].max(), \
                                                 df[df.columns[0]].mean(), df[df.columns[0]].median()
            print(df)
            print('npdes:{}, min={}, max={}, mean={}, median={}'.format(df.columns[0], themin, themax, themean, themedian))
            print('*'*150)


if __name__ == "__main__":
    ddp = DischargeDataProcessor()
    ddp.set_basins_list(basin_list=['San Antonio', 'Guadalupe'])
    ddp.set_params(num_months=150, non_null_ratio_target=0.10)
    # #### Process files ####
    ddp.process_files()
    # #### calc averages ####
    # df_counties = pd.read_excel('_WWTP_download_tool.xlsm', sheet_name='COUNTIES')
    # df_counties_download = df_counties[df_counties['PROCESS_FLAG']]
    # print(df_counties_download)
    # for i, row in df_counties_download.iterrows():
    #     ddp.calc_averages(county=row['COUNTY'])
    # #### aggregate averages ####
    # ddp.aggregate_averages()
    # #### investigate unusually large discharges ####
    # ddp.investigate_large_discharges()
