from selenium import webdriver
from selenium.webdriver.common.by import By
import xlwings as xw
import os
import time
import datetime
import re
import msgbox
import pandas as pd
import logger_handler
from os.path import join
import WRAP_read


def get_download_path():
    """Returns the default downloads path for linux or windows"""
    if os.name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return os.path.join(os.path.expanduser('~'), 'downloads')


class MasterWB:
    def __init__(self):
        self.wb_model = xw.Book.caller()
        self.app = xw.App(visible=False)
        self.working_dir = None
        self.set_working_dir()
        self.sht_active = self.wb_model.sheets.active
        self.sht_outfall = self.wb_model.sheets('TCEQ_Outfalls')
        self.sht_settings = self.wb_model.sheets('Settings')
        self.sht_wam = self.wb_model.sheets('WAM')
        self.outfall_rng = self.sht_outfall.range('A1').expand()
        self.df_outfall = self.outfall_rng.options(convert=pd.DataFrame, index=False).value
        self.start_month = self.sht_settings.range('B1').value
        self.url = None
        self.browser = None
        self.logger_debug = logger_handler.setup_logger(name='npdes_debug', log_file=join(os.getcwd(),'debug.log'))
        self.logger_exception = logger_handler.setup_logger(name='npdes_exception', log_file=join(os.getcwd(),'error.log'))


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

    @staticmethod
    def get_download_file(facility_id, dtbase):
        try:
            pattern = r'^NPDESMonitoringData_' + facility_id + r'.*\.xlsx$'
            lastfile = None
            size = None
            re_pattern = re.compile(pattern, re.IGNORECASE)
            for root, dirs, files in os.walk(get_download_path()):
                for basename in files:
                    if re_pattern.match(basename):
                        file = os.path.join(root, basename)
                        # msgbox.show_message('debug', file)
                        dtmod = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(os.path.getmtime(file)))
                        if dtmod > dtbase:
                            dtbase = dtmod
                            lastfile = file
                            size = os.path.getsize(lastfile)
            return lastfile, size
        except Exception as e:
            msgbox.show_error('get_download_file error', e)
            return None, -1

    def export_xlsx(self, facility_id, start_date):
        try:
            facility_id_input = self.browser.find_element(By.ID, 'permit_factsheets_parameters_p_npdes_id')
            # facility_id_input = self.browser.find_element('id', 'permit_factsheets_parameters_p_npdes_id')  # alternative to By.ID
            facility_id_input.clear()
            facility_id_input.send_keys(facility_id)
            # the start date input field is hidden, and the webdriver can't set the value for a hidden field.
            # the workaround is to execute a javascript command to change the value:
            self.browser.execute_script("document.getElementById('permit_factsheets_parameters_p_start_date').value='" +
                                   start_date + "'")
            # there are two submit buttons with class name = 'echo-search-btn': 
            # 1st one to download an excel file, and 2nd one to download a CSV
            # by calling .find_elements (plural), both are returned
            submit_buttons = self.browser.find_elements(By.CLASS_NAME, 'echo-search-btn')
            time.sleep(3)
            # click the 1st submit button subscripting the buttons (download excel file)
            submit_buttons[0].click()
            # give browser time to download file
            time.sleep(5)
        except Exception as e:
            self.logger_exception.exception('download error')
        finally:
            pass

    def iterate_months(self, id, dt_start):
        try:
            dtnow = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.export_xlsx(id, dt_start)
            out_file_info = self.get_download_file(id, dtnow)
            out_file = out_file_info[0]
            if out_file:
                return out_file
            else:
                return None
        except Exception as e:
            msgbox.show_error('iterate months error', e)

    def download(self):
        try:
            self.url = 'https://echo.epa.gov/trends/loading-tool/get-data/monitoring-data-download'  # url good as of 8/30/22
            self.browser = webdriver.Chrome()
            self.browser.set_page_load_timeout(360)
            self.browser.get(self.url)
            # filter where DOWNLOAD_FLAG = TRUE
            df_outfall_download = self.df_outfall[self.df_outfall['DOWNLOAD_FLAG']]
            num_download = len(df_outfall_download.index)
            if len(df_outfall_download.index) == 0:
                raise Exception('No rows were flagged for download.')
            dcnt = 0
            for index, row in df_outfall_download.iterrows():
                dcnt += 1
                if dcnt % 25 == 0:
                    self.wb_model.save()
                f_out = self.iterate_months(row['NPDES_ID'], self.start_month)
                if f_out:
                    print('{} of {}: {:5.1f}% - {}'.format(dcnt, num_download, dcnt / num_download * 100, f_out))
                    self.outfall_rng(index + 2, 7).value = '=hyperlink("' + f_out + '", "' + \
                                                            os.path.basename(f_out) + '")'
                    self.outfall_rng(index + 2, 6).value = False
            print('\n\nDone')
            # msgbox.show_message('Info', 'Done downloading ' + str(num_download) + ' files')
        except Exception as e:
            msgbox.show_error('download error', e)
        finally:
            self.browser.quit()
            self.wb_model.save()
            self.app.quit()

    def create_unappropriated_flow_csv(self):
        wam_rng = self.sht_wam.range('A1').expand()
        df_wam = wam_rng.options(convert=pd.DataFrame, index=False).value
        df_process = df_wam[df_wam['process_flag']]
        for i, row in df_process.iterrows():
            self.sht_wam.range('H1').value = 'Processing ' + row['Basin'] + ' basin'
            WRAP_read.read_unappropriated_flow(basin=row['Basin'], wrap_out_file=row['wam_out_file'])
        self.sht_wam.range('H1').value = 'Done'
        msgbox.show_message('Info', 'Done')


def download():
    wb = MasterWB()
    wb.download()


def read_unappropriated_flow():
    wb = MasterWB()
    wb.create_unappropriated_flow_csv()


if __name__ == '__main__':
    wb = MasterWB()
    wb.download()
