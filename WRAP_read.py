import pandas as pd
import math
import re
import msgbox


def read_unappropriated_flow(basin, wrap_out_file):
    # file_in = r'C:\Users\BThaman\OneDrive - HDR, Inc\Documents\TWDB\bwam3.out'
    # read file into list of lines
    with open(wrap_out_file, 'r') as f_in:
        lines = [line.strip() for line in f_in.readlines()]
    # get variables from 5th row
    yrst, nyrs, ncpts, nwrout, nreout = (int(x) for x in lines[4].split()[:5])
    # calc number of months and number of rows in a month block
    num_months = nyrs * 12
    mblock = nwrout + ncpts + nreout
    days_in_month = {1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31}
    dict_afpm = {}
    dict_cfs = {}
    # re_pattern = re.compile(r'(^FK.*|^FAKE.*|.*OSB$|^DUMM.*)', re.IGNORECASE)
    re_pattern = re.compile(r'(^FK.*|^FAK.*|^DUMM.*|^ENVCAP.*)', re.IGNORECASE)
    # re_pattern = re.compile(r'^ZZZZZ.*', re.IGNORECASE)

    for m in range(num_months):
        year = yrst + math.floor((m + 1) / 12) if (m + 1) % 12 > 0 else yrst + math.floor((m + 1) / 12) - 1
        month = (m + 1) % 12 if (m + 1) % 12 > 0 else 12
        cpt_1 = 5 + m * mblock + nwrout + 1
        cpt_n = cpt_1 + ncpts - 1
        # print(str(year) + '-' + str(month) + ': ' + str(cpt_1))
        if m == 0:
            dict_afpm['Date'] = [str(month) + '/' + str(year)]
            dict_cfs['Date'] = [str(month) + '/' + str(year)]
        else:
            dict_afpm['Date'].append(str(month) + '/' + str(year))
            dict_cfs['Date'].append(str(month) + '/' + str(year))
        # loop CPs for current month
        for i in range(cpt_1 - 1, cpt_n):
            # cpt_id = str(lines[i].split()[0])
            cpt_id = str(lines[i][:6])
            if re_pattern.match(cpt_id):
                continue
            # q_avail_afpm = float(lines[i].split()[6])
            try:
                q_avail_afpm = float(lines[i][61:72])
                q_avail_cfs = q_avail_afpm / (1.9835 * days_in_month[month])
            except Exception as e:
                q_avail_afpm = 0
                q_avail_cfs = 0
            if m == 0:
                dict_afpm[cpt_id] = [q_avail_afpm]
                dict_cfs[cpt_id] = [q_avail_cfs]
            else:
                dict_afpm[cpt_id].append(q_avail_afpm)
                dict_cfs[cpt_id].append(q_avail_cfs)
    df_afpm = pd.DataFrame(dict_afpm)
    df_cfs = pd.DataFrame(dict_cfs)
    # print('\n' + str(len(df_afpm.index)))
    # print('\n')
    # print(df_afpm.head())
    df_afpm.to_csv(basin + '_afpm.csv', index=False)
    df_cfs.to_csv(basin + '_cfs.csv', index=False)


if __name__ == '__main__':
    pth = 'bwam3.out'
    basin = 'brazos'
    read_unappropriated_flow(basin=basin, wrap_out_file=pth)
