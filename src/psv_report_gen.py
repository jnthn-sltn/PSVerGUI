# -*- coding: utf-8 -*-
"""
Created on Wed Apr 29 16:17:20 2020.

@author: joslaton
"""


import pandas as pd
import PySimpleGUI as sg
import os
import errno
import sys
import tempfile
from pathlib import Path
from string import ascii_uppercase as uc

###################################################################################################
#   Report Gen Functions   ########################################################################
###################################################################################################
def get_mtrigs_result(df):
    """
    Generate mtrigs dictionary from the test result df.

    Parameters
    ----------
    df : DataFrame
        DataFrame containing the test results
    
    Returns
    -------
    md : dict
        Dictionary containing the mappable trigger information.
    """
    def get_md(ml, mi):
        """
        Get dictionary containing the mappable trigger information from user input.

        Parameters
        ----------
        ml : [[str,str]]
            List with register, nibble pairs.
        mi : dict
            Dictionary for user input mGroup names.
        
        Returns
        -------
        md : dict
            Dictionary containing the mappable trigger information.
        """
        md = {}
        for i in range(len(ml)):   
            if ml[i][0] in md.keys():
                md[ml[i][0]][ml[i][1]] = mi[f'-{i}-']
            else:
                md[ml[i][0]] = {ml[i][1]: mi[f'-{i}-']}

        return md


    mtrigs = df.query('Type == "MASKED_WRITE"')[['Address','Mask']].values.astype(int).tolist()
    mtrigs = [(el[0], el[1]) for el in mtrigs]
    mtrigs = list(set(mtrigs))
    mtrigs = [[f'0x{el[0]:02X}', f'0x{el[1]:02X}'] for el in mtrigs]
    mtrigs.sort()
    
    col1_layout = [[sg.Text('Register')]]
    col2_layout = [[sg.Text('Nibble')]]
    col3_layout = [[sg.Text('mGroup')]]
    mnt = {'0xF0': '[3:0]', '0x0F': '[7:4]'}
    
    for i in range(len(mtrigs)):
        col1_layout += [[sg.Text(f'Reg{mtrigs[i][0]}')]]
        col2_layout += [[sg.Text(f'{mnt[mtrigs[i][1]]}')]]
        col3_layout += [[sg.DropDown([el for el in uc], size=(6, 1), key=f'-{i}-')]]
    
    layout = [[sg.Column(col1_layout), sg.Column(col2_layout), sg.Column(col3_layout)],
              [sg.Stretch(), sg.Ok(size=(10, 1)), sg.Cancel(size=(10, 1)), sg.Stretch()]]
    
    mt_window = sg.Window('mTrig Setup', layout=layout)

    while True:
        mt_evts, mt_vals = mt_window.read(timeout=100)
        if mt_evts is None or mt_evts == 'Cancel':
            mt_window.close()
            return None
        if mt_evts == 'Ok' and all(mt_vals.values()):
            md = get_md(mtrigs, mt_vals)
            mt_window.close()
            break
    return md 

def psv_save_excel(writer, output_path):
    """
    Save the excel output and use win32 to squash a format bug.

    Parameters
    ----------
    writer : TYPE
        DESCRIPTION.
    output_path : TYPE
        DESCRIPTION.

    Returns
    -------
    None.

    """
    import win32com.client

    writer.save()
    xl = win32com.client.DispatchEx("Excel.Application")
    wb = xl.workbooks.open(output_path)
    xl.Visible = False
    wb.Close(SaveChanges=1)
    xl.Quit()


def process_tbyb(df, trig_reg_ddict):
    """
    Process the TBYB df to generate result list.

    Parameters
    ----------
    df : pandas dataframe
        Contains all TBYB tests.

    Returns
    -------
    res_dict : dict of results
        Dictionary containing result of all TBYB tests.

    """
    # Get indices of test starts.
    indices = df.query('Address == 255 & Write == 2').index.tolist()
    adrs_offset = 5
    dv_offset = 5
    dv2_offset = 7
    nv1_offset = 10
    nv2_offset = 12
    fv_offset = 15
    res_list = []
    for idx in indices:
        res_dict = {}
        adrs = int(df['Address'].iloc[idx + adrs_offset])
        adrs = f'Reg0x{adrs:02X}'
        res_dict['Address'] = adrs
        dv = int(df['Read'].iloc[idx + dv_offset])
        dv2 = int(df['Read'].iloc[idx + dv2_offset])
        nv1 = int(df['Read'].iloc[idx + nv1_offset])
        nv2 = int(df['Read'].iloc[idx + nv2_offset])
        fv = int(df['Read'].iloc[idx + fv_offset])
        check1 = dv == dv2
        check2 = nv1 != dv
        check3 = nv2 != dv
        check4 = fv == nv2
        if all([check1, any([check2, check3]), check4]):
            res_dict['TBYB Result'] = 'O'
        else:
            res_dict['TBYB Result'] = '---'
        res_list += [res_dict]
    res_df = pd.DataFrame(res_list)
    debug_df = pd.DataFrame(res_list)
    return debug_df, res_df


def process_counters(df, trig_reg_ddict, mtrig=False):
    """
    Process the Counter df to generate result list.

    Each test frame begins with a write of 0x03 to Reg0xFF.
    The test frame will contain the results of each counter
    register on the register address associated with the test frame.

    Parameters
    ----------
    df : pandas dataframe
        Contains all Counter test frames.

    Returns
    -------
    debug_df : pandas dataframe
        Dataframe containing result of timed trigger tests
        with debug information.

    result_df: pandas dataframe
        Dataframe containing condensed result of timed trigger tests.

    """
    # mtrig trigger test cant reset so we need to change the offset.
    # If not mtrig, add 1
    mtrig = int(not mtrig)
    # Get indices of test frame starts.
    indices = df.query('Address == 255 & Write == 3').index.tolist()
    indices += [len(df)]
    adrs_offset = 5 + mtrig
    # dv_offset = 4 + mtrig      Works but unnecessary.
    nv1_offset = 6 + mtrig
    nv2_offset = 11 + mtrig
    count_offset = 13 + mtrig
    fv_offset = 29 + mtrig
    debug_dicts = []
    result_dicts = []
    for i in range(len(indices)-1):
        sub_df = df.iloc[indices[i]:indices[i+1]]
        sub_indices = sub_df.query(
            'Address == 255 & Write == 4').index.tolist()
        # get the address
        adrs = sub_df['Address'].loc[indices[i] + adrs_offset]
        adrs = f"0x{adrs:02X}"
        for sub_idx in sub_indices:
            tmp_dict = {'Register Address': adrs}
            # get the dv Works but unneccesary.
            # dv = sub_df['Read'].loc[sub_idx + dv_offset]
            # get the first new value
            nv1 = sub_df['Read'].loc[sub_idx + nv1_offset]
            # get the second new value
            nv2 = sub_df['Read'].loc[sub_idx + nv2_offset]
            # get the trigger counter register
            count_reg = sub_df['Address'].loc[sub_idx + count_offset]
            # get the count pass fail value
            c = sub_df['Read'].loc[
                sub_idx + count_offset:sub_idx + count_offset + 15
                ].values.tolist()
            c_chk = sub_df['Pass'].loc[
                sub_idx + count_offset:sub_idx + count_offset + 15]
            c_chk = all(c_chk)
            # get final value
            fv = sub_df['Read'].loc[sub_idx + fv_offset]
            fv = int(fv)
            # Check 1: Verify nv1 not 0
            chk1 = nv1 != 0
            # Check 2: Verify nv2 == nv1
            chk2 = nv2 == nv1
            # Check 3: Verify final value == 0
            chk3 = fv == 0
            count_reg = trig_reg_ddict[f'0x{count_reg:02X}']
            tmp_dict['Trigger Register'] = count_reg
            tmp_dict['Counter Test'] = 'O' if c_chk else '---'
            tmp_dict['Trigger Test'] = 'O' if all([chk1, chk2, chk3]) \
                else '---'
            debug_string = count_reg + ':\t'
            if chk1:
                debug_string += ""
            else:
                debug_string += "Write failed with all triggers unmasked.\t"
            if chk2:
                debug_string += ""
            else:
                debug_string += "Write occurred with all triggers masked.\t"
            if chk3:
                debug_string += ""
            else:
                debug_string += f"Final value of 0x{fv:02X}" + \
                    " is unexpected with write value of 0x00\t"
            if c_chk:
                debug_string += ''
            else:
                debug_string += 'Count failure. Count reads: '
                debug_string += '; '.join([f'0x{int(el):02X}' for el in c])
            tmp_dict['Debug'] = debug_string
            debug_dicts += [tmp_dict]
    debug_df = pd.DataFrame(debug_dicts)
    addresses = list(set(debug_df['Register Address'].values.tolist()))
    addresses.sort()
    for adrs in addresses:
        success_tc = debug_df.query(
            '`Register Address` == ' + f'"{adrs}"' +
            ' & `Trigger Test` == "O"'
                                    ).index.tolist()
        success_tc = debug_df['Trigger Register'].loc[
            success_tc].values.tolist()
        success_cnt = debug_df.query(
            '`Register Address` == ' + f'"{adrs}"' +
            ' & `Counter Test` == "O"'
                                    ).index.tolist()
        success_cnt = debug_df['Trigger Register'
                               ].loc[success_cnt].values.tolist()
        result_dicts += [{'Register Address': adrs,
                          'Successful Timed Triggers': '; '.join(success_tc),
                          'Successful Counter Readback': '; '.join(success_cnt)
                          }]
    result_df = pd.DataFrame(result_dicts)
    return debug_df, result_df


def process_triggers(df, trig_reg_ddict, mtrig=False):
    """
    Process the trigger df to generate result list.

    Each test frame begins with a write of 0x00 to Reg0xFF.
    The test frame will contain the results of each trigger
    register on the register address associated with the test frame.

    Parameters
    ----------
    df : pandas dataframe
        Contains all trigger test frames.

    Returns
    -------
    debug_df : pandas dataframe
        Dataframe containing result of trigger tests with debug information.

    result_df: pandas dataframe
        Dataframe containing condensed result of trigger tests.

    """
    # Get indices of test frame starts.
    indices = df.query('Address == 255 & Write == 0').index.tolist()
    indices += [len(df)]
    # mtrig trigger test cant reset so we need to change the offset.
    # If not mtrig, add 1
    mtrig = int(not mtrig)
    adrs_offset = 5 + mtrig
    # Default value works but is unnecessary
    # dv_offset = 5
    nv1_offset = 6 + mtrig
    nv2_offset = 11 + mtrig
    trig_offset = 12 + mtrig
    fv_offset = 13 + mtrig
    debug_dicts = []
    result_dicts = []
    for i in range(len(indices)-1):
        sub_df = df.iloc[indices[i]:indices[i+1]]
        sub_indices = sub_df.query(
            'Address == 255 & Write == 1').index.tolist()
        # get the address
        adrs = sub_df['Address'].loc[indices[i] + adrs_offset]
        adrs = f"0x{adrs:02X}"
        for sub_idx in sub_indices:
            tmp_dict = {'Register Address': adrs}
            # get the dv
            # Default value works but is unnecessary
            # dv = sub_df['Read'].loc[sub_idx + dv_offset]
            # get the first new value
            nv1 = sub_df['Read'].loc[sub_idx + nv1_offset]
            # get the second new value
            nv2 = sub_df['Read'].loc[sub_idx + nv2_offset]
            # get the trigger register
            trig_reg = sub_df['Address'].loc[sub_idx + trig_offset]
            trig_reg = f'0x{int(trig_reg):02X}'
            # get the trig value
            trig_val = sub_df['Write'].loc[sub_idx + trig_offset]
            trig_val = f'0x{int(trig_val):02X}'
            # get the trig from the register and value
            trig = trig_reg_ddict[trig_reg][trig_val]
            # get final value
            fv = sub_df['Read'].loc[sub_idx + fv_offset]
            fv = int(fv)
            # Check 1: Verify nv1 not 0
            chk1 = nv1 != 0
            # Check 2: Verify nv2 == nv1
            chk2 = nv2 == nv1
            # Check 3: Verify final value == 0
            chk3 = fv == 0
            tmp_dict['Trigger Register'] = trig
            tmp_dict['Trigger Test'] = 'O' if all([chk1, chk2, chk3]) \
                else '---'
            debug_string = trig_reg + ':\t'
            if chk1:
                debug_string += ""
            else:
                debug_string += "Write failed with all triggers unmasked.\t"
            if chk2:
                debug_string += ""
            else:
                debug_string += "Write occurred with all triggers masked.\t"
            if chk3:
                debug_string += ""
            else:
                debug_string += f"Final value of 0x{fv:02X}" + \
                    " is unexpected with write value of 0x00\t"
            tmp_dict['Debug'] = debug_string
            debug_dicts += [tmp_dict]
    debug_df = pd.DataFrame(debug_dicts)
    addresses = list(set(debug_df['Register Address'].values.tolist()))
    addresses.sort()
    for adrs in addresses:
        success_tc = debug_df.query(
            '`Register Address` == ' + f'"{adrs}"' +
            ' & `Trigger Test` == "O"'
                                    ).index.tolist()
        success_tc = debug_df['Trigger Register'].loc[
            success_tc].values.tolist()
        result_dicts += [{'Register Address': adrs,
                          'Successful Triggers': '; '.join(success_tc)}]
    result_df = pd.DataFrame(result_dicts)
    return debug_df, result_df


def process_mtrigs(df, trig_reg_ddict):
    """
    Process the mtrig df to generate result list.

    Each test frame begins with a write of 0x05 to Reg0xFF.
    The test frame will contain the results of each trigger
    register and trigger counter on the register address associated
    with the test frame.

    Parameters
    ----------
    df : pandas dataframe
        Contains all mtrig test frames.

    Returns
    -------
    debug_df : pandas dataframe
        Dataframe containing result of mtrig tests with debug information.

    result_df: pandas dataframe
        Dataframe containing condensed result of mtrig tests.

    """
    # Get indices of test frame starts.
    indices = df.query('Address == 255 & Write == 5').index.tolist()
    indices += [len(df)]
    # Offsets referenced from indices
    adrs_offset = 8
    mtrig_offset = 2
    # Offsets referenced from sub_indices
    setting_offset = -1
    result_list = []
    debug_list = []
    for i in range(len(indices)-1):
        # Split off an mtrig frame
        # get the address
        adrs = df['Address'].loc[indices[i] + adrs_offset]
        adrs = f"0x{adrs:02X}"
        mtrig_adrs = df['Address'].loc[indices[i] + mtrig_offset]
        mtrig_adrs = f"0x{mtrig_adrs:02X}"
        msk = df['Mask'].loc[indices[i] + mtrig_offset]
        msk = f"0x{int(msk):02X}"
        mtrig_grp = trig_reg_ddict[mtrig_adrs][msk]
        # Get indices of each mtrig setting.
        sub_indices = df.iloc[indices[i]:indices[i+1]].query(
            'Address == 255 & Write == 6').index.tolist()
        sub_indices = [el+2 for el in sub_indices]
        sub_indices += [indices[i+1]]
        for j in range(len(sub_indices)-1):
            setting = df['Write'].loc[sub_indices[j] + setting_offset]
            if msk == '0xF0':
                setting = int(setting)
            else:
                setting = int(setting) >> 4
            setting = trig_reg_ddict['mtrig'][f'0x{setting:02X}']
            # print(df.iloc[sub_indices[j]:sub_indices[j+1]])
            trig_res_debug_df, trig_res_df = process_triggers(
                df.iloc[sub_indices[j]:sub_indices[j+1]
                        ].reset_index(drop=True), trig_reg_ddict, mtrig=True)
            tt_res_debug_df, tt_res_df = process_counters(
                df.iloc[sub_indices[j]:sub_indices[j+1]
                        ].reset_index(drop=True), trig_reg_ddict, mtrig=True)
            success_trigs = trig_res_df[
                'Successful Triggers'].values.tolist()[0]
            success_ttrigs = tt_res_df[
                'Successful Timed Triggers'].values.tolist()[0]
            success_count = tt_res_df[
                'Successful Counter Readback'].values.tolist()[0]
            tmp_dict = {'Register Address': adrs,
                        'mTrig Group': mtrig_grp,
                        'mTrig Setting': setting,
                        'Successful Triggers': success_trigs,
                        'Successful Timed Triggers': success_ttrigs,
                        'Successful Counter Readback': success_count}
            result_list += [tmp_dict]
            # Need to add mTrig group and mTrig setting
            trig_res_debug_df['mTrig Group'] = mtrig_grp
            tt_res_debug_df['mTrig Group'] = mtrig_grp
            trig_res_debug_df['mTrig Setting'] = setting
            tt_res_debug_df['mTrig Setting'] = setting
            trig_res_debug_df['Test Name'] = 'Trigger'
            tt_res_debug_df['Test Name'] = 'Timed Trigger'
            debug_list += [trig_res_debug_df]
            debug_list += [tt_res_debug_df]
    debug_df = pd.concat(debug_list, ignore_index=True, sort=False)
    debug_df = debug_df[['Register Address',
                         'mTrig Group',
                         'mTrig Setting',
                         'Trigger Register',
                         'Trigger Test',
                         'Counter Test',
                         'Debug',
                         'Test Name']]
    result_df = pd.DataFrame(result_list)
    # print(result_df)
    return debug_df, result_df


def psv_loadwriter(src_template_path, output_path):
    """
    Load the excel writer.

    Parameters
    ----------
    src_template_path : string
        Path to the excel template.
    output_path : string
        Path to which the output will be written.

    Returns
    -------
    writer : ExcelWriter object
        ExcelWriter object facilitates the writing of Excel documents.

    """
    from shutil import copyfile
    from openpyxl import load_workbook

    copyfile(src_template_path, output_path)
    writer = pd.ExcelWriter(output_path, engine='openpyxl')
    book = load_workbook(output_path)
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    return writer


def psv_loadfile(input_file):
    """
    Load and condition the result of TestStand sequence.

    Parameters
    ----------
    input_file : string
        Path to the csv file containing the test results.

    Returns
    -------
    df : Pandas DataFrame
        Dataframe containing the conditioned output.

    """
    header_dict = {' COND_TYPE': 'Type',
                   ' COND_REG_ADDR': 'Address',
                   ' COND_WRITE_MASK': 'Mask',
                   ' COND_WRITE_DATA': 'Write',
                   ' COND_EXPECTED_DATA': 'Expected',
                   ' MEAS_READ_DATA': 'Read',
                   ' MEAS_PARITY_OK': 'Parity',
                   ' MEAS_PASS': 'Pass'}
    df = pd.read_csv(input_file, header=8,
                     usecols=[' COND_TYPE',
                              ' COND_REG_ADDR',
                              ' COND_WRITE_MASK',
                              ' COND_WRITE_DATA',
                              ' COND_EXPECTED_DATA',
                              ' MEAS_READ_DATA',
                              ' MEAS_PARITY_OK',
                              ' MEAS_PASS'])
    df = df.rename(mapper=header_dict, axis=1)
    return df


def psv_processfile(result_df, writer, trig_reg_ddict, rsa_file, mrd_file):
    """
    Process the results.

    Parameters
    ----------
    result_df : DataFrame
        Dataframe containing the conditioned test result.
    writer : ExcelWriter
        Writer used to write to the Excel Sheets.
    trig_reg_ddict : Dictionary
        Multlevel dictionary containing register information.
    rsa_file : string
        Path to Register Status Analyzer output csv.
    mrd_file : string
        Path to Machine Readable PRD csv.

    Returns
    -------
    None.

    """
    from copy import deepcopy

    df_dict = {'Trigger Test': {'Query': 'Address == 255 & Write == 0',
                                'Function': process_triggers,
                                'input DF': pd.DataFrame([]),
                                'Debug DF': pd.DataFrame([]),
                                'Result DF': pd.DataFrame([])},
               'Timed Trigger Test': {'Query': 'Address == 255 & Write == 3',
                                      'Function': process_counters,
                                      'input DF': pd.DataFrame([]),
                                      'Debug DF': pd.DataFrame([]),
                                      'Result DF': pd.DataFrame([])},
               'TBYB Test': {'Query': 'Address == 255 & Write == 2',
                             'Function': process_tbyb,
                             'input DF': pd.DataFrame([]),
                             'Debug DF': pd.DataFrame([]),
                             'Result DF': pd.DataFrame([])},
               'mTrig Test': {'Query': 'Address == 255 & Write == 5',
                              'Function': process_mtrigs,
                              'input DF': pd.DataFrame([]),
                              'Debug DF': pd.DataFrame([]),
                              'Result DF': pd.DataFrame([])}}
    test_list = ['mTrig Test', 'Timed Trigger Test',
                 'TBYB Test', 'Trigger Test']

    for k in test_list:
        index = result_df.query(df_dict[k]['Query']).head(n=1)
        if len(index) > 0:
            index = index.index.tolist()[0]
            df_dict[k]['input DF'] = deepcopy(result_df.iloc[index:])
            df_dict[k]['input DF'].reset_index(drop=True, inplace=True)
            result_df = result_df.truncate(after=index-1)
            df_dict[k]['Debug DF'], df_dict[k]['Result DF'] = df_dict[
                k]['Function'](df_dict[k]['input DF'], trig_reg_ddict)
            df_dict[k]['Result DF'].to_excel(writer, sheet_name=k, index=False)
            if k != 'TBYB Test':
                df_dict[k]['Debug DF'].to_excel(writer,
                                                sheet_name=k + ' Debug',
                                                index=False)
    rsa = pd.read_csv(rsa_file)
    rsa.rename(columns={'Unnamed: 0': 'Address'}, inplace=True)
    mrd = pd.read_csv(mrd_file)
    rsa.to_excel(writer, sheet_name='RSA', index=False)
    mrd.to_excel(writer, sheet_name='MRD', index=False)


def get_report_fn():
    fn = sg.popup_get_file('Save Report As', title='Save As', save_as=True, file_types=(("XLSX Files","*.xlsx"),))
    if fn.endswith('.xlsx'):
        return fn
    elif fn:
        return fn + '.xlsx'
    return None


def get_tr_dict(df):
    md = get_mtrigs_result(df)
    if md is None:
        return
    tr_dict = {'0x1C': {'0x01': 'T0',
                                '0x02': 'T1',
                                '0x04': 'T2'},
                        '0x2E': {'0x01': 'T3',
                                '0x02': 'T4',
                                '0x04': 'T5',
                                '0x08': 'T6',
                                '0x10': 'T7',
                                '0x20': 'T8',
                                '0x40': 'T9',
                                '0x80': 'T10'},
                        '0x2F': {'0x01': 'T11',
                                '0x02': 'T12',
                                '0x04': 'T13',
                                '0x08': 'T14',
                                '0x10': 'T15',
                                '0x20': 'T16',
                                '0x40': 'T17',
                                '0x80': 'T18 - Not Implemented'},
                        '0x31': 'TC11',
                        '0x32': 'TC12',
                        '0x33': 'TC13',
                        '0x34': 'TC14',
                        '0x35': 'TC15',
                        '0x36': 'TC16',
                        '0x37': 'TC17',
                        '0x38': 'TC3',
                        '0x39': 'TC4',
                        '0x3A': 'TC5',
                        '0x3B': 'TC6',
                        '0x3C': 'TC7',
                        '0x3D': 'TC8',
                        '0x3E': 'TC9',
                        '0x3F': 'TC10',
                        'mtrig': {'0x00': 'T3',
                                    '0x01': 'T4',
                                    '0x02': 'T5',
                                    '0x03': 'T6',
                                    '0x04': 'T7',
                                    '0x05': 'T8',
                                    '0x06': 'T9',
                                    '0x07': 'T10',
                                    '0x08': 'T11',
                                    '0x09': 'T12',
                                    '0x0A': 'T13',
                                    '0x0B': 'T14',
                                    '0x0C': 'T15',
                                    '0x0D': 'T16',
                                    '0x0E': 'T17',
                                    '0x0F': 'MASKED'}
                        }
    for k in md.keys():
        tr_dict[k] = md[k]
    return tr_dict
    


def validate_init(vd):
    """
    Validate user input for the init options menu.

    Parameters
    ----------
    vd : dictionary
    Dictionary containing user input.

    Returns
    -------
    _ : Bool
    Returns True if input is valid and False if input is invalid.
    """
    input_lst = [vd['-RSA-'], vd['-MRD-'], vd['-RSLT-'], vd['-TMP-']]
    if not all(input_lst):
        return False
    if vd['-RSA-'].split('.')[-1] != 'csv':
        return False
    elif vd['-MRD-'].split('.')[-1] != 'csv':
        return False
    elif vd['-RSLT-'].split('.')[-1] != 'csv':
        return False
    elif vd['-TMP-'].split('.')[-1] != 'xlsx':
        return False
    elif len(input_lst) != len(set(input_lst)):
        return False
    else:
        return True

def get_report_input():
    """
    Collect valid input from the user.

    Parameters
    ----------

    None

    Returns
    -------
    _ : None or tuple
        Returns tuple containing filenames for the RSA, MRD, Test Stand Results and Template or returns None.
    """
    tmp_path = str(Path(__file__).parent.resolve().as_posix()) + "/Resources/TEMPLATE_Post-Silicon Verification.xlsx"
    layout = [[sg.Text('RSA File:',size=(10, 1), justification='right'), sg.InputText('',size=(70, 1), key='-RSA-'), sg.FileBrowse(file_types=(("CSV Files","*.csv"),))],
              [sg.Text('MRD File:',size=(10, 1), justification='right'), sg.InputText('',size=(70, 1), key='-MRD-'), sg.FileBrowse(file_types=(("CSV Files","*.csv"),))],
              [sg.Text('Result File:',size=(10, 1), justification='right'), sg.InputText('',size=(70, 1), key='-RSLT-'), sg.FileBrowse(file_types=(("CSV Files","*.csv"),))],
              [sg.Text('Template File:',size=(10, 1), justification='right'), sg.InputText(tmp_path,size=(70, 1), key='-TMP-'), sg.FileBrowse(file_types=(("XLSX Files","*.xlsx"),))],
              [sg.Stretch(), sg.Ok(size=(10, 1)), sg.Cancel(size=(10, 1)), sg.Stretch()]]

    ini_window = sg.Window("Report Generator",layout=layout)

    while True:
        ini_evts, ini_vals = ini_window.read(timeout=100)
        if ini_evts is None or ini_evts == 'Cancel':
            ini_window.close()
            return None
        if ini_evts == "Ok" and validate_init(ini_vals):
            ini_window.close()
            return ini_vals['-RSA-'], ini_vals['-MRD-'], ini_vals['-RSLT-'], ini_vals['-TMP-']
    


def report_gendo():
    # Collect RSA, MRD, Results and Template
    putin = get_report_input()
    if putin is None:
        return
    rsa_fn, mrd_fn, rslt_fn, template_fn = putin

    result_df = psv_loadfile(rslt_fn)

    # Populate trigger register dict
    tr_dict = get_tr_dict(result_df)
    if tr_dict is None:
        return
    # Load Results
    report_fn = get_report_fn()
    if report_fn is None:
        return
    try:

        writer = psv_loadwriter(template_fn, report_fn)

        psv_processfile(result_df, writer, tr_dict, rsa_fn, mrd_fn)

        psv_save_excel(writer, report_fn)

        sg.popup_ok('Report Generation Complete')
    except Exception as e:
        print(e)

    return
