# -*- coding: utf-8 -*-
"""
Created on Wed Apr 29 16:17:20 2020.

@author: joslaton
"""

import PySimpleGUI as sg
import pandas as pd
import sys
from zipfile import ZipFile
from xmltodict import parse


###################################################################################################
#   MRD Gen Module   ##############################################################################
###################################################################################################


# Constant Definition
import_cols = {'standard': ['Register Class',
                            'Implementation Required',
                            'Register Address (Hex.)',
                            'Register name',
                            'Data Bits',
                            'Bit Name',
                            'Default',
                            'Broadcast Slave ID and'
                            ' Group Slave ID Support',
                            'Trigger Support',
                            'Active Trigger ',
                            'Extended Register R/W',
                            'Masked Write Support',
                            'R/W'],
               'extended': ['Register Address (Hex.)',
                            'Register Name',
                            'No. Bits',
                            'Data Bits',
                            'Function',
                            'Default',
                            'Triggered',
                            'Mask-Write Support',
                            'TBYB']}


# Class Definition


class Register:
    """
    The Register class holds all information needed for a register.

    Parameters
    ----------
    None

    """

    def __init__(self, address):
        """
        Instantiate the Register class.

        Parameters
        ----------
        address : STRING
            HEX STRING with the format '0x01', '0x0A' etc.

        Returns
        -------
        None.

        """
        self.d = {}
        self.d['Name'] = ''
        self.d['Address'] = address
        self.d['Value'] = '---'
        self.d['D7'] = '-'
        self.d['D6'] = '-'
        self.d['D5'] = '-'
        self.d['D4'] = '-'
        self.d['D3'] = '-'
        self.d['D2'] = '-'
        self.d['D1'] = '-'
        self.d['D0'] = '-'
        self.d['Trig N'] = '-'
        self.d['RZW'] = '-' if address != '0x00' else 'O'
        self.d['RW'] = '-'
        self.d['ERW'] = '-'
        self.d['MRW'] = '-'
        self.d['RR'] = '-'
        self.d['ERR'] = '-'
        self.d['RST'] = '-'
        self.d['BSID'] = '-'
        self.d['GSID1'] = '-'
        self.d['GSID2'] = '-'

    def set_name(self, df_slice):
        """
        Set the register name.

        Parameters
        ----------
        df_slice : dataframe
            A dataframe limited to the current address.

        Returns
        -------
        None.

        """
        if self.d['Address'] < '0x80':
            a = [el for el in df_slice['Register name'].tolist() if
                 type(el) == str]
            if a:
                self.d['Name'] = a[0]
        else:
            a = [el for el in df_slice['Register Name'].tolist() if
                 type(el) == str]
            if a:
                self.d['Name'] = a[0]

    def set_dv(self, df_slice):
        """
        Set the Value.

        Parameters
        ----------
        df_slice : dataframe
            A dataframe limited to the current address.

        Returns
        -------
        None.

        """
        dv = ''.join([el for el in df_slice['Default'].tolist() if
                      type(el) == str])
        if self.d['Address'] < '0x80':
            try:
                if dv[1] in ['x', 'X']:
                    self.d['Value'] = dv.replace('X', 'x')
                else:
                    dv = int(dv, 2)
                    self.d['Value'] = f'0x{dv:2X}'.replace(' ', '0')
            except ValueError:
                raise Exception(f"Register {self.d['Address']} has an invalid"
                                f" default value of \"{dv}\"."
                                "\nOnly numeric values are accepted."
                                " Please change the value in Excel and try again.")
        else:
            try:
                assert len(dv) == 4
                self.d['Value'] = dv.replace('X', 'x')
            except AssertionError:
                raise Exception(f"Register {self.d['Address']} has an invalid"
                                f" default value of \"{dv}\"."
                                "\nOnly numeric values are accepted."
                                " Please change the value in Excel and try again.")

    def set_trign(self, df_slice):
        """
        Set the trigger.

        Parameters
        ----------
        df_slice : dataframe
            A dataframe limited to the current address.

        Returns
        -------
        None.

        """
        if self.d['Address'] < '0x80':
            a = [el for el in df_slice['Active Trigger '].tolist() if
                 type(el) == str]
            if a:
                a = a[0]
                a = a[0] + a[-1] if a[0] != 'E' else 'T' + a[-1]
                self.d['Trig N'] = a
        else:
            a = [el for el in df_slice['Triggered'].tolist() if
                 type(el) == str]
            if a:
                if a[0] not in ['N', 'n', 'No']:
                    a = a[0]
                    a = a[0] + a[-1] if a[0] != 'E' else 'T' + a[-1]
                    self.d['Trig N'] = a

    def set_rw(self, df_slice):
        """
        Set the register write field.

        Parameters
        ----------
        df_slice : dataframe
            A dataframe limited to the current address.

        Returns
        -------
        None.

        """
        if self.d['Address'] < '0x80':
            a = [el for el in df_slice['R/W'].tolist() if type(el) == str]
            if all([any(['R/W' in a, 'W' in a]), self.d['Address'] < '0x20']):
                self.d['RW'] = 'O'
        else:
            pass

    def set_rr(self, df_slice):
        """
        Set the register read field.

        Parameters
        ----------
        df_slice : dataframe
            A dataframe limited to the current address.

        Returns
        -------
        None.

        """
        if self.d['Address'] < '0x20':
            a = [el for el in df_slice['R/W'].tolist() if type(el) == str]
            if any(['R/W' in a, 'R' in a]):
                self.d['RR'] = 'O'
        else:
            pass

    def set_erw(self, df_slice):
        """
        Set the extended register write field.

        Parameters
        ----------
        df_slice : dataframe
            A dataframe limited to the current address.

        Returns
        -------
        None.

        """
        if self.d['Address'] < '0x80':
            a = [el for el in df_slice['R/W'].tolist() if type(el) == str]
            b = [el for el in df_slice['Extended Register R/W'].tolist() if
                 type(el) == str]
            if all([any(['R/W' in a, 'W' in a]), 'Yes' in b]):
                self.d['ERW'] = 'O'
        else:
            a = [el for el in df_slice['TBYB'].tolist() if type(el) == str]
            b = list(set([el for el in df_slice['Register Name'].tolist() if
                          type(el) == str]))
            if all([a, b]):
                if all([a[0] in ['N', 'No'],
                        b[0].lower() not in ['sirev_id']]):
                    self.d['ERW'] = 'O'

    def set_err(self, df_slice):
        """
        Set the extended register read field.

        Parameters
        ----------
        df_slice : dataframe
            A dataframe limited to the current address.

        Returns
        -------
        None.

        """
        if self.d['Address'] < '0x80':
            a = [el for el in df_slice['R/W'].tolist() if type(el) == str]
            b = [el for el in df_slice['Extended Register R/W'].tolist() if
                 type(el) == str]
            if all([any(['R/W' in a, 'R' in a]), 'Yes' in b]):
                self.d['ERR'] = 'O'
        else:
            self.d['ERR'] = 'O'

    def set_mrw(self, df_slice):
        """
        Set the extended register write field.

        Parameters
        ----------
        df_slice : dataframe
            A dataframe limited to the current address.

        Returns
        -------
        None.

        """
        if self.d['Address'] < '0x80':
            a = [el for el in df_slice['R/W'].tolist() if type(el) == str]
            b = [el for el in df_slice['Masked Write Support'].tolist() if
                 type(el) == str]
            if all([any(['R/W' in a, 'W' in a]), 'Yes' in b]):
                self.d['MRW'] = 'O'
        else:
            a = [el for el in df_slice['TBYB'].tolist() if type(el) == str]
            if a:
                if a[0] in ['N', 'No']:
                    b = [el for el in df_slice['Mask-Write Support'] if
                         type(el) == str]
                    if b:
                        if b[0] not in ['N', 'No']:
                            self.d['MRW'] = 'O'

    def set_bgg(self, df_slice):
        """
        Set the BSID/GSID1/GSID2 write field.

        Parameters
        ----------
        df_slice : dataframe
            A dataframe limited to the current address.

        Returns
        -------
        None.

        """
        if self.d['Address'] < '0x80':
            a = [el for el in df_slice['R/W'].tolist() if type(el) == str]
            b = [el for el in df_slice[
                'Broadcast Slave ID and Group Slave ID Support'] if
                type(el) == str]
            if all([any(['R/W' in a, 'W' in a]), 'Yes' in b]):
                self.d['BSID'] = 'O'
                self.d['GSID1'] = 'O'
                self.d['GSID2'] = 'O'
        else:
            pass

    def set_rst(self, df_slice):
        """
        Set the RST field.

        Parameters
        ----------
        df_slice : dataframe
            A dataframe limited to the current address.

        Returns
        -------
        None.

        """
        if self.d['Address'] < '0x80':
            a = self.d['ERW'] == 'O'
            b = self.d['ERR'] == 'O'
            c = [el for el in df_slice['Register Class'].tolist() if
                 type(el) == str]
            d = 'RFFE Reserved'
            if all([a, b, d not in c]):
                self.d['RST'] = 'O'
        else:
            if self.d['ERW'] == 'O':
                self.d['RST'] = 'O'

    def set_bits(self, df_slice):
        """
        Set the DX fields.

        Parameters
        ----------
        df_slice : dataframe
            A dataframe limited to the current address.

        Returns
        -------
        None.

        """
        for i in range(8):
            self.d[f'D{i}'] = 'R'
        a = [el.strip('[]') for el in df_slice['Data Bits'].tolist() if
             type(el) == str]
        if self.d['Address'] < '0x80':
            b = [el for el in df_slice['Bit Name'].tolist() if
                 type(el) == str]
            c = [el for el in df_slice['R/W'].tolist() if type(el) == str]
            sb_dict = dict(zip(a, b))
            reserved_dict = {}
            # Expand the '7:5' style data bits until each bit is represented.
            for k in sb_dict.keys():
                if ':' in k:
                    hi_dex = int(k[0]) + 1
                    lo_dex = int(k[-1])
                    for i in range(lo_dex, hi_dex):
                        reserved_dict[f"{i}"] = sb_dict[k]
                else:
                    reserved_dict[k] = sb_dict[k]
            # Now must become boolean
            for k in reserved_dict.keys():
                reserved_dict[k] = 'reserved' in reserved_dict[k].lower()
                if 'R' in c:
                    reserved_dict[k] = True
                if reserved_dict[k]:
                    try:
                        val = int(self.d['Value'], 16)
                    except ValueError:
                        print(f"Default value at Register {self.d['Address']}"
                              " is invalid.\n")
                        sys.exit(1)
                    self.d[f'D{k}'] += f'{val:008b}'[-1::-1][int(k)]
                else:
                    self.d[f'D{k}'] += '/W'
        else:
            b = [el for el in df_slice['Function'].tolist() if
                 type(el) == str]
            c = [el for el in df_slice['TBYB'].tolist() if type(el) == str]
            sb_dict = dict(zip(a, b))
            # Need to expand because sometimes not all bits are implemented.
            reserved_dict = {'0': 'RESERVED', '1': 'RESERVED',
                             '2': 'RESERVED', '3': 'RESERVED',
                             '4': 'RESERVED', '5': 'RESERVED',
                             '6': 'RESERVED', '7': 'RESERVED'}
            # Expand the '7:5' style data bits until each bit is represented.
            for k in sb_dict.keys():
                if ':' in k:
                    hi_dex = int(k[0]) + 1
                    lo_dex = int(k[-1])
                    for i in range(lo_dex, hi_dex):
                        reserved_dict[f"{i}"] = sb_dict[k]
                else:
                    reserved_dict[k] = sb_dict[k]
            # Now must become boolean
            for k in reserved_dict.keys():
                reserved_dict[k] = any(['reserved' in reserved_dict[k].lower(),
                                        'y' in c[0].lower()])
                if reserved_dict[k]:
                    val = int(self.d['Value'], 16)
                    self.d[f'D{k}'] += f'{val:008b}'[-1::-1][int(k)]
                else:
                    self.d[f'D{k}'] += '/W'


# Function Definition
def get_sheet_ids(file_path):
    sheet_names = []
    with ZipFile(file_path, 'r') as zip_ref:
        xml = zip_ref.open(r'xl/workbook.xml').read()
        dictionary = parse(xml)

        if not isinstance(dictionary['workbook']['sheets']['sheet'], list):
            sheet_names.append(dictionary['workbook']['sheets']['sheet']['@name'])
        else:
            for sheet in dictionary['workbook']['sheets']['sheet']:
                sheet_names.append(sheet['@name'])
    return sheet_names


def get_register_sheets(sheet_names):
    frame1_layout = [[sg.Text('Standard Register Sheet', size=(20,1)), sg.Combo(sheet_names, default_value=sheet_names[0], size=(70, 1), key='-STD-', readonly=True)],
                    [sg.Text('Extended Register Sheet', size=(20,1)), sg.Combo(sheet_names, default_value=sheet_names[1], size=(70, 1), key='-EXT-', readonly=True)]]
    frame2_layout = [[sg.Ok(size=(10, 1)), sg.Cancel(size=(10, 1))]]
    layout = [[sg.Frame('', frame1_layout)],
              [sg.Frame('', frame2_layout)]]
    grs_window = sg.Window('MRD', layout=layout, element_justification='center')
    grs_events, grs_values = grs_window.Read()
    grs_window.Close()
    if grs_events is None or grs_events == 'Cancel':
        return 'Cancel', 'Cancel'
    return grs_values['-STD-'], grs_values['-EXT-']


def condition_df(df, ext=False):
    """
    Condition the dataframe to enforce formatting.

    Parameters
    ----------
    df : DataFrame
        DataFrame containing raw, unconditioned data.
    ext : Boolean, optional
        Flag for pSemi extended register DataFrame. The default is False.

    Returns
    -------
    df : DataFrame
        DataFrame containing conditioned data.

    """
    df.dropna(axis=0, how='all', inplace=True)
    df.reset_index(drop=True)
    if ext:
        df[
           ['Register Address (Hex.)', 'Register Name']
          ] = df[
                 ['Register Address (Hex.)', 'Register Name']
                ].fillna(method='ffill')
    else:
        df[
           ['Register Address (Hex.)', 'Bit Name']
          ] = df[
                 ['Register Address (Hex.)', 'Bit Name']
                ].fillna(method='ffill')
    return df


def get_implemented_registers(s_df, e_df):
    """
    Collect the implemented registers.

    Parameters
    ----------
    s_df : DataFrame
        DataFrame containing information on the standard registers.
    e_df : DataFrame
        DataFrame containing information on the extended registers.

    Returns
    -------
    implemented_registers : list
        List containing the addresses of implemented registers.

    """
    standard_registers = list(sorted(set(
        s_df.query('`Implementation Required` == "Yes"'
                   ' or `Implementation Required` == "Optional"')[
                            'Register Address (Hex.)'])))
    standard_registers = [address for address in
                          standard_registers if len(address) == 4]

    extended_registers = list(sorted(set(e_df['Register Address (Hex.)'])))
    extended_registers = [address for address in
                          extended_registers if len(address) == 4]
    # Want to remove all of the registers that use N/A as default value.
    tmp_reg_lst = []
    for address in extended_registers:
        tmp_list = e_df['Default'][
            e_df['Register Address (Hex.)'] == address].tolist()
        tmp_list = [n for n in tmp_list if type(n) == str]
        if tmp_list:
            tmp_reg_lst += [address]
    extended_registers = tmp_reg_lst
    # Get all implemented registers
    implemented_registers = standard_registers + extended_registers
    return implemented_registers


def populate_register_fields(s_df, e_df, implemented_registers, register_pack):
    """
    Populate the register fields for implemented registers.

    Parameters
    ----------
    s_df : DataFrame
        DataFrame containing information on the standard registers.
    e_df : DataFrame
        DataFrame containing information on the extended registers.
    implemented_registers : list
        List containing the addresses of implemented registers.
    register_pack : list
        List containing Register objects for all registers.

    Returns
    -------
    register_pack : list
        List containing Register objects for all registers.

    """
    for address in implemented_registers:
        if address < '0x80':
            df_slice = s_df[
                s_df['Register Address (Hex.)'] == address
                ]
        else:
            df_slice = e_df[
                e_df['Register Address (Hex.)'] == address
                ]
        register_pack[address].set_name(df_slice)
        register_pack[address].set_dv(df_slice)
        register_pack[address].set_trign(df_slice)
        register_pack[address].set_rw(df_slice)
        register_pack[address].set_rr(df_slice)
        register_pack[address].set_erw(df_slice)
        register_pack[address].set_err(df_slice)
        register_pack[address].set_mrw(df_slice)
        register_pack[address].set_bgg(df_slice)
        register_pack[address].set_rst(df_slice)
        register_pack[address].set_bits(df_slice)
    return register_pack


def get_register_df(prd_file, standard_sheet, extended_sheet):
    """
    Import the register maps and construct the MRD DataFrame.

    Parameters
    ----------
    prd_file : str
        string path to file containing register map sheets.
    standard_sheet : str
        string name of standard register map sheet.
    extended_sheet : str
        string name of extended register map sheet.

    Returns
    -------
    register_df : DataFrame
        DataFrame containing MRD.
    usid : int
        Base 10 integer value of device USID.

    """
    # Import the PRD Register Maps.
    try:
        s_df = pd.read_excel(prd_file,
                            sheet_name=standard_sheet,
                            header=0,
                            usecols=import_cols['standard'])
    except Exception as e:
        print(e)
        raise Exception(f"Unable to import {standard_sheet} as standard register map.\n"
                        f"Please ensure that {standard_sheet} is properly formatted as a table.")
    try:
        e_df = pd.read_excel(prd_file,
                            sheet_name=extended_sheet,
                            header=0,
                            usecols=import_cols['extended'])
    except:
        raise Exception(f"Unable to import {extended_sheet} as extended register map.\n"
                        f"Please ensure that {extended_sheet} is properly formatted as a table.")    
    s_df = condition_df(s_df)
    e_df = condition_df(e_df, ext=True)
    implemented_registers = get_implemented_registers(s_df, e_df)
    # Initialize reg_lst
    reg_lst = [f'0x{el:02X}' for el in range(256)]
    # Create empty register dict
    register_pack = {el: Register(el) for el in reg_lst}
    # Populate register fields
    register_pack = populate_register_fields(s_df,
                                             e_df,
                                             implemented_registers,
                                             register_pack)
    # Construct output from register pack dictionaries
    register_df = pd.DataFrame([register.d for register
                                in register_pack.values()])
    # usid sets the USID used during testing.
    usid = register_df.query('Address == "0x1F"')['Value'].values.tolist()
    usid = usid[0][3]
    usid = int(usid, 16)
    return register_df, usid


def save_mrd(df, output_file):
    """
    Save the MRD DataFrame to the output_file path.

    Parameters
    ----------
    df : DataFrame
        DataFrame containing the MRD.
    output_file : str
        String path to which MRD will be saved.

    Returns
    -------
    None.

    """
    df.to_csv(path_or_buf=output_file, index=False)


def mrd_gendo():
    prd_filename = sg.popup_get_file('Select file with register tables',file_types=(("Excel Files","*.xlsx"),))
    if prd_filename is None or prd_filename == '':
        return
    if prd_filename.split('.')[-1] != "xlsx":
        sg.popup(f'{prd_filename} is not an XLSX file.')
        return
    sheetnames = get_sheet_ids(prd_filename)
    if len(sheetnames) < 2:
        sg.popup('File does not meet requirements. At least two sheets are required.')
        return
    std_sheet, ext_sheet = get_register_sheets(sheetnames)
    if std_sheet == ext_sheet:
        if std_sheet == 'Cancel':
            return
        sg.popup('Standard Sheet and Extended sheet cannot be the same selections.')
        return
    try:
        register_df, _ = get_register_df(prd_filename, std_sheet, ext_sheet)
    except Exception as e:
        sg.popup(e.args[0])
        return
    mrd_filename = sg.popup_get_file('Save MRD As', save_as=True, file_types=(("CSV Files","*.csv"),))
    if mrd_filename is None or mrd_filename == '':
        return
    save_mrd(register_df, mrd_filename)

