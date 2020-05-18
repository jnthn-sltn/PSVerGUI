# -*- coding: utf-8 -*-
"""
Created on Wed Apr 29 16:17:20 2020.

@author: joslaton
"""
import PySimpleGUI as sg
import pandas as pd
import sys
from string import ascii_uppercase as uc
###################################################################################################
#   Command Sequence Gen  #########################################################################
###################################################################################################


# Constants
rffe_dict = {
    'PM_TRIG': 28,
    'UDR_RST': 35,
    'ERR_SUM': 36,
    'BUS_LD': 43,
    'TEST_PATT': 44,
    'EXT_TRIG_MASK_A': 45,
    'EXT_TRIG_REG_A': 46,
    'EXT_TRIG_REG_B': 47,
    'EXT_TRIG_MASK_B': 48,
    'EXT_TRIG_CNT_11_H': 49,
    'EXT_TRIG_CNT_12_H': 50,
    'EXT_TRIG_CNT_13_H': 51,
    'EXT_TRIG_CNT_14_H': 52,
    'EXT_TRIG_CNT_15_H': 53,
    'EXT_TRIG_CNT_16_H': 54,
    'EXT_TRIG_CNT_17_H': 55,
    'EXT_TRIG_CNT_3_H': 56,
    'EXT_TRIG_CNT_4_H': 57,
    'EXT_TRIG_CNT_5_H': 58,
    'EXT_TRIG_CNT_6_H': 59,
    'EXT_TRIG_CNT_7_H': 60,
    'EXT_TRIG_CNT_8_H': 61,
    'EXT_TRIG_CNT_9_H': 62,
    'EXT_TRIG_CNT_10_H': 63,
    'SIREV_ID': 164}

# Functions
def cmd_str_generator(enabled, typ, usid, regaddr, wrt_msk, reg_wrt_data, exp_rd_data, regwrary=[]):
    """
    Send a Command.

    Parameters
    ----------
    enabled : int
        Toggle bit to determine if Command is sent.
    typ : int
        Controls the type of Command.
        1 => Standard Write Command
        2 => Standard Read Command
        3 => Extended Write Command
        4 => Extended Read Command
        5 => Masked Write Command
    usid : int
        Sets the USID for the Command.
    regaddr : int
        Sets the address for the Command.
    wrt_msk : int
        Sets the write mask for a Masked Write Command.
    reg_wrt_data : int
        Sets the data for a Write Command.
    exp_rd_data : int
        Sets the expected data for a Read Command.
    regwrary : list, optional
        Sets the data for an Extended Write Array Command. The default is [].

    Returns
    -------
    list
        DESCRIPTION.

    """
    cmd_str = str(enabled) + ','
    cmd_str += str(typ) + ','
    cmd_str += str(usid) + ','
    cmd_str += str(regaddr) + ','
    cmd_str += str(wrt_msk) + ','
    cmd_str += str(reg_wrt_data) + ','
    cmd_str += str(exp_rd_data) + '\n'
    if regwrary:
        assert type(regwrary) == list
        assert len(regwrary) < 5
        cmd_str = cmd_str.strip('\n') + ','
        for val in regwrary:
            cmd_str += str(val) + ','
        cmd_str += '\n'
    return [cmd_str]


def read_cmd_generator(adrs, usid, erd):
    """
    Send a Standard Read Command .

    Parameters
    ----------
    adrs : int
        Set the Address for the command.
    usid : int
        Set the USID for the command.
    erd : int
        Expected read data in decimal.

    Returns
    -------
    list
        list contains a command string.

    """
    return cmd_str_generator(1, 2, usid, adrs, 0, 0, erd)


def extend_read_cmd_generator(adrs, usid, erd):
    """
    Send an Extended Read Command.

    Parameters
    ----------
    adrs : int
        Set the Address for the command.
    usid : int
        Set the USID for the command.
    erd : int
        Expected read data in decimal.

    Returns
    -------
    list
        list contains a command string.

    """
    return cmd_str_generator(1, 4, usid, adrs, 0, 0, erd)


def write_cmd_generator(adrs, usid, wrt):
    """
    Send a Standard Write Command.

    Parameters
    ----------
    adrs : int
        Set the Address for the command.
    usid : int
        Set the USID for the command.
    wrt : TYPE
        DESCRIPTION.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    return cmd_str_generator(1, 1, usid, adrs, 0, wrt, 0)


def extend_write_cmd_generator(adrs, usid, wrt, ary=[]):
    """
    Send an Extended Write Command.

    Parameters
    ----------
    adrs : int
        Set the Address for the command.
    usid : int
        Set the USID for the command.
    wrt : TYPE
        DESCRIPTION.
    ary : TYPE, optional
        DESCRIPTION. The default is [].

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    typ = 3
    if ary:
        typ = '1' + str(len(ary))
    return cmd_str_generator(1, typ, usid, adrs, 0, wrt, 0,
                             regwrary=ary)


def msk_wrt_cmd_generator(adrs, usid, msk, wrt):
    """
    Send a Masked Write Command.

    Parameters
    ----------
    reg : int
        Set the Address for the command.
    usid : int
        Set the USID for the command.
    msk : TYPE
        DESCRIPTION.
    wrt : TYPE
        DESCRIPTION.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    return cmd_str_generator(1, 5, usid, adrs, msk, wrt, 0)


def check_err_cmd_generator(usid, exp_read_err=False):
    """
    Read the ERROR SUM Register.

    Parameters
    ----------
    usid : int
        Set the USID for the command.
    exp_read_err : TYPE, optional
        DESCRIPTION. The default is False.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    if exp_read_err:
        return cmd_str_generator(1, 2, usid, rffe_dict['ERR_SUM'], 0, 0, 4)
    return cmd_str_generator(1, 2, usid, rffe_dict['ERR_SUM'], 0, 0, -1)


def pwr_rst_cmd_generator(usid, useusid=False):
    """
    Send a Powered Reset Command.

    Parameters
    ----------
    usid : int
        Set the USID for the command.
    useusid : TYPE, optional
        DESCRIPTION. The default is False.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    if useusid:
        return write_cmd_generator(rffe_dict['PM_TRIG'], usid, 64)
    return write_cmd_generator(rffe_dict['PM_TRIG'], 0, 64)


def udr_rst_cmd_generator(usid):
    """
    Send a UDR Reset Command.

    Parameters
    ----------
    usid : int
        Set the USID for the command.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    return write_cmd_generator(rffe_dict['UDR_RST'], usid, 128)


def trig_all(usid):
    """
    Send a Trigger All command.

    Parameters
    ----------
    usid : int
        Set the USID for the command.

    Returns
    -------
    cmd_lst : TYPE
        DESCRIPTION.

    """
    cmd_lst = write_cmd_generator(rffe_dict['PM_TRIG'], usid, 7)
    cmd_lst += write_cmd_generator(rffe_dict['EXT_TRIG_REG_A'], usid, 255)
    cmd_lst += write_cmd_generator(rffe_dict['EXT_TRIG_REG_B'], usid, 255)
    return cmd_lst


def pm_trig_trigger_cmd_generator(usid, trigger):
    """
    Send a PM Trigger command to the specified trigger.

    Parameters
    ----------
    usid : int
        Set the USID for the command.
    trigger : TYPE
        DESCRIPTION.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    return write_cmd_generator(rffe_dict['PM_TRIG'], usid, 2**trigger)


def ext_trig_trigger_cmd_generator(usid, trigger):
    """
    Send an Extended Trigger command to the specified trigger.

    Parameters
    ----------
    usid : int
        Set the USID for the command.
    trigger : TYPE
        DESCRIPTION.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    if trigger < 11:
        ext_trig_reg = 'EXT_TRIG_REG_A'
        trigger = 2**(trigger-3)
    else:
        ext_trig_reg = 'EXT_TRIG_REG_B'
        trigger = 2**(trigger-11)
    return write_cmd_generator(rffe_dict[ext_trig_reg], usid, trigger)


def rm_pm_trig_msk_cmd_generator(usid):
    """
    Set the PM Trigger Mask to 0.

    Parameters
    ----------
    usid : int
        Set the USID for the command.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    return write_cmd_generator(rffe_dict['PM_TRIG'], usid, 0)


def rm_ext_trig_msk_a_cmd_generator(usid):
    """
    Set the Extended Trigger A Mask to 0.

    Parameters
    ----------
    usid : int
        Set the USID for the command.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    return write_cmd_generator(rffe_dict['EXT_TRIG_MASK_A'], usid, 0)


def rm_ext_trig_msk_b_cmd_generator(usid):
    """
    Set the Extended Trigger B Mask to 0.

    Parameters
    ----------
    usid : int
        Set the USID for the command.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    return write_cmd_generator(rffe_dict['EXT_TRIG_MASK_B'], usid, 0)


def set_pm_trig_mask_cmd_generator(usid, mask='all'):
    """
    Set the PM Trig Mask.

    Parameters
    ----------
    usid : int
        Set the USID for the command.
    mask : TYPE, optional
        DESCRIPTION. The default is 'all'.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    return write_cmd_generator(rffe_dict['PM_TRIG'], usid, 56)


def set_ext_trig_mask_a_cmd_generator(usid):
    """
    Set the Extended Trigger A Mask to mask.

    Parameters
    ----------
    usid : int
        Set the USID for the command.
    mask : TYPE, optional
        DESCRIPTION. The default is 'all'.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    return write_cmd_generator(rffe_dict['EXT_TRIG_MASK_A'], usid, 255)


def set_ext_trig_mask_b_cmd_generator(usid):
    """
    Set the Extended Trigger B Mask to mask.

    Parameters
    ----------
    usid : int
        Set the USID for the command.
    mask : TYPE, optional
        DESCRIPTION. The default is 'all'.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    return write_cmd_generator(rffe_dict['EXT_TRIG_MASK_B'], usid, 255)


def trigger_cmd_generator(usid, trigger):
    """
    Catch all trigger command generator.

    Parameters
    ----------
    usid : int
        Set the USID for the command.
    trigger : TYPE
        DESCRIPTION.

    Returns
    -------
    function
        Passthrough function.

    """
    if trigger < 3:
        return pm_trig_trigger_cmd_generator(usid, trigger)
    else:
        return ext_trig_trigger_cmd_generator(usid, trigger)


###################################################################################################
#   Test Gen Functions   ##########################################################################
###################################################################################################
def validate_mt_vals(md, mv):
    """
    Test mv for uniqueness and completeness.

    Parameters
    ----------
    md : dict{dict}
        Dictionary for the mtrig control group registers.
    mv : dict
        Dictionary containing user input for the mtrig control group registers.
    
    Returns
    -------
    _ : Bool
        Returns True if mv information is unique and complete.

    """
    if '' in mv.values():
        return False 
    tmp_lst = [(mv[f'-{k[1]}R-'], mv[f'-{k[1]}N-']) for k in md.keys()]
    if len(tmp_lst) != len(set(tmp_lst)):
        return False
    return True

def update_mtrig_dict(md, mv):
    """
    Generate and prepopulate mTrig control registers dict.

    Parameters
    ----------
    md : dict{dict}
        Dictionary for the mtrig control group registers.
    mv : dict
        Dictionary containing user input for the mtrig control group registers.
    
    Returns
    -------
    md : dict{dict}
        Dictionary for the mtrig control group registers.

    """
    for k in mv:
        if 'R' in k:
            md[f'M{k[1]}']['Reg'] = mv[k]
        if 'N' in k:
            md[f'M{k[1]}']['U/L'] = mv[k]

    return md

def generate_mtrig_dict(regs, df):
    """
    Generate and prepopulate mTrig control registers dict.

    Parameters
    ----------
    regs : list[list]
        List containing the mTrig registers.
    df : DataFrame
        DataFrame containing the MRD.
    
    Returns
    -------
    md : dict{dict}
        Dictionary for the mtrig control group registers.

    """
    md = {}
    a = list(set([el[2] for el in regs]))
    b = df[df.Name.str.lower().str.contains('mtrig')][['Name', 'Address']].values.tolist()
    if b:
        for el in b:
            tmp = el[0].upper().split('_')
            pin = tmp.index('MTRIG') + 1
            i = 0
            if len(tmp[pin]) == 1:
                while ( ( pin + i ) < len(tmp)):
                    md[f'M{tmp[pin + i]}'] = {}
                    md[f'M{tmp[pin + i]}']['Reg'] = el[1]
                    md[f'M{tmp[pin + i]}']['U/L'] = 'U' if i else 'L'
                    i += 1
    for grp in a:
        if grp not in md.keys():
            md[grp] = {'Reg': '',
                               'U/L': ''}
    return md


def triggered_write_test(usid, adrs_lst, dv_lst, mtrig=False):
    """
    Test PM and Extended Triggers with this function.

    Parameters
    ----------
    usid : int
        Set the USID for the test.
    adrs_lst : list of ints
        Addresses to test.
    dv_lst : list of ints
        Default values list.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    cmd_lst = []
    for i in range(len(adrs_lst)):
        # Send address change header.
        cmd_lst += extend_write_cmd_generator(255, usid, 0)
        for trig in range(18):
            # Send trigger change header.
            cmd_lst += extend_write_cmd_generator(255, usid, 1)
            # Perform UDR Reset if not mtrig.
            if not mtrig:
                cmd_lst += udr_rst_cmd_generator(usid)
            # Set all trigger masking.
            cmd_lst += set_pm_trig_mask_cmd_generator(usid)
            cmd_lst += set_ext_trig_mask_a_cmd_generator(usid)
            cmd_lst += set_ext_trig_mask_b_cmd_generator(usid)
            # Verify default value is correct.
            cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, dv_lst[i])
            # Write 0xFF to the register.
            cmd_lst += extend_write_cmd_generator(adrs_lst[i], usid, 255)
            # Verify write occured.
            cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, 255)
            # Clear all trigger masks.
            cmd_lst += rm_pm_trig_msk_cmd_generator(usid)
            cmd_lst += rm_ext_trig_msk_a_cmd_generator(usid)
            cmd_lst += rm_ext_trig_msk_b_cmd_generator(usid)
            # Write 0x00 to the register.
            cmd_lst += extend_write_cmd_generator(adrs_lst[i], usid, 0)
            # Verify write hasnt occurred.
            cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, 0)
            # Trigger
            cmd_lst += trigger_cmd_generator(usid, trig)
            # Verify write has occurred.
            cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, 0)
    return cmd_lst


def tbyb_test(usid, adrs_lst, dv_lst):
    """
    Test all TBYB registers with this function.

    Parameters
    ----------
    usid : int
        Set the USID for the test.
    adrs_lst : list of ints
        Addresses to test.
    dv_lst : list of ints
        Default values list.

    Returns
    -------
    cmd_lst

    """
    cmd_lst = []
    for i in range(len(adrs_lst)):
        # Send address change header.
        cmd_lst += extend_write_cmd_generator(255, usid, 2)
        # Perform PWR Reset.
        cmd_lst += pwr_rst_cmd_generator(usid)
        # Set all trigger masking.
        cmd_lst += set_pm_trig_mask_cmd_generator(usid)
        cmd_lst += set_ext_trig_mask_a_cmd_generator(usid)
        cmd_lst += set_ext_trig_mask_b_cmd_generator(usid)
        # Verify default value is correct.
        cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, dv_lst[i])
        # Write 0xFF to the register.
        cmd_lst += extend_write_cmd_generator(adrs_lst[i], usid, 255)
        # Verify no write occured.
        cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, dv_lst[i])
        # Set TBYB Mode to active.
        cmd_lst += extend_write_cmd_generator(rffe_dict['SIREV_ID'], usid, 238)
        # Write 0xFF to the register.
        cmd_lst += extend_write_cmd_generator(adrs_lst[i], usid, 255)
        # Verify write occured.
        cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, 255)
        # Write 0x00 to the register.
        cmd_lst += extend_write_cmd_generator(adrs_lst[i], usid, 0)
        # Verify write occured.
        cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, 0)
        # Set TBYB Mode to inactive.
        cmd_lst += extend_write_cmd_generator(rffe_dict['SIREV_ID'], usid, 239)
        # Write 0xFF to the register.
        cmd_lst += extend_write_cmd_generator(adrs_lst[i], usid, 255)
        # Verify write did not occur.
        cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, 0)
    return cmd_lst


def trig_counter_test(usid, adrs_lst, dv_lst, mtrig=False):
    """
    Test PM and Extended Triggers with this function.

    Parameters
    ----------
    usid : int
        Set the USID for the test.
    adrs_lst : list of ints
        Addresses to test.
    dv_lst : list of ints
        Default values list.

    Returns
    -------
    cmd_lst

    """
    def set_clk_top_cmd_gen(usid, clk):
        return write_cmd_generator(clk, usid, 255)

    def read_clk_mult_cmd_gen(usid, clk):
        c = [244, 228, 211, 195, 178, 162, 145, 129, 112, 96,
             79, 63, 46, 30, 13, 0]
        cmd_lst = []
        for el in c:
            cmd_lst += read_cmd_generator(clk, usid, el)
        return cmd_lst

    cmd_lst = []
    if len(adrs_lst) < 1:
        return cmd_lst
    for i in range(len(adrs_lst)):
        # Send address change header.
        cmd_lst += extend_write_cmd_generator(255, usid, 3)
        for trig in range(3, 18):
            # Send trigger change header.
            cmd_lst += extend_write_cmd_generator(255, usid, 4)
            # Perform UDR Reset if not mtrig.
            if not mtrig:
                cmd_lst += udr_rst_cmd_generator(usid)
            # Set all trigger masking.
            cmd_lst += set_pm_trig_mask_cmd_generator(usid)
            cmd_lst += set_ext_trig_mask_a_cmd_generator(usid)
            cmd_lst += set_ext_trig_mask_b_cmd_generator(usid)
            # Verify default value is correct.
            cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, dv_lst[i])
            # Write 0xFF to the register.
            cmd_lst += extend_write_cmd_generator(adrs_lst[i], usid, 255)
            # Verify write occured.
            cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, 255)
            # Clear all trigger masks.
            cmd_lst += rm_pm_trig_msk_cmd_generator(usid)
            cmd_lst += rm_ext_trig_msk_a_cmd_generator(usid)
            cmd_lst += rm_ext_trig_msk_b_cmd_generator(usid)
            # Write 0x00 to the register.
            cmd_lst += extend_write_cmd_generator(adrs_lst[i], usid, 0)
            # Verify write hasnt occurred.
            cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, 0)
            # Set the clk to top value.
            clk = rffe_dict[f'EXT_TRIG_CNT_{trig}_H']
            cmd_lst += set_clk_top_cmd_gen(usid, clk)
            # Verify clock readback.
            cmd_lst += read_clk_mult_cmd_gen(usid, clk)
            # Verify write has occurred.
            cmd_lst += extend_read_cmd_generator(adrs_lst[i], usid, 0)
    return cmd_lst


def mtrig_test(usid, mgroup_lst, mtrig_dict):
    """
    Test all available mtrig grouped registers.

    Parameters
    ----------
    usid : TYPE
        DESCRIPTION.
    mgroup_lst : list
        List containing registers in the format [['Address', 'Value', 'Trig N']...].

    Returns
    -------
    cmd_lst : TYPE
        DESCRIPTION.

    """
    cmd_lst = []
    for reg in mgroup_lst:
        # Collect the register and mask for the mgroup
        mtrig_reg = mtrig_dict[reg[2]]['Reg']
        mtrig_mask = mtrig_dict[reg[2]]['U/L']
        mtrig_reg = int(mtrig_reg, 16)
        if mtrig_mask == 'L':
            mtrig_mask = 240
        elif mtrig_mask == 'U':
            mtrig_mask = 15
        # Send address change header
        cmd_lst += extend_write_cmd_generator(255, usid, 5)
        for i in range(15):
            # Send mtrig change header
            cmd_lst += extend_write_cmd_generator(255, usid, 6)
            if mtrig_mask == 240:
                val = i
            elif mtrig_mask == 15:
                val = i << 4
            # Change the mtrig group's trigger.
            cmd_lst += msk_wrt_cmd_generator(mtrig_reg, usid, mtrig_mask, val)
            cmd_lst += triggered_write_test(
                usid, [int(reg[0], 16)], [int(reg[1], 16)], mtrig=True)
            cmd_lst += trig_counter_test(
                usid, [int(reg[0], 16)], [int(reg[1], 16)], mtrig=True)
    return cmd_lst


def append_header():
    """
    Append the standard TestStand Sequence header.

    Returns
    -------
    hdr : list
        List of strings containing the standard header.

    """
    hdr = ['#,setup_header,test_system,Manual,,,\n']
    hdr += ['#,setup_header,version,1,,,\n']
    hdr += ['#,MipiVerification,clusters,,,,\n']
    hdr += ['enabled,type,usid,regAddr,writeMask,regWriteData,' +
            'expectedReadData,regWriteArrayItem0,regWriteArrayItem1,' +
            'regWriteArrayItem2,regWriteArrayItem3\n']
    return hdr


def get_trigger_registers(df, mtrig=False):
    """
    Get trigger enabled registers from the wider set.

    Parameters
    ----------
    df : DataFrame
        DataFrame containing the MRD.
    mtrig : Boolean, optional
        Boolean that determines whether mtrig registers are returned.
        The default is False.

    Returns
    -------
    t_lst : list
        List of lists pertaining to individual trigger registers.

    """
    t_lst = df[['Address', 'Value', 'Trig N']
               ][df['Trig N'] != '-'].values.tolist()
    if mtrig:
        t_lst = [el for el in t_lst if el[2]
                 in ['MA', 'MB', 'MC', 'MD', 'ME']]
    else:
        t_lst = [el for el in t_lst if el[2] not
                 in ['MA', 'MB', 'MC', 'MD', 'ME']]
    return t_lst


def get_tbyb_registers(df):
    """
    Get the TBYB registers from the wider set.

    Parameters
    ----------
    df : DataFrame
        DataFrame containing the MRD.

    Returns
    -------
    tbyb_lst : list
        List of lists pertaining to individual TBYB registers.

    """
    tbyb_lst = df.query(
        'ERW == "-" & ERR != "-" & Address > "0xA7"'
        )[['Name','Address', 'Value']].values.tolist()
    return tbyb_lst


def get_estimated_registers(key, df):
    """
    Choose registers from the wider set.

    Parameters
    ----------
    key: str
        Key for ed. Valid strings are "TBYB","Trig","TT", and "mTrig"
    df : DataFrame
        DataFrame containing the MRD.

    Returns
    -------
    possible : list
        List of lists pertaining to individual possible TBYB registers.
    estimated : list
        List of lists pertaining to individual presumed TBYB registers.

    """
    ed = {
        '-TBYB-': {
            'Q1': 'ERW == "-" & ERR != "-" & Address > "0xA7"',
            'Q2': 'Name != "" & Address > "0xA7"',
            'Cols': ['Name','Address','Value']},
        '-T-': {
            'Q1': '"T0" <= `Trig N` <= "T9"',
            'Q2': 'Name !="" & ("0x1C" > Address | Address > "0xA7") & `Trig N` == "-"',
            'Cols': ['Name','Address','Trig N','Value']},
        '-TT-': {
            'Q1': '"T0" <= `Trig N` <= "T9"',
            'Q2': 'Name !="" & ("0x1C" > Address | Address > "0xA7") & `Trig N` == "-"',
            'Cols': ['Name','Address','Trig N','Value']},
        '-MT-': {
            'Q1': '"MA" <= `Trig N` <= "MZ"',
            'Q2': 'Name !="" & ("0x1C" > Address | Address > "0xA7") & ((`Trig N` == "-") | ("T0" <= `Trig N` <= "T9"))',
            'Cols': ['Name','Address','Trig N','Value']}
        }
    estimated = df.query(ed[key]['Q1'])[ed[key]['Cols']].values.tolist()
    possible = df.query(ed[key]['Q2'])[ed[key]['Cols']].values.tolist()
    possible = [el for el in possible if el not in estimated]
    return estimated, possible


def get_usid(df):
    """
    Collect the USID from a DataFrame containing the MRD.

    Parameters
    ----------
    df : DataFrame
        DataFrame containing the MRD.

    Returns
    -------
    usid : int
        Integer value from 1 to 15
        
    """
    usid = df.query('Address == "0x1F"')[['Value']].values.tolist()
    usid = int(usid[0][0][-1], 16)
    return usid


def get_register_info(registers, key):
    """
    Collect address and default value information.

    Parameters
    ----------
    registers : list
        List of lists pertaining to individual registers.
    key : str
        key used to determine the offset for default values.
    Returns
    -------
    addresses : list
        List of integers containing the addresses of registers.
    default_values : list
        List of integers containing the default values of registers.

    """
    kd = {'-TBYB-': 2,
          '-T-': 3,
          '-TT-': 3,
          '-MT-': 3}
    if key == '-MT-':
        return None, None
    pos = kd[key]
    addresses = [int(el[1], 16) for el in registers]
    default_values = [int(el[pos], 16) for el in registers]
    return addresses, default_values


def get_commands(df, mtrig_dict, usid):
    """
    Parse df and generate command sequence for Test Stand sequence.

    Parameters
    ----------
    df : DataFrame
        DataFrame containing the MRD.

    Returns
    -------
    cmd_lst : list
        List of strings containing commands to Test Stand sequence.

    """
    # Collect information from MRD
    trigger_regs = get_trigger_registers(df)
    tbyb_regs = get_tbyb_registers(df)
    mtrig_regs = get_trigger_registers(df, mtrig=True)
    addresses, default_values = get_register_info(trigger_regs)
    tbyb_addresses, tbyb_default_values = get_register_info(tbyb_regs)
    # Construct command list
    cmd_lst = append_header()
    cmd_lst += triggered_write_test(usid, addresses, default_values)
    cmd_lst += tbyb_test(usid, tbyb_addresses, tbyb_default_values)
    cmd_lst += trig_counter_test(usid, addresses, default_values)
    cmd_lst += mtrig_test(usid, mtrig_regs, mtrig_dict)
    return cmd_lst


def get_test(key, usid, regs):
    """
    Get the test corresponding to the selected key.

    Parameters
    ----------
    key : str
        Key used to determine which test to generate.
    usid : int
        Integer decimal valued USID for testing purposes.
    regs: list[list]
        List of lists containing information on selected registers.
    """
    
    addresses, defaults = get_register_info(regs, key)

    if key == '-T-':
        return triggered_write_test(usid, addresses, defaults)
    elif key == '-TT-':
        return trig_counter_test(usid, addresses, defaults)
    elif key == '-TBYB-':
        return tbyb_test(usid, addresses, defaults)
    elif key == '-MT-':
        md = regs.pop()
        regs = [[el[1], el[3], el[2]] for el in regs]
        return mtrig_test(usid, regs, md)


def save_commands(cmd_lst, output_file):
    """
    Save the command list to output_file.

    Parameters
    ----------
    cmd_lst : list
        List of strings containing commands to Test Stand sequence.
    output_file : list
        Path of file to save commands in.

    Returns
    -------
    None.

    """
    if output_file.split('.')[-1] != 'csv':
        output_file += '.csv'
    with open(output_file, 'w') as f:
        f.writelines(cmd_lst)


def load_mrd():
    """
    Collect the filename, validate, and return a DataFrame containing the MRD.

    Parameters
    ----------
    None.

    Returns
    -------
    mrd_df : DataFrame
        DataFrame containing the MRD. Returns None if load or validation fails.
    """
    mrd_filename = sg.popup_get_file('Select MRD File', file_types=(("CSV Files","*.csv"),)) 
    
    if mrd_filename is None or mrd_filename == '':
        return
    if mrd_filename.split('.')[-1] != "csv":
        sg.popup(f'{mrd_filename} is not a csv file.')
        return
    try:
        mrd_df  = pd.read_csv(mrd_filename, na_filter=False)
    except:
        sg.popup(f'{mrd_filename} is not a properly formatted MRD file.')
        return
    check_cols = ['Name', 'Address', 'Value', 'D7', 'D6', 'D5', 'D4', 'D3', 'D2', 'D1', 'D0', 'Trig N', 'RZW', 'RW',
                  'ERW', 'MRW', 'RR', 'ERR', 'RST', 'BSID', 'GSID1', 'GSID2']
    cols = mrd_df.columns.values.tolist()
    
    if not all([el in cols for el in check_cols]):
        sg.popup(f'{mrd_filename} is not a properly formatted MRD file.')
        return
    
    return mrd_df


def choose_tests():
    """
    Collect the user input to determine which tests to generate.

    Parameters
    ----------
    None

    Returns
    -------
    test_values : dict
        Dictionary containing the selected tests.

    """
    frame_layout = [[sg.Checkbox('TBYB', key='-TBYB-')],
                    [sg.Checkbox('Trigger', key='-T-')],
                    [sg.Checkbox('Timed Trigger', key='-TT-')],
                    [sg.Checkbox('Mappable Trigger', key='-MT-')]]
    layout = [[sg.Frame('Select Tests', frame_layout)],
              [sg.Ok(size=(10, 1)), sg.Cancel(size=(10, 1))]]
    test_window = sg.Window('Test Generator', layout=layout)
    test_events, test_values = test_window.read()
    test_window.Close()
    if test_events is None or test_events == 'Cancel' or not any(test_values.values()):
        return None
    return test_values


def get_mgroup(reg):
    """
    Collect mTrig control group from user input.

    Parameters
    ----------
    reg : list
        List containing information on an mTrig register.
    
    Returns
    -------
    reg : list
        List containing information on an mTrig register.

    """
    layout = [[sg.Text(f'Mappable Trigger Group for Register{reg[1]}:'),sg.DropDown([el for el in uc], uc[0], key='-G-', size=(4, 1))],
              [sg.Stretch(), sg.Ok(size=(10, 1)), sg.Cancel(size=(10, 1)), sg.Stretch()]]
    mg_window = sg.Window('Get mGroup', layout=layout)
    mg_evt, mg_val = mg_window.read()
    mg_window.close()
    if mg_evt is None or mg_evt == 'Cancel':
        raise Exception(f'No mGroup selection made for Register{reg[1]}.')
    if mg_evt == 'Ok':
        reg[2] = f'M{mg_val["-G-"]}'
    return reg


def generate_mt_layout(md):
    col1 = [[sg.Text('mGroup')]]
    col2 = [[sg.Text('Control Register')]]
    col3 = [[sg.Text('Nibble')]]
    a = [f'0x{el:02X}' for el in range(256)]

    for k in md.keys():
        tt = f'Nibble used to control mGroup {k[1]}.\n"U" => [7:4]\n"L" => [3:0]'
        col1 += [[sg.Text(f'{k[1]}:')]]
        col2 += [[sg.DropDown(a, default_value=md[k]['Reg'], key=f'-{k[1]}R-', size=(6, 1))]]
        col3 += [[sg.DropDown(['U','L'], default_value=md[k]['U/L'], tooltip=tt, key=f'-{k[1]}N-', size=(4, 1))]]
    
    frame_layout = [[sg.Column(col1, element_justification='right'),
              sg.Column(col2, element_justification='center'),
              sg.Column(col3, element_justification='center')]]

    layout = [[sg.Frame('',layout=frame_layout)],
              [sg.Stretch(), sg.Ok(size=(10, 1)), sg.Cancel(size=(10, 1)), sg.Stretch()]]
    return layout

def get_mtrig_dict(regs, df):
    """
    Collect mTrig control registers dict.

    Parameters
    ----------
    regs : list[list]
        List containing the mTrig registers.
    df : DataFrame
        DataFrame containing the MRD.
    
    Returns
    -------
    regs : list[list + dict{dict}]
        List containing mtrig triggerable registers and Dictionary for the mtrig control group registers.

    """
    # Collect mtrig group of any regs for which an mTrig group is not yet assigned. 
    for i in range(len(regs)):
        if 'M' not in regs[i][2]:
            try:
                regs[i] = get_mgroup(regs[i])
            except Exception as e:
                sg.popup_ok(e[0])
                return None
    
    # Collect mapping of all mgroups
    mtrig_dict = generate_mtrig_dict(regs, df)
    layout = generate_mt_layout(mtrig_dict)
    mt_window = sg.Window('mTrig Group Config',layout=layout)
    
    while True:
        mt_evts, mt_vals = mt_window.read(timeout=100)
        if mt_evts is None or mt_evts == "Cancel":
            mt_window.close()
            return None
        if mt_evts == 'Ok':
            if validate_mt_vals(mtrig_dict, mt_vals):
                mtrig_dict = update_mtrig_dict(mtrig_dict, mt_vals)
                break
            else:
                sg.popup_ok('Error:\n\nEach mappable trigger group must be assigned a unique control register and nibble.\n\n')
                continue
            
    mt_window.close()
    return regs + [mtrig_dict]


def generate_gs_layout(key, t1_values, t2_values):
    """
    Generate the layout for the Register Selection window.

    Parameters
    ----------
    t1_values : list[list]
        List containing the presumed registers.
    t1_values : list[list]
        List containing the possible but not presumed registers.
    
    Returns
    -------
    layout : layout
        The layout for the Register Selection window.

    """
    mx_w = max([len(el[0]) for el in t1_values + t2_values]) + 4
    
    ld = {'-TBYB-': {'Headings': ["Name","Address"],
                    'Column Widths': [mx_w, 8],
                    'Title': "Select TBYB Registers"},
          '-T-': {'Headings': ["Name","Address","Trigger"],
                   'Column Widths': [mx_w, 8, 6],
                   'Title': "Select Triggered Registers"},
          '-TT-': {'Headings': ["Name","Address","Trigger"],
                   'Column Widths': [mx_w, 8, 6],
                   'Title': "Select Time Triggered Registers"},
          '-MT-': {'Headings': ["Name","Address","Trigger"],
                   'Column Widths': [mx_w, 8, 6],
                   'Title': "Select Mappable Triggered Registers"}
    }
    
    column1_layout = [[sg.Table(t1_values, headings=ld[key]['Headings'],col_widths=ld[key]['Column Widths'], auto_size_columns=False, justification="left", key='-T1-')]]
    column2_layout = [[sg.Button("Add\n<<<", size=(10, 2), key='-ADD-')],
                    [sg.Button("Remove\n>>>", size=(10, 2), key='-RMV-')]]
    column3_layout = [[sg.Table(t2_values, headings=ld[key]['Headings'],col_widths=ld[key]['Column Widths'], auto_size_columns=False, justification="left", key='-T2-')]]
    frame1_layout = [[sg.Col(column1_layout), sg.Col(column2_layout), sg.Col(column3_layout)]]
    layout = [[sg.Frame(ld[key]['Title'], frame1_layout)],
            [sg.Stretch(), sg.Ok(size=(10, 1)), sg.Cancel(size=(10, 1)), sg.Stretch()]]
    return layout


def table_updater(wndo, evt, vals, t1, t2):
    if evt == '-ADD-':
        if vals['-T2-']:
            vals['-T2-'].sort()
            vals['-T2-'].reverse()
            for idx in vals['-T2-']:
                val = t2.pop(idx)
                t1 = register_insert(t1, val)
            wndo.Element('-T1-').update(t1)
            wndo.Element('-T2-').update(t2)
    if evt == '-RMV-':
        if vals['-T1-']:
            vals['-T1-'].sort()
            vals['-T1-'].reverse()
            for idx in vals['-T1-']:
                val = t1.pop(idx)
                t2 = register_insert(t2, val)
            wndo.Element('-T1-').update(t1)
            wndo.Element('-T2-').update(t2)
    return wndo, evt, vals, t1, t2


def register_insert(reg_lst, val):
    """
    Insert a register entry in a list of registers.

    Parameters
    ----------
    reg_lst : [[str,str,str]]
        List of lists containing structured information on registers.
    val : [str,str,str]
        List of strings containing structured [<Name>, <Address>, <Value>] information on a register.
    
    Returns
    -------
    reg_lst : list[list[str]]
        Transformed list of lists containing information on registers.

    """
    if not reg_lst:
        reg_lst = [val]
        return reg_lst
    if val[1] < reg_lst[0][1]:
        reg_lst = [val] + reg_lst
        return reg_lst
    if val[1] > reg_lst[-1][1]:
        reg_lst = reg_lst + [val]
        return reg_lst
    for i in range(len(reg_lst) - 1):
        if reg_lst[i][1] < val[1] < reg_lst[i+1][1]:
            reg_lst = reg_lst[:i+1] + [val] + reg_lst[i+1:]
            return reg_lst


def get_selections(key, df):
    """
    Collect registers from user input and MRD.

    Parameters
    ----------
    df : DataFrame
        DataFrame containing the MRD.
    
    Returns
    -------
    regs : list[list[str]]
        List of lists containing information on selected registers as needed for testing.

    """
    regs = []
    sd = {'-TBYB-': 'TBYB Register Selection',
          '-T-': 'Triggered Register Selection',
          '-TT-': 'Time Triggered Register Selection',
          '-MT-': 'Mappable Triggered Register Selection'}

    t1_vals, t2_vals = get_estimated_registers(key, df)
    layout = generate_gs_layout(key, t1_vals, t2_vals)
    selection_window = sg.Window(sd[key], layout=layout)
    
    while True:
        events, values = selection_window.Read(timeout=100)
        if events is None or events == 'Cancel':
            selection_window.Close()
            break
        if events == 'Ok':
            regs = t1_vals
            selection_window.close()
            break
        selection_window, events, values, t1_vals, t2_vals = table_updater(
                selection_window, events, values, t1_vals, t2_vals)
    if key == '-MT-':
        regs = get_mtrig_dict(regs, df)
    return regs


def test_gendo():
    mrd_df = load_mrd()
    if mrd_df is None:
        return
    tv = choose_tests()
    if tv is None:
        return
    usid = get_usid(mrd_df)
    cmd_lst = append_header()
    for key in ['-T-', '-TBYB-', '-TT-', '-MT-']:
        if tv[key]:
            regs = get_selections(key, mrd_df)
            if regs:
                cmd_lst += get_test(key, usid, regs)
            else:
                return
    
    seq_filename = sg.popup_get_file('Save Test File As', save_as=True, file_types=(("CSV Files","*.csv"),))
    if seq_filename:
        save_commands(cmd_lst, seq_filename)
    return
 