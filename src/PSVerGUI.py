# -*- coding: utf-8 -*-
"""
Created on Tue Apr 21 11:43:33 2020.

@author: joslaton
"""
import PySimpleGUI as sg
from psv_mrd_gen import mrd_gendo
from psv_test_gen import test_gendo
from psv_report_gen import report_gendo

sg.theme('TanBlue')

windict = {'MRD Generator': mrd_gendo,
           'Test Generator': test_gendo,
           'Report Generator': report_gendo}

layout = [
    [sg.Frame('Launch Utility',[
        [sg.Button(button_text='MRD Generator', size=(30,1))],
        [sg.Button(button_text='Test Generator', size=(30,1))],
        [sg.Button(button_text='Report Generator', size=(30,1))],
        [sg.Button(button_text='Exit', size=(15,1))]],
        element_justification='center')
]
]

main_window = sg.Window('PSVerGUI', layout)


while True:
    main_events, main_values = main_window.Read(timeout=100)
    if main_events is None or main_events == 'Exit':
        break
    if main_events != '__TIMEOUT__':
        main_window.Disappear()
        windict[main_events]()
        main_window.Reappear()

main_window.close()