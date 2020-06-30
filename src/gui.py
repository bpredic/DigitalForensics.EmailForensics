import math
import PySimpleGUI as sg
from datetime import datetime
from mail_tool import MailTool
import matplotlib.pyplot as plt

truncate_limit = 30

def slice_dict(dict):
    items = {k: dict[k] for k in list(dict)[:truncate_limit]}
    print(items)
    return items

def print_dict(dict):
    for key in dict:
        print(key, dict[key])

def draw_bar_chart(dict, title):
    plt.style.use('ggplot')

    labels = dict.keys()
    values = dict.values()

    x_pos = [i for i, _ in enumerate(labels)]
    y_pos = [i for i, _ in enumerate(range(0, math.ceil(max(values))+1, 1))]
    plt.ion()
    plt.barh(x_pos, values, color='green', align='edge')
    plt.title(title)
    plt.yticks(x_pos, labels)
    plt.xticks(y_pos, y_pos)
    plt.tight_layout()
    fig = plt.gcf()
    plt.show()

sg.theme('DarkAmber')
layout = [[sg.InputText(key="CountMonthlyInput", default_text='yyyy'), sg.Button('Count sent messages monthly')],
          [sg.InputText(key="CountDailyInput", default_text='yyyy-mm'), sg.Button('Count sent messages daily')],
          [sg.InputText(key="CountHourlyInput", default_text='yyyy-mm-dd'), sg.Button('Count sent messages hourly')],
          [sg.InputText(key="CountDomainStartInput", default_text='yyyy-mm-dd'), sg.InputText(key="CountDomainEndInput", default_text='yyyy-mm-dd'), sg.Button('Count sent messages by domain')],
          [sg.InputText(key="CountKeywordsStartInput", default_text='yyyy-mm-dd'), sg.InputText(key="CountKeywordsEndInput", default_text='yyyy-mm-dd'), sg.Button('Count most used keywords')],
          [sg.InputText(key="ContactStartInput", default_text='yyyy-mm-dd'), sg.InputText(key="ContactEndInput", default_text='yyyy-mm-dd'), sg.Button('Get contact interactions')]]

window = sg.Window('Email forensics tool', layout, force_toplevel=True, finalize=True)
tool = MailTool()
tool.connect()

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    else:
        try:
            if event == "Count sent messages monthly":
                year = datetime.strptime(values['CountMonthlyInput'], '%Y')
                dict = tool.count_sent_messages_monthly(year)
                print_dict(dict)
                dict = slice_dict(dict)
                draw_bar_chart(dict, "Count sent messages monthly")
            if event == "Count sent messages daily":
                month = datetime.strptime(values['CountDailyInput'], '%Y-%m')
                dict = tool.count_sent_messages_daily(month)
                print_dict(dict)
                dict = slice_dict(dict)
                draw_bar_chart(dict, "Count sent messages daily")
            if event == "Count sent messages hourly":
                day = datetime.strptime(values['CountHourlyInput'], '%Y-%m-%d')
                dict = tool.count_sent_messages_hourly(day)
                print_dict(dict)
                dict = slice_dict(dict)
                draw_bar_chart(dict, "Count sent messages hourly")
            if event == "Count sent messages by domain":
                start = datetime.strptime(values['CountDomainStartInput'], '%Y-%m-%d')
                end = datetime.strptime(values['CountDomainEndInput'], '%Y-%m-%d')
                dict = tool.count_sent_messages_by_domain(start, end)
                print_dict(dict)
                dict = slice_dict(dict)
                draw_bar_chart(dict, "Count sent messages by domain")
            if event == "Count most used keywords":
                start = datetime.strptime(values['CountKeywordsStartInput'], '%Y-%m-%d')
                end = datetime.strptime(values['CountKeywordsEndInput'], '%Y-%m-%d')
                dict = tool.count_most_used_keywords(start, end)
                print_dict(dict)
                dict = slice_dict(dict)
                draw_bar_chart(dict, "Count keywords in sent messages")
            if event == "Get contact interactions":
                start = datetime.strptime(values['ContactStartInput'], '%Y-%m-%d')
                end = datetime.strptime(values['ContactEndInput'], '%Y-%m-%d')
                dict = tool.get_contact_interaction_weights(start, end)
                print_dict(dict)
                dict = slice_dict(dict)
                draw_bar_chart(dict, "Contact interaction weights")
        except:
            print("Error")
window.close()
tool.disconnect()