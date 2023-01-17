import PySimpleGUI as sg
import subprocess
import os
import time
import threading
import queue
import win32com.client
import concurrent.futures

def ping(host):
    result = subprocess.run(
        'ping -n 1 {}'.format(host),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        shell=True
    )
    return "Reply from" in result.stdout.decode()

def ping_thread(hosts, results_queue, stop_event):
    with concurrent.futures.ThreadPoolExecutor() as executor:
        while not stop_event.is_set():
            futures = [executor.submit(ping, host) for host in hosts]
            results = [[host, future.result()] for host, future in zip(hosts, futures)]
            results_queue.put(results)
            time.sleep(1)  # delay for 1 second before pinging again


# Change the theme to 'DarkAmber'
sg.theme('DarkAmber')

layout = [
    [sg.Text('Enter host names or IP addresses separated by commas:')],
    [sg.Input(key='hosts')],
    [sg.Checkbox('Continuously ping', key='continuous_ping')],
    [sg.Button('Ping')],
    [sg.Button('New Outlook Email')],
    [sg.Text('Results:')],
    [sg.Table(values=[], headings=['Host', 'Status'], key='table', size=(60, 80), font='Garamond 12')]
]

window = sg.Window('USCIS TOC Concurrent Ping Utility',size=(600, 400)).layout(layout)

results_queue = queue.Queue()
thread = None
stop_event = threading.Event()
current_hosts = None

while True:
    event, values = window.read(timeout=100)  # check for events every 100 milliseconds
    if event == sg.WIN_CLOSED:
        break

    results = []
    if event == 'Ping':
        # stop the thread if it's currently running
        if thread is not None and thread.is_alive():
            stop_event.set()
            thread.join()
            stop_event.clear()
            thread = None
            # clear the table's values
            window['table'].update(values=[])
            # reset the results list
            results = []

        # start the thread if the continuous_ping checkbox is checked
        if values['continuous_ping']:
            hosts = values['hosts'].split(',')
            current_hosts = hosts
            # initialize current_values with the values in the 'hosts' list
            current_values = [[host, None] for host in hosts]
            window['table'].update(values=current_values)

            # start the thread
            try:
                thread = threading.Thread(target=ping_thread, args=(hosts, results_queue, stop_event))
                thread.start()
            except Exception as e:
                print("Error starting thread:", e)
        # otherwise, just update the table with the results of a single ping
        else:
            hosts = values['hosts'].split(',')
            current_hosts = hosts
            results = [[host, ping(host)] for host in hosts]
            window['table'].update(values=results)

    # update the results list with the new values in the results_queue
    if results_queue.qsize() > 0:
        results = results_queue.get()
        # update the status of existing entries
        for i, (host, status) in enumerate(results):
            for j, (current_host, current_status) in enumerate(results):
                if host == current_host:
                    results[j][1] = status
                    break
            # add new entries
            else:
                results.append([host, status])
        # update the table with the modified values
        window['table'].update(values=results)

    if event == 'New Outlook Email':
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.Display()

    if results_queue.qsize() > 0:
        results = results_queue.get()
        # get the current values in the table
        current_values = window['table'].get()
        # update the status of existing entries
        for i, (host, status) in enumerate(results):
            for j, (current_host, current_status) in enumerate(current_values):
                if host == current_host:
                    current_values[j][1] = status
                    break
            # add new entries
            else:
                current_values.append([host, status])
        # update the table with the modified values
        window['table'].update(values=current_values)

window.close()
