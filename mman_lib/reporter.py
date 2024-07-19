import requests


def get_default_data():
    myDict = {
        i.split(" ")[0] : i.split(" ")[1] 
        for i in open("mman_lib/default.txt" , 'r' , encoding='utf-8').read().split("\n")
    }
    return myDict


def get_curent_time():
    import datetime

    current_time = datetime.datetime.now()
    formatted_time = current_time.strftime("%d.%m.%Y %H:%M:%S")
    return formatted_time

def send_report(text=None, process=None, responsible=None):
    txt2dict_data = get_default_data()
    URL = txt2dict_data['LINK']
    if not text: 
        text = txt2dict_data['text']
    if not responsible: 
        responsible = txt2dict_data['responsible']
    if not process: 
        process = txt2dict_data['process']

    response = requests.post(f"{URL}?time={get_curent_time()}&process={process}&responsible={responsible}&text={text}")
    return response
