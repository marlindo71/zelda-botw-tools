import openpyxl
import json

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Master"
ws1 = wb.create_sheet("Quests", 0) # insert at the end (default)

row = ['ID_UInt32','ID_Int32','Offset','Type','Name','Value']
ws.append(row)

row = ['Quest Name', 'Quest ID', 'Value']
ws1.append(row)

with open('gamedata.json') as json_file:
    data = json.load(json_file)

with open('quests.json') as json_file:
    quests = json.load(json_file)

with open('game_data.sav', 'rb') as f:
    f.seek(12)
    while True:
        c_size = 4
        rf_chunk = f.read(c_size)
        if len(rf_chunk) <= 0:
            break
        id_int = int.from_bytes(rf_chunk, byteorder='big',signed=True)
        
        if id_int == -1:
            break
        
        id_uint = int.from_bytes(rf_chunk, byteorder='big')
        offset = f.tell() - c_size
        gamedata = data.get(str(id_int))
        tipo = gamedata[0]
        nome = gamedata[1]
        if tipo == 'bool':
            rf_chunk = f.read(c_size)
            value = int.from_bytes(rf_chunk, byteorder='big')
        else:
            value = ''
            f.seek(c_size, 1)
        
        quest = quests.get(nome)
        if quest:
            row = [quest['Nome'],nome,value]
            ws1.append(row)

        row = [id_uint,id_int,offset,tipo,nome,value]
        ws.append(row)

wb.save('gamedata.xlsx')