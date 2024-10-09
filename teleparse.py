from  telethon.sync import TelegramClient #класс, позволяющий нам подключаться к клиенту мессенджера и работать с ним
 
import csv #библиотека для работы с файлами в формате CSV.
import datetime
import re

 
from telethon.tl.functions.messages import GetDialogsRequest #функция, позволяющая работать с сообщениями в чате;
from telethon.tl.types import InputPeerEmpty #конструктор для работы с InputPeer, который передаётся в качестве аргумента в GetDialogsRequest;
from telethon.tl.functions.messages import GetHistoryRequest # метод, позволяющий получить сообщения пользователей из чата и работать с ним;
from telethon.tl.types import PeerChannel #специальный тип, определяющий объекты типа «канал/чат», с помощью которого можно обратиться к нужному каналу для парсинга сообщений.

today = datetime.date.today() #Визначаю сьогоднішню дату

print(today)

api_id = 21590342
api_hash = '850984c02aaef0a36a5600aec4329456'
phone = '+380978475404'
 
client = TelegramClient(phone, api_id, api_hash)

client.start()#запускаєм клієнт


chats = []
last_date = None
size_chats = 200
groups=[]

result = client(GetDialogsRequest(
            offset_date=last_date,
            offset_id=0,
            offset_peer=InputPeerEmpty(),
            limit=size_chats,
            hash = 0
        ))
chats.extend(result.chats)


for chat in chats:
   try:
       if chat.title== 'Надходження DIP/Lak Гуменне':
           groups.append(chat)
   except:
       continue


print('Выберите номер группы из перечня:')
i=0
for g in groups:
   print(str(i) + '- ' + g.title)
   i+=1

g_index = input("Введите нужную цифру: ")
target_group=groups[int(g_index)]

print('Узнаём пользователей...')
all_participants = []
all_participants = client.get_participants(target_group)
 
print('Сохраняем данные в файл...')
with open("members.csv","w",encoding='UTF-8') as f:
   writer = csv.writer(f,delimiter=",",lineterminator="\n")
   writer.writerow(['username','name','group'])
   for user in all_participants:
       if user.username:
           username= user.username
       else:
           username= ""
       if user.first_name:
           first_name= user.first_name
       else:
           first_name= ""
       if user.last_name:
           last_name= user.last_name
       else:
           last_name= ""
       name= (first_name + ' ' + last_name).strip()
       writer.writerow([username,name,target_group.title])     
print('Парсинг участников группы успешно выполнен.')



reverse = True
all_messages = [] #для хранения спарсенных сообщений
offset_id = 0 #К ней будет обращаться метод GetHistoryRequest для того, чтобы понять, с какого сообщения начать парсинг
limit = 100 # лимит на парсинг сообщений
total_messages = 0 #cчётчик спарсенных сообщений
total_count_limit = 0#позволят нам задать ограничение на общее количество полученных сообщений.




while True: #Код в цикле будет выполняться до тех пор, пока в чате остаются сообщения, которые мы ещё не спарсили, или пока не будет достигнут установленный лимит по числу собранных сообщений
   history = client(GetHistoryRequest(
       peer=target_group,
       offset_id=offset_id,
       offset_date=None,
       add_offset=0,
       limit=limit,
       max_id=0,
       min_id=0,
       hash=0
   ))
   if not history.messages:
       break
   messages = history.messages
   for message in messages:
       all_messages.append(message.date)
       all_messages.append(message.message)
   offset_id = messages[len(messages) - 1].id
   if total_count_limit != 0 and total_messages >= total_count_limit:
       break


print(type(all_messages))      
pattern = r'(?<!\d)(?:0?[1-9]|[12][0-9]|3[01])-(?:0?[1-9]|1[0-2])-(?:19[0-9][0-9]|20[01][0-9])(?!\d)'
type_of_date = type(all_messages[0])
print(type(all_messages[0]))
print(all_messages[0])

for i in range(len(all_messages)-1):
   if type(all_messages[i]) == type_of_date:
      all_messages[i] +=  datetime.timedelta(hours=3)

print(all_messages[0])

print("Сохраняем данные в файл...") #Cообщение для пользователя о том, что начался парсинг сообщений.
 
with open("chats.csv", "w", encoding="UTF-8") as f:
   writer = csv.writer(f, delimiter=",", lineterminator="\n")
   for message in all_messages:
       writer.writerow([message])     
print("Парсинг сообщений группы успешно выполнен.") #Сообщение об удачном парсинге чата.


