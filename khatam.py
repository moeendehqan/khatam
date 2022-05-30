import requests


# token_rev = '5134840232:AAH22oCUif1ooEt-bhxH1ZwuzkMm6FeZzIw'
# url = 'https://api.telegram.org/bot{}/'.format(token_rev)
# cmd = 'sendDocument'
# chatid = '62066305'
# doc = open('Khatam.pdf', 'rb')
# req = url+cmd+f'?chat_id={chatid}'+f'&document={doc}'

# with open("Khatam.pdf'", "rb") as file:
#     context.bot.send_document(chat_id=1234, document=file,  
#           filename='this_is_for_tg_name.txt')
# up = requests.post(req).json()


bot_token = '5134840232:AAH22oCUif1ooEt-bhxH1ZwuzkMm6FeZzIw'
bot_chatID = '62066305'

file = open('Khatam.pdf', 'rb')
send_document = 'https://api.telegram.org/bot' + bot_token + '/sendDocument?chat_id=' + bot_chatID +'&document='+str(file.read())

r = requests.post(send_document)
print(r.url)

r = r.json()

