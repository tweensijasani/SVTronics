import time
time.clock = time.time
import glob
# from chatterbot import ChatBot
# from chatterbot.trainers import ChatterBotCorpusTrainer
# from chatterbot.trainers import ListTrainer
#
#
# chatbot = ChatBot('MyChatBot')
# trainer = ChatterBotCorpusTrainer(chatbot)
# trainer.train("chatterbot.corpus.english")
#
# trainer = ListTrainer(chatbot)
# trainer.train([
#                 "How are you?",
#                 "I am good.",
#                 "That is good to hear.",
#                 "Thank you",
#                 "You're welcome."
# ])
#
# response = chatbot.get_response("Hello, how are you?")
# print(response)
from chatterbot import ChatBot
from chatterbot.trainers import ChatterBotCorpusTrainer
from chatterbot.trainers import ListTrainer

chatbot = ChatBot('MyChatBot',
                  storage_adapter='chatterbot.storage.SQLStorageAdapter',
                  database_uri='sqlite:///database.sqlite3'
                  )
trainer = ChatterBotCorpusTrainer(chatbot)
trainer.train("chatterbot.corpus.english")
trainer = ListTrainer(chatbot)

directory = "Train_data/*"
for files in glob.iglob(directory):
    file = open(files, "r")
    training_data = file.read().splitlines()
    trainer.train(training_data)

# conversation = [
#     "Hello",
#     "Hi, there!",
#     "How are you doing?",
#     "I'm doing great.",
#     "That is good to hear",
#     "how can I help you?",
#     "Thank you.",
#     "You're welcome."
# ]
# trainer.train(conversation)

exit_conditions = (":q", "quit", "exit")
while True:
    query = input("> ")
    if query in exit_conditions:
        break
    else:
        print(f"{chatbot.get_response(query)}")
