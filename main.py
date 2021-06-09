from openpyxl import Workbook, load_workbook
import os
from tempfile import NamedTemporaryFile
from wxpy import *
from wechaty import Wechaty
from moviepy.editor import *
import asyncio
os.environ['WECHATY_PUPPET_HOSTIE_TOKEN'] = 'puppet_paimon_174b71642d7d65d20475f3213b854081'
os.environ['WECHATY_PUPPET'] = 'wechaty-puppet-service'


def main():
    templateDir = os.path.dirname(os.path.abspath(__file__))

    wb = Workbook()
    # sheet = wb[wb.sheetnames[0]]
    # sheet['A1'] = 'test'
    shDetail = wb.create_sheet()
    shDetail.title = "new"
    shDetail['A1'] = 'test'
    wb.save(os.path.join(templateDir, "tmp.xlsx"))

    # wb = load_workbook(os.path.join(templateDir, "tmp.xlsx"))
    # sheet = wb[wb.sheetnames[0]]
    # sheet['A1'] = 'Hello world!'
    # shDetail = wb.create_sheet()
    # shDetail.title = "new"
    # wb.save(os.path.join(templateDir, "tmp.xlsx"))
    # xlsxMd5Str, fileStream = "", bytes([])
    with NamedTemporaryFile() as tmp:
        tmp.close()
        print(tmp.name)
        wb.save(tmp.name)
        # tmp.seek(0)
        # fileStream = tmp.read()
        # xlsxMd5Str = md5bytes(fileStream)

    # return xlsxMd5Str, len(fileStream), fileStream


# def controlPhone():
class MyBot(Wechaty):
    async def on_message(self, msg: Message):
        talker = msg.talker()
        await talker.ready()
        talker.say('测试')


async def main():
    bot = MyBot()
    await bot.start()


async def ChatTest():
    bot = Wechaty()
    bot.on('scan', lambda status, qrcode, data: print(
        'Scan QR Code to login: {}\nhttps://wechaty.js.org/qrcode/{}'.format(status, qrcode)))
    bot.on('login', lambda user: print('User {} logged in'.format(user)))
    bot.on('message', lambda message: print('Message: {}'.format(message)))
    await bot.start()


def receiveMusic(musicName, filePath):
    musicName = musicName.split('.mp4')[0]
    print(musicName)
    # video = VideoFileClip("../src/src/"+musicName+'.mp4')
    video = VideoFileClip(filePath)
    audio = video.audio
    audio.write_audiofile("../src/dir/wav/"+musicName+'.wav')
    audio.write_audiofile("../src/dir/mp3/"+musicName+'.mp3')


if __name__ == '__main__':
    # main()
    # controlPhone()
    # bot = Bot()
    # my_friend = bot.friends().search('郑佳煜', sex=MALE, city="广州")[0]
    # my_friend.send('测试')
    # asyncio.run(ChatTest())
    # receiveMusic()
    path = "../src/src"
    for root, dirs, files in os.walk(path):
        print("files:", files)
        for file in files:
            print(os.path.join(root, file))
            filePath = os.path.join(root, file)
            receiveMusic(file, filePath)
