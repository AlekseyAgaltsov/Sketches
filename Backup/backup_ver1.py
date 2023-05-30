import os, time, zipfile

sources = ['C:\\Users\\AgaltsovAA\\Downloads']
target_dir = 'C:\\Backup'

target = target_dir + os.sep + time.strftime('%Y%m%d%H%M%S') + '.zip'
files = []

for source in sources:
    for i in os.walk(source):
        for file in i[2]:
            files.append(i[0] + '\\' + file)

try:
    with zipfile.ZipFile(target, 'w') as z:
        for file in files:
            z.write(file)
except:
    print('Создание резервной копии НЕ УДАЛОСЬ')
else:
    print('Резервная копия успешно создана в', target)
