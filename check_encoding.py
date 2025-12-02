import sqlite3

conn = sqlite3.connect('app.db')
cursor = conn.cursor()

# Проверяем точное содержимое адреса с ID 13
cursor.execute('SELECT address FROM records WHERE id = 13')
result = cursor.fetchone()

if result:
    address = result[0]
    print(f'Адрес: "{address}"')
    print(f'Байты: {address.encode("utf-8")}')
    print(f'LOWER: "{address.lower()}"')
    print(f'Содержит "косарева": {"косарева" in address.lower()}')
    print(f'Содержит "Косарева": {"Косарева" in address}')
    
    # Проверяем по частям
    parts = address.lower().split()
    print(f'Части: {parts}')
    for part in parts:
        if "косарева" in part:
            print(f'  Найдено в части: "{part}"')

conn.close()
