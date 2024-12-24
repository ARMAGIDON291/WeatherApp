import functions
import sys


print('*' * 70)
print("""* Привіт! Я допоможу дізнатись погоду в любому місті світу.
* Просто введи запрос в форматі city[, country_code]
* Якщо необхідно вийди з програми, тоді просто нажми Enter""")
print('*' * 70)


while True:
    q = input("Введи назву міста: ")
    if not q:
        sys.exit("До нових зустічів!")
    else:
        weather = functions.get_weather(q)
        weather = functions.print_weather(weather)
        functions.save_excel(weather)
