abusing доступа к содержимому ящиков эл.почты.

Проблематика: пользователи могут назначит доступ к содержимому эл.ящика без указания кому именно разрешен доступ

ge.py производит проверку доступа вашей учетной записи к содержимому ящиков из списка gal.txt

download.py загружает содержимое каталога "входящие" email указанных в файле emails.txt в каталог out в формате eml.

установка зависимостей: python3 -m pip install exchangelib
стабильно работает с Python2.7

использование: python ge.py -i gal.txt -d domain.loc -u user -p Qwer -s ex.server.com

