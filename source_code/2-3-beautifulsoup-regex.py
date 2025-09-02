from bs4 import BeautifulSoup
from urllib.request import urlopen
import re

# if has Chinese, apply decode()
<<<<<<< HEAD
html = urlopen(
    "https://mofanpy.com/static/scraping/table.html").read().decode('utf-8')
=======
html = urlopen("https://mofanpy.com/static/scraping/table.html").read().decode('utf-8')
>>>>>>> 501e9862249d3bb00ae77026ff1b4eeb4b4a48fb

soup = BeautifulSoup(html, features='lxml')

img_links = soup.find_all("img", {"src": re.compile('.*?\.jpg')})
for link in img_links:
    print(link['src'])

print('\n')

course_links = soup.find_all('a', {'href': re.compile('https://morvan.*')})
for link in course_links:
    print(link['href'])
