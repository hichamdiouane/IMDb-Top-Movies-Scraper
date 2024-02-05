import requests, openpyxl
from bs4 import BeautifulSoup

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IDMB Rating'])

try:
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
    source = requests.get('https://www.imdb.com/chart/top/?ref_=nv_mv_250', headers=headers)
    source.raise_for_status()
    
    soup = BeautifulSoup(source.text, 'html.parser')
    movies = soup.find('ul', class_ = 'ipc-metadata-list ipc-metadata-list--dividers-between sc-71ed9118-0 kxsUNk compact-list-view ipc-metadata-list--base').find_all('li')
    
    for movie in movies:
        rank = movie.h3.text.split('.')[0]
        name = movie.h3.text.split('.')[1]  
        year = movie.find('span', class_ = 'sc-be6f1408-8 fcCUPU cli-title-metadata-item').text
        rating = movie.find('span', 'ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating').text
        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])

except Exception as e:
    print(e)

excel.save('IMDB Movies Rating.xlsx')
