import requests, openpyxl
from bs4 import BeautifulSoup

excel = openpyxl.Workbook()
movie_sheet = excel.active
movie_sheet.title = 'Top 250 Movies'
movie_sheet.append(['Movie Rank', 'Movie Name', 'Release Year', 'IMDb Rating'])

tv_sheet = excel.create_sheet('Top 250 TV Shows')
tv_sheet.append(['Show Rank', 'Show Name', 'Release Year', 'IMDb Rating'])

try:
    movie_url = "https://www.imdb.com/chart/top/?ref_=nv_mv_250"
    movie_result = requests.get(movie_url).text
    movie_doc = BeautifulSoup(movie_result, "html.parser")

    movies = movie_doc.find('tbody', class_ = "lister-list").find_all('tr')

    tv_url = "https://www.imdb.com/chart/toptv/?ref_=nv_tvv_250"
    tv_result = requests.get(tv_url).text
    tv_doc = BeautifulSoup(tv_result, "html.parser")

    tvshows = tv_doc.find('tbody', class_ = "lister-list").find_all('tr')

    for movie in movies:
        
        movie_name = movie.find('td', class_ = "titleColumn").a.text
        
        movie_rank = movie.find('td', class_ = "titleColumn").get_text(strip = True).split('.')[0]
        
        movie_year = movie.find('td', class_ = "titleColumn").span.text.strip("()")

        movie_rating = movie.find('td', class_ = "ratingColumn imdbRating").strong.text
        
        movie_sheet.append([movie_rank, movie_name, movie_year, movie_rating])

    for show in tvshows:

        tv_name = show.find('td', class_ = "titleColumn").a.text
        
        tv_rank = show.find('td', class_ = "titleColumn").get_text(strip = True).split('.')[0]
        
        tv_year = show.find('td', class_ = "titleColumn").span.text.strip("()")

        tv_rating = show.find('td', class_ = "ratingColumn imdbRating").strong.text
        
        tv_sheet.append([tv_rank, tv_name, tv_year, tv_rating])

except Exception as e:
    print(e)


excel.save('IMDB Movie Ratings.xlsx')
